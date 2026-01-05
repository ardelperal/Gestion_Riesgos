




Option Compare Database
Option Explicit
Private s_contadorPasos As Long
Function getColCercanasACaducarResultante( _
                                        p_ColCercanasACaducar As Scripting.Dictionary, _
                                        p_ColPreparadasParaPublicar As Scripting.Dictionary, _
                                        Optional ByRef p_Error As String) As Scripting.Dictionary
    Dim m_Id As Variant
    Dim m_Edicion As Edicion
    
    On Error GoTo errores
    If p_ColCercanasACaducar Is Nothing Then
        Exit Function
    End If
    If p_ColPreparadasParaPublicar Is Nothing Then
        Set getColCercanasACaducarResultante = p_ColCercanasACaducar
        Exit Function
    End If
    For Each m_Id In p_ColCercanasACaducar
        Set m_Edicion = p_ColCercanasACaducar(m_Id)
        If Not p_ColPreparadasParaPublicar.Exists(CStr(m_Id)) Then
            If getColCercanasACaducarResultante Is Nothing Then
                Set getColCercanasACaducarResultante = New Scripting.Dictionary
                getColCercanasACaducarResultante.CompareMode = TextCompare
            End If
            If Not getColCercanasACaducarResultante.Exists(CStr(m_Id)) Then
                getColCercanasACaducarResultante.Add m_Id, m_Edicion
            End If
        End If
        

        Set m_Edicion = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getColCercanasACaducarResultante ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Function DameUntxtYHtml1(Optional ByRef p_Error As String) As String
    Dim i As Integer
    Dim m_URLTXT As String
    Dim m_URLHTML As String
    Dim m_NombreHTML As String
    Dim m_Nombretxt As String
    Dim m_URLDirLocal As String
    On Error GoTo errores
    
    m_URLDirLocal = m_ObjEntorno.URLDirectorioLocal
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    BorraHTMLs p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    For i = 1 To 50
        m_Nombretxt = "HTML" & i & ".txt"
        m_NombreHTML = "HTML" & i & ".html"
        m_URLTXT = m_URLDirLocal & m_Nombretxt
        m_URLHTML = m_URLDirLocal & m_NombreHTML
        If Not FSO.FileExists(m_URLTXT) And Not FSO.FileExists(m_URLHTML) Then
            DameUntxtYHtml1 = m_URLHTML
            Exit Function
        End If
        
    Next
    p_Error = "No se ha podido obtener ningún html"
    Err.Raise 1000
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameUntxtYHtml1 ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function GenerarArchivoConHTML( _
                                    Optional p_Mensaje As String, _
                                    Optional ByRef p_Error As String) As String
    
    
    
    
    Dim stream As ADODB.stream
    Dim m_URL As String
    Dim m_Hwd As Long
    On Error GoTo errores
    m_URL = DameUntxtYHtml1(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_URL = "" Then
        p_Error = "No se ha podido obtener el HTML"
        Err.Raise 1000
    End If
    
    
    
    Set stream = New ADODB.stream
    With stream
        .Type = 2 ' 2 indica texto
        .Charset = "UTF-8"
        .Open
        .WriteText p_Mensaje
        
        .SaveToFile m_URL, 2 ' 2 para sobrescribir si existe
        .Close
    End With
    Set stream = Nothing
    
    
    GenerarArchivoConHTML = m_URL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarArchivoConHTML ha devuelto el error: " & vbNewLine & Err.Description
    End If
   
End Function
Public Function NombreProximoAnexoSinExtension( _
                                                p_Obj As Object, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    Dim m_ObjAnexo As Anexo
    Dim m_IDAnexo As Variant
    Dim intOrdinal As Integer
    Dim intOrdinalMaximo As Integer
    Dim m_NombreArchivo As String
    Dim m_NombreSinExtension As String
    
    
    On Error GoTo errores
    'NombreProximoAnexoSinExtension=PPPEEExx
    'PPP=IDPROYECTO
    'EEE=IDEDICION
    'XX ORDINAL
    Set p_Obj.ColAnexos = Nothing
    If Not p_Obj.ColAnexos Is Nothing Then
        p_Error = p_Obj.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        For Each m_IDAnexo In p_Obj.ColAnexos
            If Nz(m_IDAnexo, "") <> "" Then
                Set m_ObjAnexo = p_Obj.ColAnexos(m_IDAnexo)
                m_NombreArchivo = m_ObjAnexo.NombreArchivo
                m_NombreSinExtension = FSO.GetBaseName(m_NombreArchivo)
                If IsNumeric(Right(m_NombreSinExtension, 2)) Then
                    intOrdinal = CInt(Right(m_NombreSinExtension, 2))
                    If intOrdinal > intOrdinalMaximo Then
                        intOrdinalMaximo = intOrdinal
                    End If
                End If
                Set m_ObjAnexo = Nothing
            End If
            
        Next
    End If
    If TypeOf p_Obj Is Proyecto Then
        NombreProximoAnexoSinExtension = Format(p_Obj.IDProyecto, "000") & "_" & _
                                        Format(intOrdinalMaximo + 1, "00")
    ElseIf TypeOf p_Obj Is Edicion Then
        NombreProximoAnexoSinExtension = Format(p_Obj.Proyecto.IDProyecto, "000") & "E" & _
                                    Format(p_Obj.IDEdicion, "000") & "_" & _
                                    Format(intOrdinalMaximo + 1, "00")
    ElseIf TypeOf p_Obj Is riesgo Then
        NombreProximoAnexoSinExtension = p_Obj.CodigoUnico & "_" & Format(intOrdinalMaximo + 1, "00")
    Else
        p_Error = "El objeto no está reconocido"
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método NombreProximoAnexoSinExtension ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function EstablecerComboUsuariosCalidad( _
                                                ByRef cmb As ComboBox, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    Dim m_objColUsuarios As Scripting.Dictionary
    Dim m_UsuarioRed  As Variant
    Dim m_ObjUsuario As Usuario
    
    On Error GoTo errores
    cmb.RowSource = ""
    Set m_objColUsuarios = m_ObjEntorno.ColUsuariosCalidad
    If m_objColUsuarios Is Nothing Then
        Exit Function
    End If
    
    For Each m_UsuarioRed In m_objColUsuarios
        Set m_ObjUsuario = m_objColUsuarios(m_UsuarioRed)
        cmb.AddItem m_ObjUsuario.Nombre
        Set m_ObjUsuario = Nothing
siguiente:
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerComboUsuariosCalidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function RellenaComboCodProyectos( _
                                        ByRef cmb As ComboBox, _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                    
    
    Dim m_Proyecto As Proyecto
    Dim m_Id As Variant
    
    On Error GoTo errores
    
    cmb.RowSource = ""
    
    
    
    If Not m_ObjEntorno.ColProyectosTotales Is Nothing Then
        For Each m_Id In m_ObjEntorno.ColProyectosTotales
           Set m_Proyecto = m_ObjEntorno.ColProyectosTotales(m_Id)
            cmb.AddItem m_Proyecto.Proyecto
            Set m_Proyecto = Nothing
        Next
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenaComboCodProyectos ha devuelto el error: " & Err.Description
    End If
End Function
Public Function UsuarioEsAdministrador( _
                                        p_UsuarioRed As String, _
                                        Optional ByRef p_Error As String _
                                        ) As EnumSiNo
    On Error GoTo errores
    
    If p_UsuarioRed = "" Then
        p_Error = "No se ha indicado el usuario"
        Err.Raise 1000
    End If
    If m_ObjEntorno.ColUsuariosAdministradores Is Nothing Then
        p_Error = m_ObjEntorno.Error
        Err.Raise 1000
        UsuarioEsAdministrador = EnumSiNo.No
        Exit Function
    End If
    If m_ObjEntorno.ColUsuariosAdministradores.Exists(p_UsuarioRed) Then
        UsuarioEsAdministrador = EnumSiNo.Sí
    Else
        UsuarioEsAdministrador = EnumSiNo.No
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioEsAdministrador ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function UsuarioEsDeCalidad( _
                                    p_UsuarioRed As String, _
                                    Optional ByRef p_Error As String _
                                    ) As EnumSiNo
    On Error GoTo errores
    
    If p_UsuarioRed = "" Then
        p_Error = "No se ha indicado el usuario"
        Err.Raise 1000
    End If
    With m_ObjEntorno
        If .ColUsuariosCalidad Is Nothing Then
            p_Error = .Error
            Err.Raise 1000
            UsuarioEsDeCalidad = EnumSiNo.No
            Exit Function
        End If
        If .ColUsuariosCalidad.Exists(p_UsuarioRed) Then
            UsuarioEsDeCalidad = EnumSiNo.Sí
        Else
            UsuarioEsDeCalidad = EnumSiNo.No
        End If
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioEsDeCalidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function UsuarioAutorizado( _
                                    p_Objeto As Object, _
                                    Optional p_Usuario As Usuario, _
                                    Optional ByRef p_Error As String _
                                    ) As EnumSiNo
                                    
    Dim m_ObjProyecto As Proyecto
    On Error GoTo errores
    
    If p_Usuario Is Nothing Then
        Set p_Usuario = m_ObjUsuarioConectado
        
    End If
    If UsuarioEsAdministrador(p_Usuario.UsuarioRed) = EnumSiNo.Sí Or UsuarioEsDeCalidad(p_Usuario.UsuarioRed) = EnumSiNo.Sí Then
        UsuarioAutorizado = EnumSiNo.Sí
        Exit Function
    End If
    Set m_ObjProyecto = getProyectoPorObjeto(p_Objeto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjProyecto Is Nothing Then
        p_Error = "No se ha podido obtener el Proyecto del que depende"
        Err.Raise 1000
    End If
    UsuarioAutorizado = m_ObjProyecto.UsuarioAutorizado(p_Usuario, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 0 Then
        p_Error = "El método UsuarioAutorizado ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Function FicheroAbierto( _
                        m_URL As String, _
                        Optional ByRef p_Error As String _
                        ) As Boolean
    
    Dim intfilenum As Integer
    
    On Error GoTo errores
    
    intfilenum = FreeFile()
    Open m_URL For Binary Access Read Write Lock Read Write As #intfilenum
    Close #intfilenum
    If Err.Number <> 0 Then
        FicheroAbierto = True
    Else
        FicheroAbierto = False
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FicheroAbierto ha producido el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function CerrarFormulariosAbiertos(Optional ByRef p_Error As String) As String
    Dim obj As AccessObject, dbs As Object
    
    Set dbs = Application.CurrentProject
    ' Search for open AccessObject objects in AllForms collection.
    For Each obj In dbs.AllForms
        If obj.IsLoaded = True Then
            ' Print name of obj.
            
            DoCmd.Close acForm, obj.Name, acSaveNo
        End If
    Next obj
End Function
Public Function FormularioAbierto( _
                                    p_NombreFormulario As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Boolean
    
    Dim m_Estado As String
    
    On Error GoTo errores
    
    m_Estado = SysCmd(acSysCmdGetObjectState, acForm, p_NombreFormulario)
    If m_Estado = "0" Then
        FormularioAbierto = False
    Else
        FormularioAbierto = True
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FormularioAbierto ha producido el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function AbrirEnLocal( _
                                m_URLRemoto As String, _
                                lngHwnd As Long, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    
    Dim URLLocal As String
    On Error GoTo errores
    
    If Not FSO.FileExists(m_URLRemoto) Then
        p_Error = "No es accesible la ruta del archivo que se pretende abrir" & vbNewLine & m_URLRemoto
        Err.Raise 1000
    End If
    URLLocal = m_ObjEntorno.URLDirectorioLocal & FSO.GetFile(m_URLRemoto).Name
    If FSO.FileExists(URLLocal) Then
        If FicheroAbierto(URLLocal) Then
           p_Error = "Tiene el archivo abierto"
            Err.Raise 1000
        End If
    End If
    FSO.CopyFile m_URLRemoto, URLLocal, True
    Ejecutar lngHwnd, "open", URLLocal, "", "", 1
    AbrirEnLocal = URLLocal
    Exit Function
errores:
    If Err.Number = 70 Then
        p_Error = "El archivo parece estar abierto"
    Else
        If Err.Number <> 1000 Then
            p_Error = "El método AbrirEnLocal ha devuelto el error: " & vbNewLine & Err.Description
        End If
    End If
    
End Function
Public Function EnOficina(Optional ByRef p_Error As String) As EnumSiNo
    
    Dim strIPS As String
    On Error GoTo errores
    strIPS = GetIPAddresses
    If InStr(1, strIPS, SubRedOficina) = 0 Then
        EnOficina = EnumSiNo.No
    Else
        EnOficina = EnumSiNo.Sí
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnOficina ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Private Function getUsuarioMaquina( _
                            Optional ByRef p_Error As String _
                            ) As String
    Dim objNetwork As Object
    On Error GoTo errores
    Set objNetwork = CreateObject("Wscript.Network")
    With objNetwork
        getUsuarioMaquina = .UserName & "|" & .computername
    End With
   
    Set objNetwork = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuarioMaquina ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getMaquina( _
                            Optional ByRef p_Error As String _
                            ) As String
    Dim flag As String
    Dim dato As Variant
    On Error GoTo errores
    flag = getUsuarioMaquina
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        getMaquina = dato(1)
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getMaquina ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getUsuariodeMaquina( _
                                Optional ByRef p_Error As String _
                                ) As String
    Dim flag As String
    Dim dato As Variant
    On Error GoTo errores
    flag = getUsuarioMaquina
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        getUsuariodeMaquina = dato(0)
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariodeMaquina ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Public Function CorreoAlAdministrador( _
                                        m_MensajeError As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                        
   
    Dim m_mensaje As String
    Dim m_Asunto As String
    Dim m_Nombre As String
    Dim m_NombreFormulario As String
    Dim m_TextoEnOficina As String
    Dim m_Version As String
    
    Dim m_ObjCorreo As CORREO
    On Error GoTo errores
    
    If m_MensajeError = "" Then
        p_Error = "No hay mensaje que enviar"
        Err.Raise 1000
    End If
    If Not m_ObjEntorno Is Nothing Then
        m_Version = "Versión: " & m_ObjEntorno.VersionAplicacion
    Else
        m_Version = "Versión:Desconocida "
    End If
    If m_EnOficina = Empty Then
        m_EnOficina = EnOficina(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If m_EnOficina = Empty Then
        m_TextoEnOficina = "En Oficina Desconocido"
    Else
        If m_EnOficina = EnumSiNo.Sí Then
            m_TextoEnOficina = "En Oficina"
        Else
            m_TextoEnOficina = "Fuera de Oficina"
        End If
    End If
    On Error Resume Next
    m_NombreFormulario = Application.Screen.ActiveForm.Name
    If Err.Number <> 0 Then
        Err.Clear
        m_NombreFormulario = "Desconocido"
    End If
    On Error GoTo errores
    m_Nombre = getNombreUsuarioConectado()
    
    m_Asunto = "Error en RIESGOS m_Version" & m_Version & ") " & m_Nombre & " " & m_TextoEnOficina
    m_mensaje = "FORMULARIO del ERROR: " & m_NombreFormulario & vbNewLine
    m_mensaje = m_mensaje & "<BR> </BR>" & vbNewLine
    On Error Resume Next
    m_mensaje = m_mensaje & "NOMBRE EQUIPO: " & VBA.Environ("COMPUTERNAME")
    m_mensaje = m_mensaje & "<BR> </BR>" & vbNewLine
    On Error GoTo errores
    
    m_mensaje = m_mensaje & "DETALLE: " & m_MensajeError
    
    Set m_ObjCorreo = New CORREO
    With m_ObjCorreo
        .Asunto = m_Asunto
        .Cuerpo = m_mensaje
        .Destinatarios = "ardelperal@gmail.com;andres.romandelperal@telefonica.com"
        .FechaGrabacion = Now()
        .EnviarCorreo p_Error
    End With
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreoAlAdministrador ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Sub AjustarTamaño(frmFormulario As Form)

    Dim i As Integer

    On Error GoTo AjustarTamaño_TratamientoErrores

    ' ajusto el ancho del formulario teniendo en cuenta si tiene o no selector de registros
    If Not frmFormulario.RecordSelectors Then
        frmFormulario.InsideWidth = frmFormulario.Width
    Else
        frmFormulario.InsideWidth = frmFormulario.Width + 250
    End If

    ' si se abre en vista formulario simple
    If frmFormulario.DefaultView = 0 Then
        'ajusto el alto incluyendo las distintas secciones, encabezado, pie, grupos...
        ' como no sé el número de secciones del formulario, me salgo al producirse un error
        frmFormulario.InsideHeight = 0
        For i = 0 To 100
            frmFormulario.InsideHeight = frmFormulario.InsideHeight + frmFormulario.Section(i).Height
        Next
    End If

AjustarTamaño_Salir:
   DoCmd.Restore
   On Error GoTo 0
   Exit Sub
   
AjustarTamaño_TratamientoErrores:
   If Not Err = 2462 Then  ' "El número de sección que introdujo no es válido."
      MsgBox "Error " & Err.Number & " en proc.: AjustarTamaño de Módulo: Módulo1 (" & Err.Description & ")"
   End If
   Resume AjustarTamaño_Salir
End Sub         ' Aju

Public Function NumeroDePalabras( _
                                    m_Texto As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Long
    
    Dim m_NumPalabras As Long
    Dim dato
    On Error GoTo errores
    
    If m_Texto <> "" Then
        m_Texto = Replace(m_Texto, vbNewLine, "")
        dato = Split(m_Texto, " ")
        
        m_NumPalabras = UBound(dato) + 1
    End If
    NumeroDePalabras = m_NumPalabras
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método NumeroDePalabras ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function Seleccionar( _
                            p_EsArchivo As Boolean, _
                            Optional p_Titulo As String, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_ObjfDialog As Object
    Dim varFile As Variant
    
    On Error GoTo errores
    
    If p_Titulo = "" Then
        p_Titulo = "Seleccione el archivo"
    End If
    If p_EsArchivo = True Then
        Set m_ObjfDialog = Application.FileDialog(msoFileDialogFilePicker)
    Else
        Set m_ObjfDialog = Application.FileDialog(msoFileDialogFolderPicker)
    End If
    With m_ObjfDialog
        .Show
        If p_EsArchivo Then
            .AllowMultiSelect = False
            .InitialFileName = m_ObjEntorno.URLUltimoArchivo
            .Title = p_Titulo
            .Filters.Clear
            .Filters.Add "All Files", "*.*"
        End If
        For Each varFile In .SelectedItems
            Seleccionar = CStr(varFile)
        Next
    End With
    If p_EsArchivo Then
        m_ObjEntorno.URLUltimoArchivo = CStr(varFile)
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Seleccionar ha producido el error : " & vbNewLine & Err.Description
    End If
End Function

Public Function getProyectoPorObjeto( _
                                        ByRef p_ObjDatos As Object, _
                                        Optional ByRef p_Error As String _
                                        ) As Proyecto
    
    On Error GoTo errores
    
    If TypeOf p_ObjDatos Is Proyecto Then
        Set getProyectoPorObjeto = p_ObjDatos
    ElseIf TypeOf p_ObjDatos Is Edicion Then
        Set getProyectoPorObjeto = p_ObjDatos.Proyecto
        p_Error = p_ObjDatos.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    ElseIf TypeOf p_ObjDatos Is riesgo Then
        If p_ObjDatos.Edicion Is Nothing Then
            p_Error = p_ObjDatos.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Function
        End If
        Set getProyectoPorObjeto = p_ObjDatos.Edicion.Proyecto
        p_Error = p_ObjDatos.Edicion.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    ElseIf TypeOf p_ObjDatos Is PM Then
        If p_ObjDatos.riesgo Is Nothing Then
            p_Error = p_ObjDatos.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Function
        End If
        If p_ObjDatos.riesgo.Edicion Is Nothing Then
            p_Error = p_ObjDatos.Edicion.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Function
        End If
        Set getProyectoPorObjeto = p_ObjDatos.riesgo.Edicion.Proyecto
        p_Error = p_ObjDatos.riesgo.Edicion.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    ElseIf TypeOf p_ObjDatos Is PC Then
        If p_ObjDatos.riesgo Is Nothing Then
            p_Error = p_ObjDatos.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Function
        End If
        If p_ObjDatos.riesgo.Edicion Is Nothing Then
            p_Error = p_ObjDatos.Edicion.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Function
        End If
        Set getProyectoPorObjeto = p_ObjDatos.riesgo.Edicion.Proyecto
        p_Error = p_ObjDatos.riesgo.Edicion.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    ElseIf TypeOf p_ObjDatos Is PMAccion Then
        If p_ObjDatos.Mitigacion Is Nothing Then
            p_Error = p_ObjDatos.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Function
        End If
        If p_ObjDatos.Mitigacion.riesgo Is Nothing Then
            p_Error = p_ObjDatos.Mitigacion.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Function
        End If
        If p_ObjDatos.Mitigacion.riesgo.Edicion Is Nothing Then
            p_Error = p_ObjDatos.Mitigacion.riesgo.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Function
        End If
        Set getProyectoPorObjeto = p_ObjDatos.Mitigacion.riesgo.Edicion.Proyecto
        p_Error = p_ObjDatos.Mitigacion.riesgo.Edicion.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    ElseIf TypeOf p_ObjDatos Is PCAccion Then
        If p_ObjDatos.Contingencia Is Nothing Then
            p_Error = p_ObjDatos.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Function
        End If
        If p_ObjDatos.Contingencia.riesgo Is Nothing Then
            p_Error = p_ObjDatos.Contingencia.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Function
        End If
        If p_ObjDatos.Contingencia.riesgo.Edicion Is Nothing Then
            p_Error = p_ObjDatos.Contingencia.riesgo.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Function
        End If
        Set getProyectoPorObjeto = p_ObjDatos.Contingencia.riesgo.Edicion.Proyecto
        p_Error = p_ObjDatos.Contingencia.riesgo.Edicion.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    ElseIf TypeOf p_ObjDatos Is Anexo Then
        If Not p_ObjDatos.Proyecto Is Nothing Then
            p_Error = p_ObjDatos.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set getProyectoPorObjeto = p_ObjDatos.Proyecto
        End If
    ElseIf TypeOf p_ObjDatos Is AnexoAntiguo Then
        If Not p_ObjDatos.Proyecto Is Nothing Then
            p_Error = p_ObjDatos.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set getProyectoPorObjeto = p_ObjDatos.Proyecto
        End If
    
    ElseIf TypeOf p_ObjDatos Is RiesgoExterno Then
        If Not p_ObjDatos.Edicion.Proyecto Is Nothing Then
            p_Error = p_ObjDatos.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set getProyectoPorObjeto = p_ObjDatos.Proyecto
        End If
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getProyectoPorObjeto ha devuelto el error: " & Err.Description
    End If
                                
End Function




Public Function MostrarHTML( _
                            p_Mensaje As ADODB.stream, _
                            Optional p_EnFormulario As EnumSiNo = EnumSiNo.Sí, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    
    Dim m_ObjHTML As html
    On Error GoTo errores
    
    Set m_ObjHTML = New html
    m_ObjHTML.MostrarHTML p_Mensaje:=p_Mensaje, p_EnFormulario:=p_EnFormulario, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjHTML = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MostrarHTML ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function ActualizarObjetosActivos( _
                                        Optional p_EstablecerlosANulos As EnumSiNo = EnumSiNo.No, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    
    On Error GoTo errores
    If Not m_ObjProyectoActivo Is Nothing Then
        If p_EstablecerlosANulos = EnumSiNo.No Then
            Set m_ObjProyectoActivo = getProyecto(m_ObjProyectoActivo.IDProyecto, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Else
            Set m_ObjProyectoActivo = Nothing
        End If
    End If
    If Not m_ObjEdicionActiva Is Nothing Then
        If p_EstablecerlosANulos = EnumSiNo.No Then
            Set m_ObjEdicionActiva = getEdicion(m_ObjEdicionActiva.IDEdicion, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Else
            Set m_ObjEdicionActiva = Nothing
        End If
    End If
    If p_EstablecerlosANulos = EnumSiNo.No Then
         If Not m_ObjRiesgoActivo Is Nothing Then
            Set m_ObjRiesgoActivo = getRiesgo(m_ObjRiesgoActivo.IDRiesgo, , , p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    Else
        Set m_ObjRiesgoActivo = Nothing
    End If
    If p_EstablecerlosANulos = EnumSiNo.No Then
        If Not m_ObjPMActivo Is Nothing Then
            Set m_ObjPMActivo = getPM(m_ObjPMActivo.IDMitigacion, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    Else
        Set m_ObjPMActivo = Nothing
    End If
    If p_EstablecerlosANulos = EnumSiNo.No Then
        If Not m_ObjPCActivo Is Nothing Then
            Set m_ObjPCActivo = getPC(m_ObjPCActivo.IDContingencia, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    Else
        Set m_ObjPCActivo = Nothing
    End If
    If p_EstablecerlosANulos = EnumSiNo.No Then
        If Not m_ObjPMAccionActiva Is Nothing Then
            Set m_ObjPMAccionActiva = getPMAccion(m_ObjPMAccionActiva.IDAccionMitigacion, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    Else
        Set m_ObjPMAccionActiva = Nothing
    End If
    If p_EstablecerlosANulos = EnumSiNo.No Then
        If Not m_ObjPCAccionActiva Is Nothing Then
            Set m_ObjPCAccionActiva = getPCAccion(m_ObjPCAccionActiva.IDAccionContingencia, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    Else
        Set m_ObjPCAccionActiva = Nothing
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizarObjetosActivos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function GenerarInformePublicacion( _
                                           p_Edicion As Edicion, _
                                           Optional ByRef p_EnWord As EnumSiNo = EnumSiNo.No, _
                                           Optional p_FechaCierre As String, _
                                           Optional p_FechaPublicacion As String, _
                                           Optional ByRef p_Error As String _
                                           ) As String

    Dim Tipo As EnumTipoInformePublicacion

    On Error GoTo errores

    If p_EnWord = EnumSiNo.Sí Then
        Tipo = EnumTipoInformePublicacion.Word
    Else
        Tipo = m_ObjEntorno.GenerarInformeTipo
    End If

    Select Case Tipo
        Case EnumTipoInformePublicacion.Word
            GenerarInformePublicacion = GenerarInformeWord(p_Edicion, EnumSiNo.No, p_FechaCierre, p_FechaPublicacion, p_Error, p_EnWord)
        Case EnumTipoInformePublicacion.html
            GenerarInformePublicacion = InformeRiesgoHTML.GenerarInformeEdicionHTML(p_Edicion, 0, p_FechaCierre, p_FechaPublicacion, p_Error)
        Case Else
            GenerarInformePublicacion = ExcelInforme.GenerarInforme(p_Edicion, EnumSiNo.Sí, p_FechaCierre, p_FechaPublicacion, p_Error)
    End Select
    If p_Error <> "" Then Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarInformePublicacion ha producido el error nº: " & Err.Number & _
                  vbNewLine & "Detalle: " & Err.Description
    End If
End Function


Public Function MostrarProyecto( _
                                    m_ObjProyecto As Proyecto, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim m_ObjHTML As html
    Dim m_mensaje As ADODB.stream
    Dim m_HTML As String
    On Error GoTo errores
    
    If m_ObjProyecto Is Nothing Then
        p_Error = "No hay datos "
        Err.Raise 1000
    End If
    m_HTML = m_ObjProyecto.HTMLProyecto
    p_Error = m_ObjProyecto.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_mensaje = New ADODB.stream
    With m_mensaje
        .Open
        .WriteText m_HTML
        MostrarHTML p_Mensaje:=m_mensaje, p_Error:=p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        .Close
    End With
    Set m_mensaje = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTML.MostrarProyecto ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function MostrarEdicionCambios( _
                                            p_ObjColCambios As Scripting.Dictionary, _
                                            p_Edicion As String, _
                                            p_EdicionAnterior As String, _
                                            p_NombreProyecto As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    Dim m_ObjHTML As html
    
    On Error GoTo errores
    
    
    Set m_ObjHTML = New html
    MostrarEdicionCambios = m_ObjHTML.MostrarEdicionCambios(p_ObjColCambios, p_Edicion, p_EdicionAnterior, p_NombreProyecto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjHTML = Nothing
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MostrarEdicionCambios ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Public Function MostrarControlCambiosVersionActual( _
                                                    Optional p_QueMostrar As EnumControlCambios = EnumControlCambios.Todo, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    Dim m_ObjHTML As html
    Dim m_Version As CCVersion
    On Error GoTo errores
    Set m_Version = Constructor.getCCVersion(, m_ObjEntorno.VersionAplicacion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
   
    Set m_ObjHTML = New html
    MostrarControlCambiosVersionActual = m_ObjHTML.MostrarControlCambiosVersion(m_Version, p_QueMostrar, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjHTML = Nothing
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MostrarControlCambiosVersion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function MostrarControlCambiosVersion( _
                                                p_Version As CCVersion, _
                                                Optional p_QueMostrar As EnumControlCambios = EnumControlCambios.Todo, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim m_ObjHTML As html
    
    On Error GoTo errores
   
   
    Set m_ObjHTML = New html
    MostrarControlCambiosVersion = m_ObjHTML.MostrarControlCambiosVersion(p_Version, p_QueMostrar, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjHTML = Nothing
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MostrarControlCambiosVersion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getHTMCabecera( _
                                    Optional p_Titulo As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim m_mensaje As String
    On Error GoTo errores
    
    If p_Titulo = "" Then
        p_Titulo = "N/A"
    End If
    m_mensaje = m_ObjEntorno.CabeceraHTML
    m_mensaje = Replace(m_mensaje, "#titulo", p_Titulo)
    getHTMCabecera = m_mensaje
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMCabecera ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getHTMLUsuariosConectados( _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim m_mensaje As String
    Dim m_CorreoUsuarioVicario As String
    Dim m_CorreoUsuarioOriginal As String
    Dim m_NombreUsuarioOriginal As String
    Dim m_NombreUsuarioVicario As String
    Dim m_UsuarioVicario As String
    
    On Error GoTo errores
    If Not m_ObjUsuarioConectadoInicialmente Is Nothing Then
        m_CorreoUsuarioOriginal = m_ObjUsuarioConectadoInicialmente.CorreoUsuario
        m_NombreUsuarioOriginal = m_ObjUsuarioConectadoInicialmente.Nombre
    End If
    If Not m_ObjUsuarioConectado Is Nothing Then
        m_CorreoUsuarioVicario = m_ObjUsuarioConectado.CorreoUsuario
        m_NombreUsuarioVicario = m_ObjUsuarioConectado.Nombre
    End If
    If m_NombreUsuarioOriginal <> "" And m_NombreUsuarioVicario <> "" Then
        If m_NombreUsuarioOriginal <> m_NombreUsuarioVicario Then
            m_mensaje = "<a href='mailto:" & m_CorreoUsuarioOriginal & "'>correo enviado por gestión de riesgos en nombre de: " & _
                m_NombreUsuarioOriginal & " entrando como : " & m_NombreUsuarioVicario & "</a>" & vbNewLine
        Else
            m_mensaje = "<a href='mailto:" & m_CorreoUsuarioOriginal & "'>correo enviado por gestión de riesgos en nombre de: " & _
                m_NombreUsuarioOriginal & "</a>" & vbNewLine
        End If
    Else
        If m_NombreUsuarioOriginal <> "" Then
            m_mensaje = "<a href='mailto:" & m_CorreoUsuarioOriginal & "'>correo enviado por gestión de riesgos en nombre de: " & _
                m_NombreUsuarioOriginal & "</a>" & vbNewLine
        Else
            m_mensaje = "<a href='mailto:" & m_CorreoUsuarioVicario & "'>correo enviado por gestión de riesgos en nombre de: " & _
                m_NombreUsuarioVicario & "</a>" & vbNewLine
        End If
    End If
    
    
   
    getHTMLUsuariosConectados = m_mensaje
    Exit Function
errores:
    getHTMLUsuariosConectados = "-1"
End Function





Public Function getHTMLTablaCambiosExplicaciones( _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    Dim m_Col As Scripting.Dictionary
    Dim m_mensaje As String
    Dim m_Apartado As String
    Dim m_Explicacion As String
    Dim m_CambioExplicacion As CambioExplicacion
    Dim m_Id As Variant
    
    
    On Error GoTo errores
    Set m_Col = m_ObjEntorno.ColCambiosExplicaciones
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    m_mensaje = m_mensaje & "<table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""ColespanArriba"" colspan='2'> Explicación de los Apartados</td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"">APARTADO</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">EXPLICACIÓN</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        For Each m_Id In m_Col
            Set m_CambioExplicacion = m_Col(m_Id)
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
                m_mensaje = m_mensaje & "<td>" & m_CambioExplicacion.Apartado & "</td>" & vbNewLine
                m_mensaje = m_mensaje & "<td>" & m_CambioExplicacion.Explicacion & "</td>" & vbNewLine
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
            Set m_CambioExplicacion = Nothing
        Next
    
    m_mensaje = m_mensaje & "</table>" & vbNewLine
    
    getHTMLTablaCambiosExplicaciones = m_mensaje
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLTablaCambiosExplicaciones ha devuelto el error: " & Err.Description
    End If
End Function



Public Function getHTMLTablaExplicacionesEstadosRiesgos( _
                                                        Optional ByRef p_Error As String _
                                                        ) As String
    
    Dim m_mensaje As String
    Dim m_Apartado As String
    Dim m_Explicacion As String
    
    
    On Error GoTo errores
    
    m_mensaje = m_mensaje & "<table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""ColespanArriba"" colspan='2'> ESTADOS DE UN RIESGO</td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
            m_Apartado = "Incompleto"
            m_Explicacion = "Primera edición de riesgos en fase de oferta grabada por Calidad, a falta de cumplimentación por parte del técnico."
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
                m_mensaje = m_mensaje & "<td><b>" & m_Apartado & "</b></td>" & vbNewLine
                m_mensaje = m_mensaje & "<td>" & m_Explicacion & "</td>" & vbNewLine
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
            m_Apartado = "Detectado"
            m_Explicacion = "Riesgo identificado, sin ningún plan definido o con planes finalizados."
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
                m_mensaje = m_mensaje & "<td><b>" & m_Apartado & "</b></td>" & vbNewLine
                m_mensaje = m_mensaje & "<td>" & m_Explicacion & "</td>" & vbNewLine
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
            m_Apartado = "Planificado"
            m_Explicacion = "Riesgo para el cual se ha definido un plan demitigación, pero sin fechas previstas de activación. "
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
                m_mensaje = m_mensaje & "<td><b>" & m_Apartado & "</b></td>" & vbNewLine
                m_mensaje = m_mensaje & "<td>" & m_Explicacion & "</td>" & vbNewLine
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
            m_Apartado = "Activo"
            m_Explicacion = "Riesgo con posibilidad de materialización, por lo que, al menos, tiene un plan de mitigación activo (con fechas planificadas) "
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
                m_mensaje = m_mensaje & "<td><b>" & m_Apartado & "</b></td>" & vbNewLine
                m_mensaje = m_mensaje & "<td>" & m_Explicacion & "</td>" & vbNewLine
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
            m_Apartado = "Materializado"
            m_Explicacion = "Riesgo realizado, acompañado de la manifestaciñón de las consecuencias asociadas al mismo, por lo que podría derivar en una no conformidad. " & vbNewLine & _
                            "El riesgo tiene registrada una fecha de materialización y no tiene un plan de contingencia finalizado con fecha de activación del plan posterior a la fecha de materailizadión del riesgo."
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
                m_mensaje = m_mensaje & "<td><b>" & m_Apartado & "</b></td>" & vbNewLine
                m_mensaje = m_mensaje & "<td>" & m_Explicacion & "</td>" & vbNewLine
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
            m_Apartado = "Aceptado"
            m_Explicacion = "Riesgo cuya posibilidad de materialización se ha aceptado ante la imposibilidad de ejecutar acciones que puedan eliminar el riesgo. Es visado por parte de Calidad. "
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
                m_mensaje = m_mensaje & "<td><b>" & m_Apartado & "</b></td>" & vbNewLine
                m_mensaje = m_mensaje & "<td>" & m_Explicacion & "</td>" & vbNewLine
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
            m_Apartado = "Aceptado pendiente de Calidad"
            m_Explicacion = "Riesgo cuya posibilidad de materialización se ha aceptado ante la imposibilidad de ejecutar acciones que puedan eliminar el riesgo. Es visado por parte de Calidad. " & vbNewLine & _
                            "Aceptación de Riesgo pendiente de visado por parte de Calidad"
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
                m_mensaje = m_mensaje & "<td><b>" & m_Apartado & "</b></td>" & vbNewLine
                m_mensaje = m_mensaje & "<td>" & m_Explicacion & "</td>" & vbNewLine
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
            m_Apartado = "Aceptado rechazado por Calidad"
            m_Explicacion = "Riesgo cuya posibilidad de materialización se ha aceptado ante la imposibilidad de ejecutar acciones que puedan eliminar el riesgo. Es visado por parte de Calidad. " & vbNewLine & _
                            "Aceptación de Riesgo pendiente de visado por parte de Calidad, por lo que debe revaluarse."
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
                m_mensaje = m_mensaje & "<td><b>" & m_Apartado & "</b></td>" & vbNewLine
                m_mensaje = m_mensaje & "<td>" & m_Explicacion & "</td>" & vbNewLine
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
            m_Apartado = "Retirado"
            m_Explicacion = "Riesgo que deja de afectar al proyecto. Debe registrase la fecha de retiro y es visado por parte de Calidad. "
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
                m_mensaje = m_mensaje & "<td><b>" & m_Apartado & "</b></td>" & vbNewLine
                m_mensaje = m_mensaje & "<td>" & m_Explicacion & "</td>" & vbNewLine
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
            m_Apartado = "Retirado pendiente de Calidad"
            m_Explicacion = "Riesgo que deja de afectar al proyecto. Fecha de retiro pendiente de visado por parte de Calidad. "
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
                m_mensaje = m_mensaje & "<td><b>" & m_Apartado & "</b></td>" & vbNewLine
                m_mensaje = m_mensaje & "<td>" & m_Explicacion & "</td>" & vbNewLine
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
            m_Apartado = "Retirado rechazado por Calidad"
            m_Explicacion = "Riesgo que deja de afectar al proyecto. Fecha de retiro rechazada por parte de Calidad. "
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
                m_mensaje = m_mensaje & "<td><b>" & m_Apartado & "</b></td>" & vbNewLine
                m_mensaje = m_mensaje & "<td>" & m_Explicacion & "</td>" & vbNewLine
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
            m_Apartado = "Cerrado"
            m_Explicacion = "Riesgo carente de la posibilidad de materialización o cambio por haber finalizado las actuaciones previstas en los distintos planes de prevención, " & _
                            "mitigación o contingencia y el impacto del riesgo ha sido subsanado o corregido. " & _
                            "Los riesgos se cierran automáticamente al publcar la última edición, con fecha de estado de la última publicación."
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
                m_mensaje = m_mensaje & "<td><b>" & m_Apartado & "</b></td>" & vbNewLine
                m_mensaje = m_mensaje & "<td>" & m_Explicacion & "</td>" & vbNewLine
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
        
    m_mensaje = m_mensaje & "</table>" & vbNewLine
    
    getHTMLTablaExplicacionesEstadosRiesgos = m_mensaje
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLTablaExplicacionesEstadosRiesgos ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getHTMLTablaExplicacionesEstadosPlanes( _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    
    Dim m_mensaje As String
    Dim m_Apartado As String
    Dim m_Explicacion As String
    
    
    On Error GoTo errores
    
    m_mensaje = m_mensaje & "<table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""ColespanArriba"" colspan='2'> ESTADOS DE UN PLAN DE MITIGACIÓN/CONTINGENCIA</td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        
        m_Apartado = "Plan incompleto"
        m_Explicacion = "No tiene registrado ninguna acción como última acción"
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td><b>" & m_Apartado & "</b></td>" & vbNewLine
            m_mensaje = m_mensaje & "<td>" & m_Explicacion & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_Apartado = "Plan definido"
        m_Explicacion = "No todas las acciones tienen definidas una fecha de fin prevista "
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td><b>" & m_Apartado & "</b></td>" & vbNewLine
            m_mensaje = m_mensaje & "<td>" & m_Explicacion & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
         m_Apartado = "Plan planificado"
        m_Explicacion = "Todas las acciones tienen fecha de fin prevista "
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td><b>" & m_Apartado & "</b></td>" & vbNewLine
            m_mensaje = m_mensaje & "<td>" & m_Explicacion & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_Apartado = "Plan activo"
        m_Explicacion = "Al menos una acción tiene fecha de inicio y la última acción no tiene fecha de fin real"
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td><b>" & m_Apartado & "</b></td>" & vbNewLine
            m_mensaje = m_mensaje & "<td>" & m_Explicacion & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_Apartado = "Plan finalizado"
        m_Explicacion = "Todas las acciones tienen fecha de fin real"
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td><b>" & m_Apartado & "</b></td>" & vbNewLine
            m_mensaje = m_mensaje & "<td>" & m_Explicacion & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        
    m_mensaje = m_mensaje & "</table>" & vbNewLine
    
    getHTMLTablaExplicacionesEstadosPlanes = m_mensaje
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLTablaExplicacionesEstadosPlanes ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getHTMLTablaExplicacionesEstados( _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    
    Dim m_mensaje As String
    Dim m_HTMLTablaExplicacionesEstadosPlanes As String
    Dim m_HTMLTablaExplicacionesEstadosRiesgos As String
    
    On Error GoTo errores
    
    m_HTMLTablaExplicacionesEstadosPlanes = getHTMLTablaExplicacionesEstadosPlanes(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTablaExplicacionesEstadosRiesgos = getHTMLTablaExplicacionesEstadosRiesgos(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_mensaje = m_HTMLTablaExplicacionesEstadosPlanes & vbNewLine
    m_mensaje = m_mensaje & "<Br>" & vbNewLine
    m_mensaje = m_mensaje & "<Br>" & vbNewLine
    m_mensaje = m_mensaje & m_HTMLTablaExplicacionesEstadosRiesgos & vbNewLine
    
    getHTMLTablaExplicacionesEstados = m_mensaje
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLTablaExplicacionesEstados ha devuelto el error: " & Err.Description
    End If
End Function



Public Function EnvioCorreoAltaProyecto( _
                                        ByRef p_ObjProyecto As Proyecto, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    Dim m_ObjCorreo As CORREO
    On Error GoTo errores
    
    
    
    Set m_ObjCorreo = New CORREO
    m_ObjCorreo.AltaProyecto p_ObjProyecto, p_Error
    Set m_ObjCorreo = Nothing
    If p_Error <> "" Then
       Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnvioCorreoAltaProyecto ha devuelto el error: " & Err.Description
    End If
End Function

Public Function EnvioCorreoNuevaPublicacion( _
                                            ByRef p_ObjProyecto As Proyecto, _
                                            p_URLInforme As String, _
                                            Optional p_ConEnvioAlRAC As EnumSiNo, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim m_ObjCorreo As CORREO
    On Error GoTo errores
    
    Set m_ObjCorreo = New CORREO
    m_ObjCorreo.NuevaPublicacion p_ObjProyecto, p_URLInforme, p_ConEnvioAlRAC, p_Error
    Set m_ObjCorreo = Nothing
    If p_Error <> "" Then
       Err.Raise 1000
    End If
   
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnvioCorreoNuevaPublicacion ha devuelto el error: " & Err.Description
    End If
End Function


Public Function EstablecerTareasCalidad( _
                                        Optional ByRef p_Reseteando As EnumSiNo, _
                                        Optional ByRef p_Regularizando As EnumSiNo, _
                                        Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    
    If p_Reseteando = Empty Then
        p_Reseteando = EnumSiNo.Sí
    End If
    If p_Reseteando = EnumSiNo.Sí Then
        Set m_ObjTareasCalidad = New TareasCalidad
    Else
        If m_ObjTareasCalidad Is Nothing Then
            Set m_ObjTareasCalidad = New TareasCalidad
        End If
    End If
    
    
    EstablecerContadoresCalidad p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerTareasCalidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function EstablecerTareasTecnico( _
                                        Optional ByRef p_Reseteando As EnumSiNo, _
                                        Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    
    If p_Reseteando = Empty Then
        p_Reseteando = EnumSiNo.Sí
    End If
    If p_Reseteando = EnumSiNo.Sí Then
        Set m_ObjTareasTecnico = New TareasTecnico
    Else
        If m_ObjTareasTecnico Is Nothing Then
            Set m_ObjTareasTecnico = New TareasTecnico
        End If
    End If
    
    
    EstablecerContadoresTecnico p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerTareasTecnico ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function EstablecerContadoresCalidad( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim m_Etiqueta As Label
    Dim m_Contador As String
    Dim m_FormPpal As Form

    On Error GoTo errores
   
    If m_ObjTareasCalidad Is Nothing Then
        p_Error = "El objeto tarea no está establecido"
        Err.Raise 1000
    End If
    If FormularioAbierto("Form0BDOpciones") Then
        Set m_FormPpal = Forms("Form0BDOpciones")
        Set m_Etiqueta = m_FormPpal.lblTareasPendientes
        m_Contador = m_ObjTareasCalidad.TareasTotales
        m_Etiqueta.Caption = "Tareas pendientes (#)"
        m_Etiqueta.Caption = Replace(m_Etiqueta.Caption, "#", m_Contador)
        If m_Contador = "0" Then
            m_Etiqueta.ForeColor = 16737792
        Else
            m_Etiqueta.ForeColor = 683236
        End If
    End If
    
    If FormularioAbierto("FormCalidadTareas") Then
        Form_FormCalidadTareas.CargarArbol p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerContadoresCalidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function EstablecerContadoresTecnico( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim m_Etiqueta As Label
    Dim m_Contador As String
    Dim m_FormPpal As Form

    On Error GoTo errores
   
    If m_ObjTareasTecnico Is Nothing Then
        p_Error = "El objeto tarea no está establecido"
        Err.Raise 1000
    End If
    
    If FormularioAbierto("Form0BDOpcionesTecnico") Then
        Set m_FormPpal = Forms("Form0BDOpcionesTecnico")
        Set m_Etiqueta = m_FormPpal.lblTareasPendientes
        m_Contador = m_ObjTareasTecnico.TareasTotales
        m_Etiqueta.Caption = "Tareas Pendientes (#)"
        m_Etiqueta.Caption = Replace(m_Etiqueta.Caption, "#", m_Contador)
        If m_Contador = "0" Then
            m_Etiqueta.ForeColor = 16737792
        Else
            m_Etiqueta.ForeColor = 683236
        End If
    End If
    
    If FormularioAbierto("FormTecnicoTareas") Then
        Form_FormTecnicoTareas.CargarArbol p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerContadoresTecnico ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function DameID( _
                        p_NombreTabla As String, _
                        p_NombreCampoID As String, _
                        Optional ByRef p_db As DAO.Database, _
                        Optional ByRef p_Error As String _
                        ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim lngIDMax As Long
    On Error GoTo errores
    
    If p_NombreTabla = "" Or p_NombreCampoID = "" Then
        p_Error = "Se ha de indicar el nombre de la tabla y de su campo ID"
        Err.Raise 1000
    End If
    If p_db Is Nothing Then
        Set p_db = getdb()
    End If
    m_SQL = "SELECT Max(" & p_NombreTabla & "." & p_NombreCampoID & ") AS MaxID " & _
            "FROM " & p_NombreTabla & ";"
    Set rcdDatos = p_db.OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(Nz(.Fields("MaxID"), "")) Then
                lngIDMax = .Fields("MaxID")
            End If

        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DameID = CStr(lngIDMax + 1)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameID ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function RellenarCombo( _
                                cmb As ComboBox, _
                                p_NombreTabla As String, _
                                p_NombreCampo As String, _
                                Optional p_db As DAO.Database, _
                                Optional ByRef p_Error As String _
                                ) As EnumSiNo
    
    Dim m_valor As Variant
    Dim col As Scripting.Dictionary
    On Error GoTo errores
    
    p_Error = ""
    cmb.RowSource = ""
    Set col = Constructor.getValoresDistintos(p_NombreTabla, p_NombreCampo, p_db, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If col Is Nothing Then
        Exit Function
    End If
    For Each m_valor In col
        cmb.AddItem m_valor
    Next
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo RellenarCombo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function RellenarComboRespCalidad( _
                                            cmb As ComboBox, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim m_Id As Variant
    Dim m_Usuario As Usuario
    On Error GoTo errores
    
    p_Error = ""
    cmb.RowSource = ""
    If m_ObjEntorno.ColUsuariosCalidad Is Nothing Then
        Exit Function
    End If
    For Each m_Id In m_ObjEntorno.ColUsuariosCalidad
        Set m_Usuario = m_ObjEntorno.ColUsuariosCalidad(m_Id)
        cmb.AddItem m_Usuario.ID & ";" & m_Usuario.Nombre
        Set m_Usuario = Nothing
    Next
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo RellenarComboRespCalidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function




Public Function CampoRepetido( _
                                    p_NombreTabla As String, _
                                    p_NombreCampo As String, _
                                    p_ValorCampo As String, _
                                    Optional p_ElCampoEsTexto As EnumSiNo = EnumSiNo.Sí, _
                                    Optional ByRef p_Error As String _
                                    ) As EnumSiNo
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    On Error GoTo errores
    
    If p_NombreTabla = "" Then
        p_Error = "Ha de indicar p_NombreTabla"
        Err.Raise 1000
    End If
    If p_NombreCampo = "" Then
        p_Error = "Ha de indicar p_NombreCampo"
        Err.Raise 1000
    End If
    If p_ValorCampo = "" Then
        p_Error = "Ha de indicar p_ValorCampo"
        Err.Raise 1000
    End If
    If p_ElCampoEsTexto = EnumSiNo.Sí Then
        m_SQL = "SELECT " & p_NombreCampo & " " & _
                "FROM " & p_NombreTabla & " " & _
                "WHERE " & p_NombreCampo & "='" & p_ValorCampo & "';"
    Else
        m_SQL = "SELECT " & p_NombreCampo & " " & _
                "FROM " & p_NombreTabla & " " & _
                "WHERE " & p_NombreCampo & "=" & p_ValorCampo & ";"
    End If
     
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            CampoRepetido = EnumSiNo.No
        Else
            CampoRepetido = EnumSiNo.Sí
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Proyecto.CampoRepetido ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function EnviarCorreoTecnicoRiesgoAceptado( _
                                                    ByRef p_ObjRiesgo As riesgo, _
                                                    Optional ByRef p_Error As String _
                                                    ) As CORREO
    
    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    Dim m_ObjUsuarioCalidad As Usuario
    
    On Error GoTo errores
    p_Error = ""
    m_Asunto = "Riesgo aceptado por el técnico pendiente aprobación por Calidad: " & m_ObjProyectoActivo.Proyecto & " (gestión de riesgos)"
    m_Cuerpo = p_ObjRiesgo.getHTMLTecnicoRiesgoAceptado(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjUsuarioCalidad = getUsuario(, , m_ObjProyectoActivo.NombreUsuarioCalidad, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Destinatarios = m_ObjUsuarioCalidad.CorreoUsuario
    
    'm_DestinatariosConCopia = m_ObjUsuarioConectadoInicialmente.CorreoUsuario & ";" & m_ObjEntorno.CorreoCalidad
    m_DestinatariosConCopia = m_ObjUsuarioConectadoInicialmente.CorreoUsuario
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set EnviarCorreoTecnicoRiesgoAceptado = New CORREO
    With EnviarCorreoTecnicoRiesgoAceptado
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoTecnicoRiesgoAceptado ha devuelto el error: " & Err.Description
    End If
End Function

Public Function EnviarCorreoTecnicoRiesgoRetirado( _
                                                    ByRef p_ObjRiesgo As riesgo, _
                                                    Optional ByRef p_Error As String _
                                                    ) As CORREO
    
    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    Dim m_ObjUsuarioCalidad As Usuario
    
    On Error GoTo errores
    p_Error = ""
    m_Asunto = "Riesgo retirado por el Técnico pendiente de aprobación por Calidad: " & m_ObjProyectoActivo.Proyecto & " (gestión de riesgos)"
    m_Cuerpo = p_ObjRiesgo.getHTMLTecnicoRiesgoRetirado(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjUsuarioCalidad = getUsuario(, , m_ObjProyectoActivo.NombreUsuarioCalidad, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Destinatarios = m_ObjUsuarioCalidad.CorreoUsuario
    m_DestinatariosConCopia = m_ObjUsuarioConectadoInicialmente.CorreoUsuario
    'm_DestinatariosConCopia = m_ObjUsuarioConectadoInicialmente.CorreoUsuario & ";" & m_ObjEntorno.CorreoCalidad
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set EnviarCorreoTecnicoRiesgoRetirado = New CORREO
    With EnviarCorreoTecnicoRiesgoRetirado
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoTecnicoRiesgoRetirado ha devuelto el error: " & Err.Description
    End If
End Function
Public Function EnviarCorreoCalidadApruebaRiesgoAceptado( _
                                                        ByRef p_ObjRiesgo As riesgo, _
                                                        Optional ByRef p_Error As String _
                                                        ) As CORREO

    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    
    
    On Error GoTo errores
    p_Error = ""
    m_Asunto = "Calidad aprueba riesgo aceptado por el técnico: " & p_ObjRiesgo.Edicion.Proyecto.Proyecto & " (gestión de riesgos)"
    
    m_Cuerpo = p_ObjRiesgo.getHTMLCalidadApruebaRiesgoAceptado(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Destinatarios = p_ObjRiesgo.Edicion.Proyecto.CadenaCorreoAutorizados
    p_Error = p_ObjRiesgo.Edicion.Proyecto.Error
    m_DestinatariosConCopia = m_ObjUsuarioConectado.CorreoUsuario
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set EnviarCorreoCalidadApruebaRiesgoAceptado = New CORREO
    With EnviarCorreoCalidadApruebaRiesgoAceptado
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoCalidadApruebaRiesgoAceptado ha devuelto el error: " & Err.Description
    End If
End Function

Public Function EnviarCorreoCalidadQuitarAprobacionRiesgoAceptado( _
                                                                    ByRef p_ObjRiesgo As riesgo, _
                                                                    Optional ByRef p_Error As String _
                                                                    ) As CORREO

    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    
    
    On Error GoTo errores
    p_Error = ""
    m_Asunto = "Calidad quita la aprobación del riesgo aceptado por el técnico: " & p_ObjRiesgo.Edicion.Proyecto.Proyecto & " (gestión de riesgos)"
    
    m_Cuerpo = p_ObjRiesgo.getHTMLCalidadQuitaAprobacionRiesgoAceptado(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Destinatarios = p_ObjRiesgo.Edicion.Proyecto.CadenaCorreoAutorizados
    p_Error = p_ObjRiesgo.Edicion.Proyecto.Error
    m_DestinatariosConCopia = m_ObjUsuarioConectado.CorreoUsuario
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set EnviarCorreoCalidadQuitarAprobacionRiesgoAceptado = New CORREO
    With EnviarCorreoCalidadQuitarAprobacionRiesgoAceptado
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoCalidadQuitarAprobacionRiesgoAceptado ha devuelto el error: " & Err.Description
    End If
End Function

Public Function EnviarCorreoCalidadRechazaRiesgoAceptado( _
                                                        ByRef p_ObjRiesgo As riesgo, _
                                                        Optional ByRef p_Error As String _
                                                        ) As CORREO

    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    
    
    On Error GoTo errores
    p_Error = ""
    m_Asunto = "Calidad rechaza riesgo aceptado por el técnico: " & p_ObjRiesgo.Edicion.Proyecto.Proyecto & " (gestión de riesgos)"
    
    m_Cuerpo = p_ObjRiesgo.getHTMLCalidadRechazaRiesgoAceptado(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Destinatarios = p_ObjRiesgo.Edicion.Proyecto.CadenaCorreoAutorizados
    p_Error = p_ObjRiesgo.Edicion.Proyecto.Error
    m_DestinatariosConCopia = m_ObjUsuarioConectado.CorreoUsuario
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set EnviarCorreoCalidadRechazaRiesgoAceptado = New CORREO
    With EnviarCorreoCalidadRechazaRiesgoAceptado
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoCalidadRechazaRiesgoAceptado ha devuelto el error: " & Err.Description
    End If
End Function
Public Function EnviarCorreoCalidadQuitarRechazoRiesgoAceptado( _
                                                        ByRef p_ObjRiesgo As riesgo, _
                                                        Optional ByRef p_Error As String _
                                                        ) As CORREO

    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    
    
    On Error GoTo errores
    p_Error = ""
    m_Asunto = "Calidad quita el rechazo del riesgo aceptado por el técnico: " & p_ObjRiesgo.Edicion.Proyecto.Proyecto & " (gestión de riesgos)"
    
    m_Cuerpo = p_ObjRiesgo.getHTMLCalidadQuitarRechazoRiesgoAceptado(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Destinatarios = p_ObjRiesgo.Edicion.Proyecto.CadenaCorreoAutorizados
    p_Error = p_ObjRiesgo.Edicion.Proyecto.Error
    m_DestinatariosConCopia = m_ObjUsuarioConectado.CorreoUsuario
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set EnviarCorreoCalidadQuitarRechazoRiesgoAceptado = New CORREO
    With EnviarCorreoCalidadQuitarRechazoRiesgoAceptado
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoCalidadQuitarRechazoRiesgoAceptado ha devuelto el error: " & Err.Description
    End If
End Function
Public Function EnviarCorreoCalidadApruebaRiesgoRetirado( _
                                                        ByRef p_ObjRiesgo As riesgo, _
                                                        Optional ByRef p_Error As String _
                                                        ) As CORREO

    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    
    
    On Error GoTo errores
    p_Error = ""
    m_Asunto = "Calidad aprueba riesgo retirado por el técnico: " & p_ObjRiesgo.Edicion.Proyecto.Proyecto & " (gestión de riesgos)"
    
    m_Cuerpo = p_ObjRiesgo.getHTMLCalidadApruebaRiesgoRetirado(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Destinatarios = p_ObjRiesgo.Edicion.Proyecto.CadenaCorreoAutorizados
    p_Error = p_ObjRiesgo.Edicion.Proyecto.Error
    m_DestinatariosConCopia = m_ObjUsuarioConectado.CorreoUsuario
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set EnviarCorreoCalidadApruebaRiesgoRetirado = New CORREO
    With EnviarCorreoCalidadApruebaRiesgoRetirado
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoCalidadApruebaRiesgoRetirado ha devuelto el error: " & Err.Description
    End If
End Function
Public Function EnviarCorreoCalidadQuitarAprobacionRiesgoRetirado( _
                                                                ByRef p_ObjRiesgo As riesgo, _
                                                                Optional ByRef p_Error As String _
                                                                ) As CORREO

    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    
    
    On Error GoTo errores
    p_Error = ""
    m_Asunto = "Calidad quita la aprobación del riesgo retirado por el técnico: " & p_ObjRiesgo.Edicion.Proyecto.Proyecto & " (gestión de riesgos)"
    
    m_Cuerpo = p_ObjRiesgo.getHTMLCalidadQuitaAprobacionRiesgoRetirado(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Destinatarios = p_ObjRiesgo.Edicion.Proyecto.CadenaCorreoAutorizados
    p_Error = p_ObjRiesgo.Edicion.Proyecto.Error
    m_DestinatariosConCopia = m_ObjUsuarioConectado.CorreoUsuario
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set EnviarCorreoCalidadQuitarAprobacionRiesgoRetirado = New CORREO
    With EnviarCorreoCalidadQuitarAprobacionRiesgoRetirado
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoCalidadQuitarAprobacionRiesgoRetirado ha devuelto el error: " & Err.Description
    End If
End Function
Public Function EnviarCorreoCalidadRechazaRiesgoRetirado( _
                                                        ByRef p_ObjRiesgo As riesgo, _
                                                        Optional ByRef p_Error As String _
                                                        ) As CORREO

    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    
    
    On Error GoTo errores
    p_Error = ""
    m_Asunto = "Calidad rechaza la retirada del riesgo por el técnico: " & p_ObjRiesgo.Edicion.Proyecto.Proyecto & " (gestión de riesgos)"
    
    m_Cuerpo = p_ObjRiesgo.getHTMLCalidadRechazaRiesgoRetirado(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Destinatarios = p_ObjRiesgo.Edicion.Proyecto.CadenaCorreoAutorizados
    p_Error = p_ObjRiesgo.Edicion.Proyecto.Error
    m_DestinatariosConCopia = m_ObjUsuarioConectado.CorreoUsuario
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set EnviarCorreoCalidadRechazaRiesgoRetirado = New CORREO
    With EnviarCorreoCalidadRechazaRiesgoRetirado
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoCalidadRechazaRiesgoRetirado ha devuelto el error: " & Err.Description
    End If
End Function

Public Function EnviarCorreoCalidadQuitarRechazoRiesgoRetirado( _
                                                        ByRef p_ObjRiesgo As riesgo, _
                                                        Optional ByRef p_Error As String _
                                                        ) As CORREO

    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    
    
    On Error GoTo errores
    p_Error = ""
    m_Asunto = "Calidad quita el rechazo de la retirada del riesgo por el técnico: " & p_ObjRiesgo.Edicion.Proyecto.Proyecto & " (gestión de riesgos)"
    
    m_Cuerpo = p_ObjRiesgo.getHTMLCalidadQuitarRechazoRiesgoRetirado(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Destinatarios = p_ObjRiesgo.Edicion.Proyecto.CadenaCorreoAutorizados
    p_Error = p_ObjRiesgo.Edicion.Proyecto.Error
    m_DestinatariosConCopia = m_ObjUsuarioConectado.CorreoUsuario
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set EnviarCorreoCalidadQuitarRechazoRiesgoRetirado = New CORREO
    With EnviarCorreoCalidadQuitarRechazoRiesgoRetirado
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoCalidadQuitarRechazoRiesgoRetirado ha devuelto el error: " & Err.Description
    End If
End Function

Public Function EnviarCorreoPropuestaNuevaPublicacion( _
                                                        ByRef p_ObjEdicion As Edicion, _
                                                        Optional ByRef p_Error As String _
                                                        ) As CORREO
    
    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    Dim m_Edicion As String
    Dim m_CorreoResponsableCalidad As String
    
    On Error GoTo errores
    p_Error = ""
    With p_ObjEdicion
        m_Edicion = .Edicion
        m_Asunto = "Propuesta de Publicación de " & m_Edicion & ": " & _
                    .Proyecto.NombreProyecto & " (gestión de riesgos)"
        m_Cuerpo = .HTMLEdicionPropuestaPublicacion
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_DestinatariosConCopia = .Proyecto.CadenaCorreoAutorizados
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_CorreoResponsableCalidad = .CorreoResponsableCalidad
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
    End With
    m_Destinatarios = m_CorreoResponsableCalidad
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set EnviarCorreoPropuestaNuevaPublicacion = New CORREO
    With EnviarCorreoPropuestaNuevaPublicacion
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoPropuestaNuevaPublicacion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function EnviarCorreoQuitarPropuestaNuevaPublicacion( _
                                                            ByRef p_ObjEdicion As Edicion, _
                                                            Optional ByRef p_Error As String _
                                                            ) As CORREO
    
    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    Dim m_Edicion As String
    
    
    On Error GoTo errores
    p_Error = ""
    
    With p_ObjEdicion
        m_Edicion = .Edicion
        m_Asunto = "Eliminada Propuesta de Publicación de " & m_Edicion & ": " & _
                    .Proyecto.NombreProyecto & " (gestión de riesgos)"
        m_Cuerpo = .HTMLEdicionQuitarPropuestaPublicacion
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_DestinatariosConCopia = .Proyecto.CadenaCorreoAutorizados
        p_Error = .Proyecto.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_Destinatarios = .CorreoResponsableCalidad
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set EnviarCorreoQuitarPropuestaNuevaPublicacion = New CORREO
    With EnviarCorreoQuitarPropuestaNuevaPublicacion
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoQuitarPropuestaNuevaPublicacion ha devuelto el error: " & Err.Description
    End If
End Function
Public Function EnviarCorreoRechazoPropuestaPublicacion( _
                                                            ByRef p_ObjEdicion As Edicion, _
                                                            Optional ByRef p_Error As String _
                                                            ) As CORREO
    
    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    Dim m_Edicion As String
    Dim m_CorreoResponsableCalidad As String
    
    On Error GoTo errores
    p_Error = ""
    
    With p_ObjEdicion
        m_Edicion = .Edicion
        m_Asunto = "Rechazada por Calidad Propuesta de Publicación de " & m_Edicion & ": " & _
                    .Proyecto.NombreProyecto & " (gestión de riesgos)"
        m_Cuerpo = .HTMLEdicionRechazarPropuestaPublicacion
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_CorreoResponsableCalidad = .CorreoResponsableCalidad
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_DestinatariosConCopia = m_CorreoResponsableCalidad
        m_Destinatarios = .Proyecto.CadenaCorreoAutorizados
        p_Error = .Proyecto.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set EnviarCorreoRechazoPropuestaPublicacion = New CORREO
    With EnviarCorreoRechazoPropuestaPublicacion
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoRechazoPropuestaPublicacion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function EnviarCorreoRiesgoRequiereRetipificacion( _
                                                        ByRef p_ObjRiesgo As riesgo, _
                                                        Optional ByRef p_Error As String _
                                                        ) As CORREO
    
    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    Dim m_ObjUsuarioCalidad As Usuario
    
    On Error GoTo errores
    p_Error = ""
    m_Asunto = "Riesgo No encontrado en Biblioteca por el Técnico: " & m_ObjProyectoActivo.Proyecto & " (gestión de riesgos)"
    m_Cuerpo = p_ObjRiesgo.HTMLRiesgoPorRetipificar
    p_Error = p_ObjRiesgo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjUsuarioCalidad = getUsuario(, , m_ObjProyectoActivo.NombreUsuarioCalidad, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Destinatarios = m_ObjUsuarioCalidad.CorreoUsuario
    
    m_DestinatariosConCopia = m_ObjUsuarioConectado.CorreoUsuario
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set EnviarCorreoRiesgoRequiereRetipificacion = New CORREO
    With EnviarCorreoRiesgoRequiereRetipificacion
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoRiesgoRequiereRetipificacion ha devuelto el error: " & Err.Description
    End If
End Function
Public Function EnviarCorreoRiesgoRetipificacion( _
                                                ByRef p_ObjRiesgo As riesgo, _
                                                ByRef p_ObjRiesgoAlInicio As riesgo, _
                                                Optional ByRef p_Error As String _
                                                ) As CORREO
    
    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    Dim m_ObjUsuarioCalidad As Usuario
    
    On Error GoTo errores
    p_Error = ""
    With p_ObjRiesgo
        m_Asunto = "Riesgo Asignado de la Biblioteca por Calidad: " & .Edicion.Proyecto.Proyecto & _
                    " (gestión de riesgos)"
        m_Cuerpo = .getHTMLRiesgoRetipificado(p_ObjRiesgoAlInicio, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ObjUsuarioCalidad = getUsuario(, , p_ObjRiesgo.Edicion.Proyecto.NombreUsuarioCalidad, , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_Destinatarios = .Edicion.Proyecto.CadenaCorreoAutorizados
        p_Error = .Edicion.Proyecto.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    m_DestinatariosConCopia = m_ObjUsuarioConectado.CorreoUsuario
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set EnviarCorreoRiesgoRetipificacion = New CORREO
    With EnviarCorreoRiesgoRetipificacion
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoRiesgoRetipificacion ha devuelto el error: " & Err.Description
    End If
End Function
Public Function EnviarCorreoRiesgoMaterializado( _
                                                ByRef p_ObjRiesgo As riesgo, _
                                                Optional ByRef p_Error As String _
                                                ) As CORREO
    
    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    Dim m_ObjUsuarioCalidad As Usuario
    Dim m_CadenaCorreoAutorizados As String
    
    On Error GoTo errores
    p_Error = ""
    m_Asunto = "Riesgo Materializado: " & m_ObjProyectoActivo.Proyecto & " (gestión de riesgos)"
    m_Cuerpo = p_ObjRiesgo.HTMLRiesgoMaterializado
    p_Error = p_ObjRiesgo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjUsuarioCalidad = getUsuario(, , m_ObjProyectoActivo.NombreUsuarioCalidad, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Destinatarios = m_ObjEntorno.CorreoCalidad
    m_CadenaCorreoAutorizados = p_ObjRiesgo.Edicion.Proyecto.CadenaCorreoAutorizados
    p_Error = p_ObjRiesgo.Edicion.Proyecto.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_DestinatariosConCopia = m_ObjUsuarioConectado.CorreoUsuario & ";" & m_CadenaCorreoAutorizados
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set EnviarCorreoRiesgoMaterializado = New CORREO
    With EnviarCorreoRiesgoMaterializado
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoRiesgoMaterializado ha devuelto el error: " & Err.Description
    End If
End Function

Public Function EnviarCorreoRevisionEdicion( _
                                                ByRef p_ObjEdicion As Edicion, _
                                                Optional ByRef p_Error As String _
                                                ) As CORREO
    
    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    Dim m_ObjUsuarioCalidad As Usuario
    Dim m_CadenaCorreoAutorizados As String
    
    On Error GoTo errores
    p_Error = ""
    m_Asunto = "Revisa edición : " & p_ObjEdicion.Proyecto.Proyecto & " (gestión de riesgos)"
    m_Cuerpo = p_ObjEdicion.HTMLEdicionRevision
    p_Error = p_ObjEdicion.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjUsuarioCalidad = getUsuario(, , p_ObjEdicion.Proyecto.NombreUsuarioCalidad, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_CadenaCorreoAutorizados = p_ObjEdicion.Proyecto.CadenaCorreoAutorizados
    p_Error = p_ObjEdicion.Proyecto.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Destinatarios = m_CadenaCorreoAutorizados
    
    m_DestinatariosConCopia = m_ObjUsuarioConectado.CorreoUsuario
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set EnviarCorreoRevisionEdicion = New CORREO
    With EnviarCorreoRevisionEdicion
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoRevisionEdicion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function EnviarCorreoAltaProyecto( _
                                            ByRef p_ObjProyecto As Proyecto, _
                                            Optional ByRef p_Error As String _
                                            ) As String

    
    On Error GoTo errores
    
    getCorreoByProyecto p_ObjProyecto, p_Error
    If p_Error <> "" Then
       Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoAltaProyecto ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getCorreoByProyecto( _
                                        p_ObjProyecto As Proyecto, _
                                        Optional ByRef p_Error As String _
                                        ) As CORREO
                                
    Dim m_ObjCorreo As CORREO
    On Error GoTo errores
    
    Set m_ObjCorreo = New CORREO
    Set getCorreoByProyecto = m_ObjCorreo.SetCorreoAltaProyecto(p_ObjProyecto, p_Error)
    Set m_ObjCorreo = Nothing
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getCorreoByProyecto ha devuelto el error: " & Err.Description
    End If
End Function



Public Function getCorreoNuevaPublicacion( _
                                            p_ObjProyecto As Proyecto, _
                                            p_URLInforme As String, _
                                            Optional p_ConEnvioAlRAC As EnumSiNo, _
                                            Optional ByRef p_Error As String _
                                            ) As CORREO
                                
    
    Dim m_ObjCorreo As CORREO
    On Error GoTo errores
    
    Set m_ObjCorreo = New CORREO
    Set getCorreoNuevaPublicacion = m_ObjCorreo.SetCorreoNuevaPublicacion(p_ObjProyecto, p_URLInforme, p_ConEnvioAlRAC, p_Error)
    Set m_ObjCorreo = Nothing
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
    
    
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getCorreoNuevaPublicacion ha devuelto el error: " & Err.Description
    End If
End Function




Public Function EstaEnElIntervaloDado( _
                                        m_FechaInicialIntervalo As String, _
                                        m_FechaFinalIntervalo As String, _
                                        Optional m_FechaInicial As String, _
                                        Optional m_FechaFinal, _
                                        Optional ByRef p_Error As String _
                                        ) As EnumSiNo
    
    
    On Error GoTo errores
    
    p_Error = ""
    If Not IsDate(m_FechaInicialIntervalo) Then
        p_Error = "Se ha de indicar la fecha inicial del intervalo"
        Err.Raise 1000
    End If
    If Not IsDate(m_FechaFinalIntervalo) Then
        p_Error = "Se ha de indicar la fecha final del intervalo"
        Err.Raise 1000
    End If
    If CDate(m_FechaFinalIntervalo) < CDate(m_FechaInicialIntervalo) Then
        p_Error = "Se ha de indicar la fecha final del intervalo posterior o igual a la inicial"
        Err.Raise 1000
    End If
    If IsDate(m_FechaInicial) And IsDate(m_FechaFinal) Then
        If CDate(m_FechaFinal) < CDate(m_FechaInicial) Then
            p_Error = "Las fechas iniciales y finales de expediente no son obligatorias, pero de rellenarse la final ha de ser posterior a la inicial."
            Err.Raise 1000
        End If
    End If
    If IsDate(m_FechaInicial) And IsDate(m_FechaFinal) Then
        If CDate(m_FechaInicial) < CDate(m_FechaInicialIntervalo) And CDate(m_FechaFinal) < CDate(m_FechaInicialIntervalo) Then
            '---------------------
            ' cualquiera de las fechas de inicio y fin de expediente son anteriores de la fecha inicial y final del intervalo
            '---------------------
            EstaEnElIntervaloDado = EnumSiNo.No
        ElseIf (CDate(m_FechaInicial) >= CDate(m_FechaInicialIntervalo) And CDate(m_FechaInicial) <= CDate(m_FechaFinalIntervalo)) Or _
                (CDate(m_FechaFinal) >= CDate(m_FechaInicialIntervalo) And CDate(m_FechaFinal) <= CDate(m_FechaFinalIntervalo)) Then
            '---------------------
            ' cualquiera de las fechas de inicio y fin de expediente está en el intervalo entre fecha inicial intervalo y final intervalo
            '---------------------
            EstaEnElIntervaloDado = EnumSiNo.Sí
        ElseIf CDate(m_FechaInicial) > CDate(m_FechaFinalIntervalo) Then
            '---------------------
            ' cualquiera de las fechas de inicio y fin de expediente son posteriores de la fecha inicial y final del intervalo
            '---------------------
            EstaEnElIntervaloDado = EnumSiNo.No
        ElseIf CDate(m_FechaInicial) < CDate(m_FechaInicialIntervalo) And CDate(m_FechaFinal) > CDate(m_FechaFinalIntervalo) Then
            '---------------------
            ' El expediente empieza antes del inicio del intervalo y acaba después del fin del intervalo
            '---------------------
            EstaEnElIntervaloDado = EnumSiNo.Sí
        Else
        
            p_Error = "Situación de fechas desconocida"
            Err.Raise 1000
        End If
    ElseIf IsDate(m_FechaInicial) And Not IsDate(m_FechaFinal) Then
        EstaEnElIntervaloDado = EnumSiNo.Sí
    ElseIf Not IsDate(m_FechaInicial) And IsDate(m_FechaFinal) Then
        '---------------------
        ' la fecha final de expediente ha de estar entre la de inicio y final de intervalo
        '---------------------
        If CDate(m_FechaFinal) >= CDate(m_FechaInicialIntervalo) And CDate(m_FechaFinal) <= CDate(m_FechaFinalIntervalo) Then
            EstaEnElIntervaloDado = EnumSiNo.Sí
        Else
            EstaEnElIntervaloDado = EnumSiNo.No
        End If
    Else
        '---------------------
        ' Expediente sin fecha de firmadecontrato y sin fecharecepción adjudicado, cualquier valor de fechas de intervalo entra
        '---------------------
        EstaEnElIntervaloDado = EnumSiNo.Sí
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstaEnElIntervaloDado ha devuelto el error : " & vbNewLine & Err.Description
    End If
End Function


Public Function ValorRedondeado( _
                                    p_Valor As Double, _
                                    p_NumeroDecimales As Integer, _
                                    Optional ByRef p_Error As String _
                                    ) As Double
    Dim m_ParteDecimal As String
    Dim m_ParteNoDecimal As String
    Dim m_DecimalSiguiente As Integer
    Dim m_UltimoDecimal As Integer
    Dim dato
    On Error GoTo errores
    
    p_Error = ""
    If p_NumeroDecimales > 0 Then
        If InStr(1, p_Valor, ",") <> 0 Or InStr(1, p_Valor, ".") <> 0 Then
            If InStr(1, p_Valor, ",") <> 0 Then
                dato = Split(p_Valor, ",")
                m_ParteNoDecimal = dato(0)
                m_ParteDecimal = dato(1)
            Else
                If InStr(1, p_Valor, ".") <> 0 Then
                    dato = Split(p_Valor, ".")
                    m_ParteNoDecimal = dato(0)
                    m_ParteDecimal = dato(1)
                End If
            End If
            If Len(m_ParteDecimal) > p_NumeroDecimales Then
                m_UltimoDecimal = Mid(m_ParteDecimal, p_NumeroDecimales, 1)
                m_DecimalSiguiente = Mid(m_ParteDecimal, p_NumeroDecimales + 1, 1)
                If m_DecimalSiguiente >= 5 Then
                    m_UltimoDecimal = m_UltimoDecimal + 1
                End If
                ValorRedondeado = CDbl(m_ParteNoDecimal & "," & Left(m_ParteDecimal, p_NumeroDecimales - 1) & m_UltimoDecimal)
            Else
                ValorRedondeado = p_Valor
            End If
        Else
            ValorRedondeado = p_Valor
        End If
    Else
        ValorRedondeado = p_Valor
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ValorRedondeado ha devuelto el error: " & Err.Description
    End If
End Function
Public Function HTMLENTXT( _
                            Optional p_HTML As String, _
                            Optional m_mensaje As ADODB.stream, _
                            Optional p_EnFormulario As EnumSiNo = EnumSiNo.Sí, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim F1 As Object
    Dim m_URLHTML As String
    Dim m_URLTXT As String
    Dim m_URLCompletaArchivo As String
    Dim m_Nombre As String
    Dim mHWD As Long
    On Error GoTo errores
    
    If p_HTML = "" And m_mensaje Is Nothing Then
        p_Error = "No se ha indicado el HTML"
        Err.Raise 1000
    End If
       
    DameUntxtYHtml m_URLTXT, m_URLHTML, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
   
    If p_HTML <> "" Then
        Set F1 = FSO.CreateTextFile(m_URLTXT, True)
        F1.WriteLine p_HTML
        F1.Close
        FSO.GetFile(m_URLTXT).Name = FSO.GetBaseName(m_URLTXT) & ".html"
    Else
        m_mensaje.SaveToFile m_URLHTML
    End If
    If p_EnFormulario = EnumSiNo.Sí Then
        m_URLHTMLActivo = m_URLHTML
        If FormularioAbierto("FormWeb") Then
            DoCmd.Close acForm, "FormWeb", acSaveNo
        End If
        DoCmd.OpenForm "FormWeb"
    Else
        On Error Resume Next
        mHWD = Screen.ActiveForm.hWnd
        If Err.Number <> 0 Then
            Err.Clear
            mHWD = 1
        End If
        On Error GoTo errores
        Ejecutar mHWD, "open", m_URLHTML, "", "", 1
    End If
    
    HTMLENTXT = m_URLHTML
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTMLENTXT ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
        
    
End Function

Private Function DameUntxtYHtml( _
                                ByRef p_URLTXT As String, _
                                ByRef p_URLHTML As String, _
                                Optional ByRef p_Error As String) As String
    
    Dim m_URLHTML As String
    Dim m_URLTXT As String
    Dim i As Integer
    
    
    Dim m_NombreHTML As String
    Dim m_Nombretxt As String
    
    On Error GoTo errores
    BorraHTMLs
    For i = 1 To 50
        m_Nombretxt = "HTML" & i & ".txt"
        m_NombreHTML = "HTML" & i & ".html"
        m_URLTXT = m_ObjEntorno.URLDirectorioLocal & m_Nombretxt
        m_URLHTML = m_ObjEntorno.URLDirectorioLocal & m_NombreHTML
        If Not FSO.FileExists(m_URLHTML) And Not FSO.FileExists(m_URLHTML) Then
            p_URLTXT = m_URLTXT
            p_URLHTML = m_URLHTML
            Exit Function
        End If
    Next
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameUntxtYHtml ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Private Function BorraHTMLs( _
                            Optional ByRef p_Error As String) As String
    
    Dim fichero As File
    
    On Error GoTo errores
    
    For Each fichero In FSO.GetFolder(m_ObjEntorno.URLDirectorioLocal).Files
        If FSO.GetExtensionName(fichero.Path) = "html" Or FSO.GetExtensionName(fichero.Path) = "htm" Then
            If Not FicheroAbierto(fichero.Path) Then
                FSO.DeleteFile fichero.Path
            End If
        End If
    Next
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método BorraHTMLs ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function


Public Function getFechaDetectadoMaterializadoRiesgo( _
                                                        p_CodUnico As String, _
                                                        Optional ByRef p_ObjRiesgoDetectado As riesgo, _
                                                        Optional ByRef p_ObjRiesgoMaterializado As riesgo, _
                                                        Optional ByRef p_Error As String _
                                                        ) As String
    Dim m_IdRiesgo As Variant
    Dim m_ObjRiesgo As riesgo
    Dim m_objEdicion As Edicion
    Dim m_NumeroEdiciones As Integer
    Dim m_ObjProyecto As Proyecto
    Dim m_IDProyecto As String
    Dim m_IDEdicion As Variant
    Dim m_FechaDetectado As String
    Dim m_FechaMaterializado As String
    Dim dato
    On Error GoTo errores
    If InStr(1, p_CodUnico, "R") = 0 Then
        p_Error = "El código único no tiene el formato adecuado"
        Err.Raise 1000
    End If
    dato = Split(p_CodUnico, "R")
    m_IDProyecto = CStr(CInt(dato(0)))
    Set m_ObjProyecto = getProyecto(m_IDProyecto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    For Each m_IDEdicion In m_ObjProyecto.colEdiciones.keys
        Set m_objEdicion = m_ObjProyecto.colEdiciones(m_IDEdicion)
        For Each m_IdRiesgo In m_objEdicion.colRiesgos.keys
            Set m_ObjRiesgo = m_objEdicion.colRiesgos(m_IdRiesgo)
            If m_ObjRiesgo.CodigoUnico = p_CodUnico Then
                If p_ObjRiesgoDetectado Is Nothing Then
                    If m_ObjRiesgo.FechaDetectado <> "" Then
                        Set p_ObjRiesgoDetectado = m_ObjRiesgo
                    End If
                End If
                If p_ObjRiesgoMaterializado Is Nothing Then
                    If m_ObjRiesgo.FechaMaterializado <> "" Then
                        Set p_ObjRiesgoMaterializado = m_ObjRiesgo
                    End If
                End If
            End If
            Set m_ObjRiesgo = Nothing
        Next
        Set m_objEdicion = Nothing
    Next
    m_FechaDetectado = p_ObjRiesgoDetectado.FechaDetectado
    If Not p_ObjRiesgoMaterializado Is Nothing Then
        m_FechaMaterializado = p_ObjRiesgoMaterializado.FechaMaterializado
    End If
    
    getFechaDetectadoMaterializadoRiesgo = m_FechaDetectado & "|" & m_FechaMaterializado
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getFechaDetectadoMaterializadoRiesgo ha devuelto el error: " & Err.Description
    End If
End Function



Public Function RegistrarUltimoProyecto( _
                                            Optional ByRef p_Error As String _
                                            ) As String
                                        
    
   
    
    
    On Error GoTo errores
    p_Error = ""
    If m_ObjProyectoActivo Is Nothing Or m_ObjUsuarioConectado Is Nothing Then
        Exit Function
    End If
    Set m_ObjUltimoProyecto = New UltimoProyecto
    With m_ObjUltimoProyecto
        .IDProyecto = m_ObjProyectoActivo.IDProyecto
        .Usuario = m_ObjUsuarioConectado.UsuarioRed
        .Registrar m_ObjProyectoActivo.IDProyecto, m_ObjUsuarioConectado.UsuarioRed, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    PintarUltimoProyecto m_ObjUltimoProyecto.Proyecto.NombreProyecto, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistrarUltimoProyecto ha devuelto el error: " & Err.Description
    End If
End Function

Public Function PintarUltimoProyecto( _
                                        Optional p_NombreProyecto As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                        
    Dim m_ObjProyecto As Proyecto
    
    Dim m_Etiqueta As Label
    
    On Error GoTo errores
    p_Error = ""
    
   
    If FormularioAbierto("Form0BDOpciones") Then
        Set m_Etiqueta = Forms("Form0BDOpciones").lblUltimoProyecto
    ElseIf FormularioAbierto("Form0BDOpcionesTecnico") Then
        Set m_Etiqueta = Forms("Form0BDOpcionesTecnico").lblUltimoProyecto
    End If
    If p_NombreProyecto = "" Then p_NombreProyecto = "Último Proyecto"
    m_Etiqueta.Caption = p_NombreProyecto
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PintarUltimoProyecto ha devuelto el error: " & Err.Description
    End If
End Function

Public Function ProyectoYaExistente( _
                                        p_Proyecto As String, _
                                        Optional ByRef p_Error As String _
                                        ) As EnumSiNo
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    If p_Proyecto = "" Then
        p_Error = "No se ha indicado el p_Proyecto"
        Err.Raise 1000
    End If
    m_SQL = "SELECT * FROM TbProyectos " & _
        "WHERE Proyecto='" & p_Proyecto & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            ProyectoYaExistente = EnumSiNo.No
        Else
            ProyectoYaExistente = EnumSiNo.Sí
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ProyectoYaExistente ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Function ValidacionEdicionNuevaDeAnterior( _
                                                    p_ObjEdicion As Edicion, _
                                                    Optional p_Error As String _
                                                ) As String
    
    On Error GoTo errores
    
    '-------------------------------------------------
    ' COMPROBACIONES PREVIAS
    '-------------------------------------------------
    
    
    If p_ObjEdicion Is Nothing Then
        p_Error = "No se conoce la Edición de partida"
        Err.Raise 1000
    End If
    If p_ObjEdicion.EsUltimaEdicionCalculado = EnumSiNo.No Then
        p_Error = "No es la última Edición"
        Err.Raise 1000
    End If
    If p_ObjEdicion.Proyecto Is Nothing Then
        p_Error = "No se conoce el Proyecto de la Edición de partida"
        Err.Raise 1000
    End If
    If p_ObjEdicion.Proyecto.FechaCierre <> "" Then
        p_Error = "La gestión de riesgos ya está cerrada"
        Err.Raise 1000
    End If
    If p_ObjEdicion.FechaPublicacion = "" Then
        p_Error = "La Edición no está publicada"
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ValidacionEdicionNuevaDeAnterior ha devuelto el error: " & Err.Description
    End If
End Function

Public Function GenerarEdicionNuevaAPartirDeAnterior( _
                                                        p_IDEdicion As String, _
                                                        Optional p_FechaSiguientePublicacion As String, _
                                                        Optional p_db As DAO.Database, _
                                                        Optional p_Error As String _
                                                        ) As Edicion
    Dim m_objEdicion As Edicion
    Dim m_ObjProyecto As Proyecto
    Dim m_ObjEdicionNueva As Edicion
    Dim m_objColRiesgos As Scripting.Dictionary
    Dim m_IdRiesgo As Variant
    Dim m_ObjRiesgo As riesgo
    Dim m_IDPX As Variant
    Dim m_ObjPM As PM
    Dim m_ObjPC As PC
    Dim m_IdAccion As Variant
    Dim m_ObjPMAccion As PMAccion
    Dim m_ObjPCAccion As PCAccion
    Dim m_IDEdicionNueva As String
    Dim m_ObjRiesgoNuevo As riesgo
    Dim m_IDRiesgoNuevo As String
    Dim m_ObjPMNuevo As PM
    Dim m_IDPMNuevo As String
    Dim m_ObjPMAccionNueva As PMAccion
    Dim m_IDPMAccionNueva As String
    Dim m_ObjPCNuevo As PC
    Dim m_IDPCNuevo As String
    Dim m_ObjPCAccionNueva As PCAccion
    Dim m_IDPCAccionNueva As String
    Dim m_PriorizacionManual As EnumSiNo
    
    
    On Error GoTo errores
    
    Set m_objEdicion = Constructor.getEdicion(p_IDEdicion, p_Error, p_db)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    ValidacionEdicionNuevaDeAnterior m_objEdicion, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjProyecto = m_objEdicion.Proyecto
    p_Error = m_objEdicion.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjProyecto Is Nothing Then
        p_Error = "No se ha podido obtener el proyecto de la edición"
        Err.Raise 1000
    End If
    Set m_ObjEdicionNueva = CopiarEdicion(p_IDEdicion, m_ObjProyecto.IDProyecto, p_db, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_IDEdicionNueva = m_ObjEdicionNueva.IDEdicion
    m_PriorizacionManual = m_ObjProyecto.RequiereRiesgoDeBibliotecaCalculado
    Set m_objColRiesgos = Nothing
    Set m_objColRiesgos = m_objEdicion.colRiesgos
    p_Error = m_objEdicion.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    For Each m_IdRiesgo In m_objColRiesgos.keys
        If Nz(m_IdRiesgo, "") = "" Then
            GoTo siguienteRiesgo
        End If
        Set m_ObjRiesgo = m_objColRiesgos(m_IdRiesgo)
        Set m_ObjRiesgoNuevo = CopiarRiesgo(CStr(m_IdRiesgo), m_IDEdicionNueva, m_PriorizacionManual, p_db, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_IDRiesgoNuevo = m_ObjRiesgoNuevo.IDRiesgo
        If m_ObjRiesgo.TienePMs = EnumSiNo.Sí Then
            For Each m_IDPX In m_ObjRiesgo.ColPMs.keys
                If Nz(m_IDPX, "") = "" Then
                    GoTo siguientePM
                End If
                Set m_ObjPM = m_ObjRiesgo.ColPMs(m_IDPX)
                Set m_ObjPMNuevo = CopiarPM(m_ObjPM.IDMitigacion, m_IDRiesgoNuevo, p_db, p_Error)
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If Not m_ObjPMNuevo Is Nothing Then
                    m_IDPMNuevo = m_ObjPMNuevo.IDMitigacion
                    If m_ObjPM.TieneAcciones = EnumSiNo.Sí Then
                        For Each m_IdAccion In m_ObjPM.colAcciones.keys
                            If Nz(m_IdAccion, "") = "" Then
                                GoTo siguientePMAccion
                            End If
                            Set m_ObjPMAccion = m_ObjPM.colAcciones(m_IdAccion)
                            Set m_ObjPMAccionNueva = CopiarPMAccion(CStr(m_IdAccion), m_IDPMNuevo, p_db, p_Error)
                            If p_Error <> "" Then
                                Err.Raise 1000
                            End If
                            Set m_ObjPMAccion = Nothing
                            Set m_ObjPMAccionNueva = Nothing
siguientePMAccion:
                        Next
                    End If
                End If
                
                Set m_ObjPM = Nothing
                Set m_ObjPMNuevo = Nothing
siguientePM:
            Next
        End If
        If m_ObjRiesgo.TienePCs = EnumSiNo.Sí Then
            For Each m_IDPX In m_ObjRiesgo.ColPCs.keys
                If Nz(m_IDPX, "") = "" Then
                    GoTo siguientePC
                End If
                Set m_ObjPC = m_ObjRiesgo.ColPCs(m_IDPX)
                Set m_ObjPCNuevo = CopiarPC(CStr(m_IDPX), m_IDRiesgoNuevo, p_db, p_Error)
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                m_IDPCNuevo = m_ObjPCNuevo.IDContingencia
                If Not m_ObjPC.colAcciones Is Nothing Then
                    For Each m_IdAccion In m_ObjPC.colAcciones.keys
                        If Nz(m_IdAccion, "") = "" Then
                            GoTo siguientePCAccion
                        End If
                        Set m_ObjPCAccion = m_ObjPC.colAcciones(m_IdAccion)
                        Set m_ObjPCAccionNueva = CopiarPCAccion(CStr(m_IdAccion), m_IDPCNuevo, p_db, p_Error)
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                        Set m_ObjPCAccion = Nothing
                        Set m_ObjPCAccionNueva = Nothing
siguientePCAccion:
                    Next
                End If
                Set m_ObjPC = Nothing
                Set m_ObjPCNuevo = Nothing
siguientePC:
            Next
        End If
        
        Set m_ObjRiesgo = Nothing
siguienteRiesgo:
    Next
    
    SetFechaMaximaPublicacion p_IDProyecto:=m_ObjProyecto.IDProyecto, _
                            p_FechaSiguientePublicacion:=p_FechaSiguientePublicacion, _
                            p_db:=p_db, _
                            p_Error:=p_Error
    Set m_ObjProyecto = m_ObjEdicionNueva.Proyecto
    p_Error = m_ObjEdicionNueva.Error
    
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjEdicionNueva = Constructor.getEdicion(m_IDEdicionNueva, p_Error, p_db)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    TrasladarSuminstradoresAEdicion m_ObjEdicionNueva, p_db, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set GenerarEdicionNuevaAPartirDeAnterior = m_ObjEdicionNueva
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarEdicionNuevaAPartirDeAnterior ha devuelto el error: " & Err.Description
    End If
    If Not m_ObjEdicionNueva Is Nothing Then
        m_ObjEdicionNueva.Borrar
    End If
End Function

Public Function TrasladarSuminstradoresAEdicion( _
                                                p_Edicion As Edicion, _
                                                Optional p_db As DAO.Database, _
                                                Optional p_Error As String _
                                                ) As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ProyectoSuministrador As ProyectoSuministrador
    Dim m_Id As Variant
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_db As DAO.Database
    
    On Error GoTo errores
    
    If p_db Is Nothing Then
        Set m_db = getdb()
    Else
        Set m_db = p_db
    End If
    
    Set m_Col = Constructor.getSuministradoresParaCalidadEnProyecto(p_Edicion.IDProyecto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_Id In m_Col
        Set m_ProyectoSuministrador = m_Col(m_Id)
        m_Id = DameID("TbProyectosEdicionesSuministradores", "ID", m_db, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_SQL = "SELECT * " & _
                "FROM TbProyectosEdicionesSuministradores " & _
                "WHERE " & _
                "IDEdicion=" & p_Edicion.IDEdicion & " " & _
                "AND IDSuministrador=" & m_ProyectoSuministrador.IDSuministrador & ";"
        Set rcdDatos = m_db.OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                
                .AddNew
                    .Fields("ID") = m_Id
                    .Fields("IDEdicion") = p_Edicion.IDEdicion
                    .Fields("IDSuministrador") = m_ProyectoSuministrador.IDSuministrador
                .Update
            End If
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        Set m_ProyectoSuministrador = Nothing
    Next
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método TrasladarSuminstradoresAEdicion ha devuelto el error: " & Err.Description
    End If
    
End Function

Private Function CopiarEdicion( _
                                p_IDEdicion As String, _
                                p_IDProyecto As String, _
                                Optional p_db As DAO.Database, _
                                Optional ByRef p_Error As String _
                                 ) As Edicion
    
    Dim m_objEdicion As Edicion
    Dim m_ObjEdicionUltima As Edicion
    Dim m_ObjProyecto As Proyecto
    Dim m_Edicion As String
    Dim rcdDatosOrigen As DAO.Recordset
    Dim rcdDatosDestino As DAO.Recordset
    Dim m_SQL As String
    Dim m_Id As String
    Dim m_FechaMax As String
    Dim m_FechaRef As Date
    Dim m_db As DAO.Database
    
    On Error GoTo errores
    If p_db Is Nothing Then
        Set m_db = getdb()
    Else
        Set m_db = p_db
    End If
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    Set m_ObjProyecto = Constructor.getProyecto(p_IDProyecto, p_Error, m_db)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjProyecto Is Nothing Then
        p_Error = "No existe el proyecto con ID: " & p_IDProyecto
        Err.Raise 1000
    End If
    Set m_ObjEdicionUltima = m_ObjProyecto.EdicionUltima
    p_Error = m_ObjProyecto.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjEdicionUltima Is Nothing Then
        p_Error = "No existe la última edición del proyecto"
        Err.Raise 1000
    End If
    If Not IsNumeric(m_ObjEdicionUltima.Edicion) Then
        p_Error = "La última edición no tiene numeración"
        Err.Raise 1000
    End If
    m_Edicion = CStr(CLng(m_ObjEdicionUltima.Edicion + 1))
    m_FechaMax = m_ObjEdicionUltima.FechaMaxProximaPublicacionCalculado
    p_Error = m_ObjEdicionUltima.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "TbProyectosEdiciones"
    Set rcdDatosDestino = m_db.OpenRecordset(m_SQL)
    
    m_Id = DameID("TbProyectosEdiciones", "IDEdicion", m_db, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "SELECT * FROM TbProyectosEdiciones " & _
            "WHERE IDEdicion=" & p_IDEdicion & ";"
    Set rcdDatosOrigen = m_db.OpenRecordset(m_SQL)
    If rcdDatosOrigen.EOF Then
        Exit Function
    End If
    rcdDatosDestino.AddNew
        rcdDatosDestino.Fields("IDEdicion") = m_Id
        rcdDatosDestino.Fields("IDProyecto") = p_IDProyecto
        rcdDatosDestino.Fields("FechaEdicion") = CStr(m_FechaRef)
        rcdDatosDestino.Fields("Edicion") = m_Edicion
        rcdDatosDestino.Fields("Elaborado") = rcdDatosOrigen.Fields("Elaborado")
        rcdDatosDestino.Fields("Revisado") = rcdDatosOrigen.Fields("Revisado")
        rcdDatosDestino.Fields("Aprobado") = rcdDatosOrigen.Fields("Aprobado")
        rcdDatosDestino.Fields("PermitidoImprimirExcel") = rcdDatosOrigen.Fields("PermitidoImprimirExcel")
        If IsDate(m_FechaMax) Then
            rcdDatosDestino.Fields("FechaMaxProximaPublicacion") = m_FechaMax
        End If
        
    rcdDatosDestino.Update
    Set m_objEdicion = Constructor.getEdicion(m_Id, p_Error, m_db)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set CopiarEdicion = m_objEdicion
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CopiarEdicion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function CopiarPM( _
                            p_IDPM As String, _
                            p_IDRiesgo As String, _
                            Optional ByRef p_Error As String, _
                            Optional p_db As DAO.Database = Nothing _
                             ) As PM
    Dim m_ObjPM As PM
    Dim rcdDatosOrigen As DAO.Recordset
    Dim rcdDatosDestino As DAO.Recordset
    Dim m_SQL As String
    Dim m_Id As String
    Dim m_db As DAO.Database
    
    On Error GoTo errores
    
    If p_db Is Nothing Then
        Set m_db = getdb()
    Else
        Set m_db = p_db
    End If
    
    m_SQL = "TbRiesgosPlanMitigacionPpal"
    Set rcdDatosDestino = m_db.OpenRecordset(m_SQL)
    
    m_Id = DameID("TbRiesgosPlanMitigacionPpal", "IDMitigacion", m_db, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_SQL = "SELECT * FROM TbRiesgosPlanMitigacionPpal " & _
            "WHERE IDMitigacion=" & p_IDPM & ";"
    Set rcdDatosOrigen = m_db.OpenRecordset(m_SQL)
    If rcdDatosOrigen.EOF Then
        Exit Function
    End If
    rcdDatosDestino.AddNew
        rcdDatosDestino.Fields("IDMitigacion") = m_Id
        rcdDatosDestino.Fields("IDRiesgo") = p_IDRiesgo
        rcdDatosDestino.Fields("CodMitigacion") = rcdDatosOrigen.Fields("CodMitigacion")
        rcdDatosDestino.Fields("DisparadorDelPlan") = rcdDatosOrigen.Fields("DisparadorDelPlan")
        rcdDatosDestino.Fields("FechaDeActivacion") = rcdDatosOrigen.Fields("FechaDeActivacion")
        rcdDatosDestino.Fields("FechaDesactivacion") = rcdDatosOrigen.Fields("FechaDesactivacion")
        'rcdDatosDestino.Fields("Estado") = rcdDatosOrigen.Fields("Estado")
    rcdDatosDestino.Update
    Set m_ObjPM = Constructor.getPM(m_Id, p_Error, m_db)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set CopiarPM = m_ObjPM
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CopiarPM ha devuelto el error: " & Err.Description
    End If
End Function

Public Function CopiarPMAccion( _
                                p_IDPMAccion As String, _
                                p_IDPM As String, _
                                Optional ByRef p_Error As String, _
                                Optional p_db As DAO.Database = Nothing _
                                 ) As PMAccion
    Dim m_ObjPMAccion As PMAccion
    Dim rcdDatosOrigen As DAO.Recordset
    Dim rcdDatosDestino As DAO.Recordset
    Dim m_SQL As String
    Dim m_Id As String
    Dim m_db As DAO.Database
    
    On Error GoTo errores
    
    If p_db Is Nothing Then
        Set m_db = getdb()
    Else
        Set m_db = p_db
    End If
    
    m_SQL = "TbRiesgosPlanMitigacionDetalle"
    Set rcdDatosDestino = m_db.OpenRecordset(m_SQL)
    
    m_Id = DameID("TbRiesgosPlanMitigacionDetalle", "IDAccionMitigacion", m_db, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_SQL = "SELECT * FROM TbRiesgosPlanMitigacionDetalle " & _
            "WHERE IDAccionMitigacion=" & p_IDPMAccion & ";"
    Set rcdDatosOrigen = m_db.OpenRecordset(m_SQL)
    If rcdDatosOrigen.EOF Then
        Exit Function
    End If
    rcdDatosDestino.AddNew
        rcdDatosDestino.Fields("IDAccionMitigacion") = m_Id
        rcdDatosDestino.Fields("IDMitigacion") = p_IDPM
        rcdDatosDestino.Fields("CodAccion") = rcdDatosOrigen.Fields("CodAccion")
        rcdDatosDestino.Fields("Accion") = rcdDatosOrigen.Fields("Accion")
        rcdDatosDestino.Fields("ResponsableAccion") = rcdDatosOrigen.Fields("ResponsableAccion")
        rcdDatosDestino.Fields("FechaInicio") = rcdDatosOrigen.Fields("FechaInicio")
        rcdDatosDestino.Fields("FechaFinPrevista") = rcdDatosOrigen.Fields("FechaFinPrevista")
        rcdDatosDestino.Fields("FechaFinReal") = rcdDatosOrigen.Fields("FechaFinReal")
        
    rcdDatosDestino.Update
    Set m_ObjPMAccion = Constructor.getPMAccion(m_Id, p_Error, m_db)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set CopiarPMAccion = m_ObjPMAccion
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CopiarPMAccion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function CopiarPC( _
                            p_IDPC As String, _
                            p_IDRiesgo As String, _
                            Optional ByRef p_Error As String _
                             ) As PC
    Dim m_ObjPC As PC
    Dim rcdDatosOrigen As DAO.Recordset
    Dim rcdDatosDestino As DAO.Recordset
    Dim m_SQL As String
    Dim m_Id As String
    
    On Error GoTo errores
    
    
    m_SQL = "TbRiesgosPlanContingenciaPpal"
    Set rcdDatosDestino = getdb().OpenRecordset(m_SQL)
    
    m_Id = DameID("TbRiesgosPlanContingenciaPpal", "IDContingencia", , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_SQL = "SELECT * FROM TbRiesgosPlanContingenciaPpal " & _
            "WHERE IDContingencia=" & p_IDPC & ";"
    Set rcdDatosOrigen = getdb().OpenRecordset(m_SQL)
    If rcdDatosOrigen.EOF Then
        Exit Function
    End If
    rcdDatosDestino.AddNew
        rcdDatosDestino.Fields("IDContingencia") = m_Id
        rcdDatosDestino.Fields("IDRiesgo") = p_IDRiesgo
        rcdDatosDestino.Fields("CodContingencia") = rcdDatosOrigen.Fields("CodContingencia")
        rcdDatosDestino.Fields("DisparadorDelPlan") = rcdDatosOrigen.Fields("DisparadorDelPlan")
        rcdDatosDestino.Fields("FechaDeActivacion") = rcdDatosOrigen.Fields("FechaDeActivacion")
        rcdDatosDestino.Fields("FechaDesactivacion") = rcdDatosOrigen.Fields("FechaDesactivacion")
        rcdDatosDestino.Fields("Estado") = rcdDatosOrigen.Fields("Estado")
    rcdDatosDestino.Update
    Set m_ObjPC = Constructor.getPC(m_Id, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set CopiarPC = m_ObjPC
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CopiarPC ha devuelto el error: " & Err.Description
    End If
End Function

Public Function CopiarPCAccion( _
                                p_IDPCAccion As String, _
                                p_IDPC As String, _
                                Optional ByRef p_Error As String _
                                 ) As PCAccion
    Dim m_ObjPCAccion As PCAccion
    Dim rcdDatosOrigen As DAO.Recordset
    Dim rcdDatosDestino As DAO.Recordset
    Dim m_SQL As String
    Dim m_Id As String
    
    On Error GoTo errores
    
    
    m_SQL = "TbRiesgosPlanContingenciaDetalle"
    Set rcdDatosDestino = getdb().OpenRecordset(m_SQL)
    
    m_Id = DameID("TbRiesgosPlanContingenciaDetalle", "IDAccionContingencia", , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_SQL = "SELECT * FROM TbRiesgosPlanContingenciaDetalle " & _
            "WHERE IDAccionContingencia=" & p_IDPCAccion & ";"
    Set rcdDatosOrigen = getdb().OpenRecordset(m_SQL)
    If rcdDatosOrigen.EOF Then
        Exit Function
    End If
    rcdDatosDestino.AddNew
        rcdDatosDestino.Fields("IDAccionContingencia") = m_Id
        rcdDatosDestino.Fields("IDContingencia") = p_IDPC
        rcdDatosDestino.Fields("CodAccion") = rcdDatosOrigen.Fields("CodAccion")
        rcdDatosDestino.Fields("Accion") = rcdDatosOrigen.Fields("Accion")
        rcdDatosDestino.Fields("ResponsableAccion") = rcdDatosOrigen.Fields("ResponsableAccion")
        rcdDatosDestino.Fields("FechaInicio") = rcdDatosOrigen.Fields("FechaInicio")
        rcdDatosDestino.Fields("FechaFinPrevista") = rcdDatosOrigen.Fields("FechaFinPrevista")
        rcdDatosDestino.Fields("FechaFinReal") = rcdDatosOrigen.Fields("FechaFinReal")
        
    rcdDatosDestino.Update
    Set m_ObjPCAccion = Constructor.getPCAccion(m_Id, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set CopiarPCAccion = m_ObjPCAccion
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CopiarPCAccion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function CopiarRiesgo( _
                                p_IDRiesgo As String, _
                                p_IDEdicion As String, _
                                Optional p_PriorizacionManual As EnumSiNo = EnumSiNo.Sí, _
                                Optional ByRef p_Error As String, _
                                Optional p_db As DAO.Database = Nothing _
                                 ) As riesgo
    Dim m_ObjRiesgo As riesgo
    Dim m_SQL As String
    Dim fld As DAO.Field
    Dim rcdDatosOrigen As DAO.Recordset
    Dim rcdDatosDestino As DAO.Recordset
    Dim m_Id As String
    Dim m_Priorizacion As String
    Dim m_ColCamposANoCopiar As Scripting.Dictionary
    Dim m_db As DAO.Database

    On Error GoTo errores
    
    If p_db Is Nothing Then
        Set m_db = getdb()
    Else
        Set m_db = p_db
    End If
    
    m_SQL = "TbRiesgos"
    Set rcdDatosDestino = m_db.OpenRecordset(m_SQL)
    
    m_Id = DameID("TbRiesgos", "IDRiesgo", m_db, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ColCamposANoCopiar = New Scripting.Dictionary
    m_ColCamposANoCopiar.CompareMode = TextCompare
    With m_ColCamposANoCopiar
        .Add "DiasSinRespuestaCalidadAceptacion", "DiasSinRespuestaCalidadAceptacion"
        .Add "DiasSinRespuestaCalidadRetiro", "DiasSinRespuestaCalidadRetiro"
        .Add "DiasSinRespuestaCalidadRetipificacion", "DiasSinRespuestaCalidadRetipificacion"
        .Add "Priorizacion", "Priorizacion"
    End With
    m_SQL = "SELECT * FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & p_IDRiesgo & ";"
    Set rcdDatosOrigen = m_db.OpenRecordset(m_SQL)
    If rcdDatosOrigen.EOF Then
        Exit Function
    End If
    rcdDatosDestino.AddNew
    For Each fld In rcdDatosOrigen.Fields
        If m_ColCamposANoCopiar.Exists(fld.Name) Then
            GoTo siguienteCampo
        End If
        If fld.Name = "IDRiesgo" Then
            rcdDatosDestino.Fields(fld.Name).Value = m_Id
        ElseIf fld.Name = "IDEdicion" Then
            rcdDatosDestino.Fields(fld.Name).Value = p_IDEdicion
        ElseIf fld.Name = "Priorizacion" Then
            If p_PriorizacionManual <> EnumSiNo.Sí Then
                rcdDatosDestino.Fields(fld.Name).Value = rcdDatosOrigen.Fields(fld.Name).Value
            End If
        
        Else
            rcdDatosDestino.Fields(fld.Name).Value = rcdDatosOrigen.Fields(fld.Name).Value
        End If
siguienteCampo:
    Next
    rcdDatosDestino.Update
    
    Set m_ObjRiesgo = Constructor.getRiesgo(m_Id, , , p_Error, m_db)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
   
    Set CopiarRiesgo = m_ObjRiesgo
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CopiarRiesgo ha devuelto el error: " & Err.Description
    End If
End Function
Public Function SetFechaMaximaPublicacion( _
                                            p_IDProyecto As String, _
                                            Optional p_FechaSiguientePublicacion As String, _
                                            Optional ByVal p_db As DAO.Database, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    
    Dim m_Proyecto As Proyecto
    Dim m_EdicionUltimaActiva As Edicion
    Dim blnGrabar As Boolean
    Dim m_SQL As String
    Dim db As DAO.Database
    On Error GoTo errores
    
    Set m_Proyecto = Constructor.getProyecto(p_IDProyecto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Proyecto Is Nothing Then
        p_Error = "No se ha podido obtener los datos del proyecto"
        Err.Raise 1000
    End If
    If m_Proyecto.EsActivo = EnumSiNo.No Then
        Exit Function
    End If
    If Not IsDate(p_FechaSiguientePublicacion) Then
        p_FechaSiguientePublicacion = m_Proyecto.FechaMaxProximaPublicacionCalculada
        If Not IsDate(p_FechaSiguientePublicacion) Then
            Exit Function
        End If
    End If
    
    
    If IsDate(m_Proyecto.FechaMaxProximaPublicacion) Then
        If CDate(p_FechaSiguientePublicacion) <> CDate(m_Proyecto.FechaMaxProximaPublicacion) Then
            blnGrabar = True
        Else
            blnGrabar = False
        End If
    Else
        blnGrabar = True
    End If
    If blnGrabar = False Then
        SetFechaMaximaPublicacion = m_Proyecto.FechaMaxProximaPublicacion
        Exit Function
    End If
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then Err.Raise 1000
    Else
        Set db = p_db
    End If
    m_SQL = "UPDATE TbProyectos SET FechaMaxProximaPublicacion = #" & Format(p_FechaSiguientePublicacion, "mm/dd/yyyy") & "# " & _
                "WHERE IDProyecto=" & p_IDProyecto & ";"
    db.Execute m_SQL
    If Not m_Proyecto.EdicionUltima Is Nothing Then
        If Not IsDate(m_Proyecto.EdicionUltima.FechaPublicacion) Then
            m_SQL = "UPDATE TbProyectosEdiciones SET FechaMaxProximaPublicacion = #" & Format(p_FechaSiguientePublicacion, "mm/dd/yyyy") & "# " & _
                "WHERE IDEdicion=" & m_Proyecto.EdicionUltima.IDEdicion & ";"
            db.Execute m_SQL
        End If
    End If
    
   
    SetFechaMaximaPublicacion = p_FechaSiguientePublicacion
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Edicion.SetFechaMaximaPublicacion ha devuelto el error: " & Err.Description
    End If
End Function





Public Function getTextoToAStream(p_Texto As String, Optional ByRef p_Error As String) As ADODB.stream
    
    
    
    On Error GoTo errores
    Set getTextoToAStream = New ADODB.stream
    getTextoToAStream.Open
    
    getTextoToAStream.WriteText p_Texto
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getTextoToAStream ha producido el error num " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Public Function DameCabeceraHTML( _
                                p_Titulo As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_mensaje As String
    
    On Error GoTo errores
    
    
    m_mensaje = m_ObjEntorno.CabeceraHTML
    m_mensaje = Replace(m_mensaje, "#titulo", p_Titulo)
    DameCabeceraHTML = m_mensaje
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameCabeceraHTML ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function


Public Function UsuarioTienePermisoParaEscribirEnProyecto( _
                                                            Optional ByRef p_Proyecto As Proyecto, _
                                                            Optional p_Usuario As Usuario, _
                                                            Optional ByRef p_Error As String _
                                                            ) As EnumSiNo
                    
       
    
    On Error GoTo errores
    If p_Usuario Is Nothing Then
        Set p_Usuario = m_ObjUsuarioConectado
    End If
    If p_Usuario.EsAdministradorCalculado = EnumSiNo.Sí Or p_Usuario.EsCalidadCalculado = EnumSiNo.Sí Then
        UsuarioTienePermisoParaEscribirEnProyecto = EnumSiNo.Sí
        Exit Function
    End If
    
    If p_Proyecto Is Nothing Then
        p_Error = "No se ha podido determinar el Proyecto"
        Err.Raise 1000
    End If
    If InStr(1, p_Proyecto.CadenaNombreAutorizadosCalculados, p_Usuario.Nombre) = 0 Then
        UsuarioTienePermisoParaEscribirEnProyecto = EnumSiNo.No
    Else
        UsuarioTienePermisoParaEscribirEnProyecto = EnumSiNo.Sí
    End If
    
    
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioTienePermisoParaEscribirEnProyecto ha devuelto el error: " & Err.Description
    End If
End Function

Public Function UsuarioTienePermisoParaEscribir( _
                                                ByRef p_ObjObjeto As Object, _
                                                Optional p_Usuario As Usuario, _
                                                Optional ByRef p_Error As String _
                                                ) As EnumSiNo
        
    
    Dim m_UsuarioPuedeEscribirEnProyecto As EnumSiNo
    Dim m_Edicion As Edicion
    Dim m_Proyecto As Proyecto
    
    On Error GoTo errores
    If p_ObjObjeto Is Nothing Then
        p_Error = "No se conoce el proyecto"
        Err.Raise 1000
    End If
    If p_Usuario Is Nothing Then
        Set p_Usuario = m_ObjUsuarioConectado
    
    End If
    If p_Usuario.EsAdministradorCalculado = EnumSiNo.Sí Or p_Usuario.EsCalidadCalculado = EnumSiNo.Sí Then
        UsuarioTienePermisoParaEscribir = EnumSiNo.Sí
        Exit Function
    End If
    If TypeOf p_ObjObjeto Is Proyecto Then
        Set m_Proyecto = p_ObjObjeto
        
    ElseIf TypeOf p_ObjObjeto Is Edicion Then
        Set m_Proyecto = p_ObjObjeto.Edicion
    ElseIf TypeOf p_ObjObjeto Is riesgo Then
        Set m_Proyecto = p_ObjObjeto.Edicion.Proyecto
    ElseIf TypeOf p_ObjObjeto Is PM Then
        Set m_Proyecto = p_ObjObjeto.riesgo.Edicion.Proyecto
    
    ElseIf TypeOf p_ObjObjeto Is PC Then
        Set m_Proyecto = p_ObjObjeto.riesgo.Edicion.Proyecto
    ElseIf TypeOf p_ObjObjeto Is PMAccion Then
        Set m_Proyecto = p_ObjObjeto.Mitigacion.riesgo.Edicion.Proyecto
    ElseIf TypeOf p_ObjObjeto Is PCAccion Then
    
    Else
        p_Error = "Tipo no reconocido"
        Err.Raise 1000
    End If
    
    UsuarioTienePermisoParaEscribir = UsuarioTienePermisoParaEscribirEnProyecto(m_Proyecto, p_Usuario, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
        
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioTienePermisoParaEscribir ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getEsEdicionPrimeraYActivaConPermisoUsuario( _
                                                                ByRef p_Edicion As Edicion, _
                                                                Optional p_Usuario As Usuario, _
                                                                Optional ByRef p_Error As String _
                                                                ) As EnumSiNo
        
    
    Dim m_UsuarioPuedeEscribirEnProyecto As EnumSiNo
    Dim m_EsPrimeraEdicion As EnumSiNo
    Dim m_EsEdicionActiva As EnumSiNo
    
    On Error GoTo errores
    If p_Edicion Is Nothing Then
        p_Error = "No se conoce la edición"
        Err.Raise 1000
    End If
    If p_Usuario Is Nothing Then
        Set p_Usuario = m_ObjUsuarioConectado
    
    End If
    m_EsEdicionActiva = p_Edicion.EsActivo
    p_Error = p_Edicion.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_EsEdicionActiva = EnumSiNo.No Then
        getEsEdicionPrimeraYActivaConPermisoUsuario = EnumSiNo.No
        Exit Function
    End If
    m_EsPrimeraEdicion = p_Edicion.EsPrimeraEdicion
    p_Error = p_Edicion.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_EsPrimeraEdicion = EnumSiNo.No Then
        getEsEdicionPrimeraYActivaConPermisoUsuario = EnumSiNo.No
        Exit Function
    End If
    
    m_UsuarioPuedeEscribirEnProyecto = p_Edicion.Proyecto.UsuarioAutorizado
    p_Error = p_Edicion.Proyecto.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_UsuarioPuedeEscribirEnProyecto = EnumSiNo.No Then
        getEsEdicionPrimeraYActivaConPermisoUsuario = EnumSiNo.No
        Exit Function
    End If
    getEsEdicionPrimeraYActivaConPermisoUsuario = EnumSiNo.Sí
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEsEdicionPrimeraYActivaConPermisoUsuario ha devuelto el error: " & Err.Description
    End If
End Function
Public Sub Avance(ByRef p_Linea As Variant)
    Dim frm As Form
    On Error Resume Next
    
    Set frm = Screen.ActiveForm
    If Not frm Is Nothing Then
        If Not frm.Controls("lblEstado") Is Nothing Then
            frm.Controls("lblEstado").Visible = True
            frm.Controls("lblEstado").Caption = p_Linea
        End If
    End If
    
    If FormularioAbierto("frmSplash") Then
        With Forms("frmSplash")
            Dim totalPasos As Long
            Dim anchoMaximo As Long, nuevoAncho As Long
            
            Dim tempEntorno As New Entorno
            totalPasos = tempEntorno.ColItems.Count
            Set tempEntorno = Nothing
            
            s_contadorPasos = s_contadorPasos + 1
            
            anchoMaximo = .lblProgresoFondo.Width
            
            If totalPasos > 0 Then
                ' Cálculo proporcional
                nuevoAncho = (s_contadorPasos / totalPasos) * anchoMaximo
                
                ' SALVAGUARDA: Asegurarse de que el nuevo ancho no supere el máximo.
                If nuevoAncho > anchoMaximo Then
                    nuevoAncho = anchoMaximo
                End If
            End If
            
            .lblProgresoBarra.Width = nuevoAncho
        End With
    End If
    
    VBA.DoEvents
End Sub
Public Function AvanceCerrar() As String
    
    
    
    
    On Error Resume Next
    Application.Screen.ActiveForm.Controls("lblEstado").Visible = False
    
    Exit Function
    
End Function

Public Function EjecutarShell( _
                                strComando As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    
    Dim ManejadorProceso As Long
    Dim IDProceso As Long
    Dim lpExitCode As Long
    On Error GoTo errores
    
    If strComando = "" Then
        p_Error = "No se ha indicado el comando"
        Err.Raise 1000
    End If
    IDProceso = Shell(strComando, vbHide)
    ManejadorProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, IDProceso)
    ' Mientras lp_ExitCode = STATUS_PENDING, se ejecuta el do
    Do
        Call GetExitCodeProcess(ManejadorProceso, lpExitCode)
        DoEvents
    Loop While lpExitCode = STATUS_PENDING
    Call CloseHandle(ManejadorProceso)
    
    
    
    EjecutarShell = "OK"
    Exit Function
errores:
    
    If Err.Number <> 1000 Then
        p_Error = "El método EjecutarShell ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function



Public Function DAMESSIDPorArchivo(Optional ByRef p_Error As String) As String
    
    
    Dim fichero As Scripting.TextStream
    
    Dim m_ArchivoSSID As String
    
    
    On Error GoTo errores
    If m_ObjEntorno Is Nothing Then
        Exit Function
    End If
    m_ArchivoSSID = m_ObjEntorno.URLArchivoSSID
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not FSO.FileExists(m_ArchivoSSID) Then
        Exit Function
    End If
    Set fichero = FSO.OpenTextFile(m_ArchivoSSID)
    DAMESSIDPorArchivo = Trim(fichero.ReadLine)
    fichero.Close
    

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DAMESSIDPorArchivo ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function



Public Function EliminarPublicacionesLog( _
                                        p_IDEdicion As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
     
    
    Dim m_SQL As String
    
    On Error GoTo errores
    
    If p_IDEdicion = "" Then
        Exit Function
    End If
    m_SQL = "DELETE * FROM TbLogPublicaciones " & _
            "WHERE IDEdicion=" & p_IDEdicion & ";"
    getdb().Execute m_SQL
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EliminarPublicacionesLog ha producido el error : " & vbNewLine & Err.Description
    End If
    
    
End Function

Public Function RegistrarEnAGEDO( _
                                ByRef p_ObjEdicionAPublicar As Edicion, _
                                p_URLInicial As String, _
                                Optional ByRef p_Error As String _
                                ) As Documento
    
    
    Dim m_ObjDocumento As Documento
    Dim m_NombreArchivo As String
    Dim m_Codigo As String
    Dim m_Edicion As String
    Dim m_FechaPublicacion As String
    Dim m_Proyecto As Proyecto
    Dim m_FechaRef As Date
    
    On Error GoTo errores
    
    Avance "Publicación...Registrando Documento En AGEDO..."
    Set m_Proyecto = Constructor.getProyecto(p_ObjEdicionAPublicar.IDProyecto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    
    m_Codigo = m_Proyecto.CodigoDocumento
    
    If m_Codigo = "" Then
        p_Error = "NO se ha podido obtener el código del documento"
        Err.Raise 1000
    End If
    m_NombreArchivo = FSO.GetFileName(p_URLInicial)
    m_Edicion = p_ObjEdicionAPublicar.Edicion
    m_FechaPublicacion = p_ObjEdicionAPublicar.FechaPublicacion
    
    Set m_ObjDocumento = New Documento
    With m_ObjDocumento
        .Tipo = "IS"
        .Area = "E"
        .Archivo = "Electrónico"
        .Versionable = "Sí"
        .Titulo = "Informe de seguimiento de Riesgos"
        .FechaAlta = CStr(m_FechaRef)
        .FEdicion = CStr(m_FechaRef)
        .FENTRADAENVIGOR = CStr(m_FechaRef)
        .Clasificacion = "SINCLAS"
        .Observaciones = "Alta Automática desde gestión de riesgos"
        .CadenaNombreArchivos = m_NombreArchivo
        .URLCarpetaAGEDO = m_ObjEntorno.URLDirectorioTemporalAGEDO
        .URLCarpetaArchivosLocales = FSO.GetParentFolderName(p_URLInicial)
        .codigo = m_Codigo
        .Edicion = m_Edicion
    End With
    With m_ObjDocumento
        Set .EdicionObj = p_ObjEdicionAPublicar
        Set .Proyecto = m_Proyecto
        .Registrar p_URLInicial, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
    End With
    
    Set RegistrarEnAGEDO = m_ObjDocumento
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistrarEnAGEDO ha producido el error : " & vbNewLine & Err.Description
    End If
End Function

Public Function getObjetoAnterior( _
                                    p_Obj As Object, _
                                    Optional ByRef p_Error As String _
                                    ) As Object
    
    
    
    On Error GoTo errores
    
    If TypeOf p_Obj Is Edicion Then
        Set getObjetoAnterior = p_Obj.EdicionAnterior
        p_Error = p_Obj.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    ElseIf TypeOf p_Obj Is riesgo Then
        Set getObjetoAnterior = p_Obj.RiesgoEdicionAnterior
        p_Error = p_Obj.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    ElseIf TypeOf p_Obj Is PM Then
        Set getObjetoAnterior = p_Obj.PMEdicionAnterior
        p_Error = p_Obj.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    ElseIf TypeOf p_Obj Is PC Then
        Set getObjetoAnterior = p_Obj.PCEdicionAnterior
        p_Error = p_Obj.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    ElseIf TypeOf p_Obj Is PMAccion Then
        Set getObjetoAnterior = p_Obj.PMAccionEdicionAnterior
        p_Error = p_Obj.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    ElseIf TypeOf p_Obj Is PCAccion Then
        Set getObjetoAnterior = p_Obj.PCAccionEdicionAnterior
        p_Error = p_Obj.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getObjetoAnterior ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function AbrirAyuda( _
                            Optional ByRef p_Error As String _
                            ) As String

    
    Dim m_NombreFormulario As String
    Dim m_NombreArchivo As String
    Dim m_URL As String
    On Error Resume Next
    If Application.Screen.ActiveForm Is Nothing Then
        If Err.Number <> 0 Then
            Exit Function
        End If
        Exit Function
    End If
    If Err.Number <> 0 Then
        Exit Function
    End If
    Err.Clear
    On Error GoTo errores
    m_NombreFormulario = Application.Screen.ActiveForm.Name
    m_NombreArchivo = m_NombreFormulario & ".pdf"
    m_URL = m_ObjEntorno.URLDirectorioAyuda & m_NombreArchivo
    
    
    AbrirEnLocal m_URL, Application.Screen.ActiveForm.hWnd, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
   
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AbrirAyuda ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function RellenarListaImagenes( _
                                        ByRef p_ListImages As Object, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    Dim m_VarItem As Variant
    Dim m_URLIcono As String
    Dim m_Col As Scripting.Dictionary
    Dim i As Long
    
    On Error GoTo errores
    
    Set m_Col = m_ObjEntorno.ColImagenes
    
    i = 1
    For Each varItem In m_Col
        m_URLIcono = CStr(varItem)
        'Debug.Print m_URLIcono
        On Error Resume Next
        p_ListImages.ListImages.Add i, FSO.GetFileName(m_URLIcono), LoadPicture(m_URLIcono)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo errores
            p_ListImages.ListImages.Remove i
            p_ListImages.ListImages.Add i, FSO.GetFileName(m_URLIcono), LoadPicture(m_URLIcono)
            
        End If
        i = i + 1
    Next
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarListaImagenes ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function EstablecerControlesPublicabilidad( _
                                                    p_Form As Form, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
            
    
    
    
    Dim lblCom As Label
    
    On Error GoTo errores
    
    Set lblCom = p_Form.lbl3Complementado
    
    lblCom.Visible = True
    With m_ObjEdicionActiva
        If .EsActivo = EnumSiNo.No Then
            lblCom.ForeColor = m_ColorComplementado
            lblCom.Caption = "No Aplica"
            Exit Function
        End If
        If m_URLInforme = "" Then
            lblCom.ForeColor = m_ColorNoComplementado
            lblCom.Caption = "No Generado"
            
            Exit Function
        End If
        If .Publicable = EnumSiNo.Sí Then
            lblCom.ForeColor = m_ColorComplementado
            lblCom.Caption = "Publicable"
        Else
            lblCom.ForeColor = m_ColorNoComplementado
            lblCom.Caption = "No Publicable"
        End If
        
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo EstablecerControlesPublicabilidad datos ha producido el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function EstablecerControlesEvidenciaUTE( _
                                                    p_Form As Form, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
            
    
    
    Dim lblReq As Label
    Dim lblCom As Label
    
    On Error GoTo errores
    
    Set lblReq = p_Form.lbl1Requerido
    Set lblCom = p_Form.lbl1Complementado
    With m_ObjEdicionActiva
        If .Proyecto.EnUTE = "Sí" Then
            lblReq.Visible = True
            lblReq.Caption = "Requerido"
            lblCom.Visible = True
            If .EsActivo = EnumSiNo.No Then
                lblCom.ForeColor = m_ColorComplementado
                lblCom.Caption = "No Aplica"
                Exit Function
            End If
            If Not m_ObjAnexoEvicenciaUTE Is Nothing Then
                lblCom.ForeColor = m_ColorComplementado
                lblCom.Caption = "Complementado"
            Else
                lblCom.ForeColor = m_ColorNoComplementado
                lblCom.Caption = "No Complementado"
            End If
            
        Else
            lblReq.Visible = True
            lblReq.Caption = "No Requerido"
            lblCom.Visible = False
        End If
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo EstablecerControlesEvidenciaUTE datos ha producido el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function EstablecerControlesEvidenciaSuministradores( _
                                                                p_Form As Form, _
                                                                Optional ByRef p_Error As String _
                                                                ) As String
            
    
    
    Dim lblReq As Label
    Dim lblCom As Label
    On Error GoTo errores
    
    Set lblReq = p_Form.lbl2Requerido
    Set lblCom = p_Form.lbl2Complementado
    
    With m_ObjEdicionActiva
        Set m_ObjEdicionActiva.ColSuministradores = Nothing
        If .EvidenciasSuministradoresRequeridas = EnumSiNo.Sí Then
            lblReq.Visible = True
            lblReq.Caption = "Requerido"
            lblCom.Visible = True
            If .EsActivo = EnumSiNo.No Then
                lblCom.ForeColor = m_ColorComplementado
                lblCom.Caption = "No Aplica"
                Exit Function
            End If
            If .EvidenciasSuministradoresCompletadas = EnumSiNo.Sí Then
                lblCom.ForeColor = m_ColorComplementado
                lblCom.Caption = "Complementado"
            Else
                lblCom.ForeColor = m_ColorNoComplementado
                lblCom.Caption = "No Complementado"
            End If
        Else
            lblReq.Visible = True
            lblReq.Caption = "No Requerido"
            lblCom.Visible = False
        End If
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo EstablecerControlesEvidenciaSuministradores datos ha producido el error: " & vbNewLine & Err.Description
    End If
End Function


Public Function EstablecerControlesPropuesta( _
                                                p_Form As Form, _
                                                Optional ByRef p_Error As String _
                                                ) As String
            
    
    Dim lblCom As Label
    Dim btn As Object
    On Error GoTo errores
    
    
    Set lblCom = p_Form.lbl4Complementado
    Set btn = p_Form.ComandoVerRechazo
    
    
    On Error GoTo errores
    
    lblCom.Visible = True
    btn.Visible = False
    With m_ObjEdicionActiva
        If .EsActivo = EnumSiNo.No Then
            lblCom.ForeColor = m_ColorComplementado
            lblCom.Caption = "No Aplica"
            Exit Function
        End If
        If IsDate(.FechaPreparadaParaPublicar) Then
            lblCom.ForeColor = m_ColorComplementado
            lblCom.Caption = Format(.FechaPreparadaParaPublicar, "dd/mm/yyyy")
        Else
            If .PublicacionRechazada = EnumSiNo.Sí Then
                lblCom.Caption = "Rechazada"
                lblCom.ForeColor = m_ColorNoComplementado
                btn.Visible = True
                btn.Enabled = True
            Else
                lblCom.ForeColor = m_ColorNoComplementado
                lblCom.Caption = "No Complementado"
            End If
            
        End If
    
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo EstablecerControlesPropuesta datos ha producido el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function EstablecerEtiquetas( _
                                        p_Form As Form, _
                                        Optional ByRef p_Error As String _
                                        ) As String
            
    
    
    
    On Error GoTo errores
    EstablecerControlesEvidenciaUTE p_Form, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerControlesEvidenciaSuministradores p_Form, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerControlesPublicabilidad p_Form, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerControlesPropuesta p_Form, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    
      
        
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo EstablecerEtiquetas datos ha producido el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function EstablecerComboFamilia(cmb As ComboBox, Optional ByRef p_Error As String) As String
    
    
    Dim m_Familia As Variant
    
    
    On Error GoTo errores
    
    
    cmb.RowSource = ""
    If Not m_ObjEntorno.ColRiesgosBibliotecasFamilias Is Nothing Then
        
        For Each m_Familia In m_ObjEntorno.ColRiesgosBibliotecasFamilias
            cmb.AddItem m_Familia
            
        Next
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerComboFamilia ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function



Public Function FormatoCorrectoURL( _
                                    p_URL As String _
                                    ) As EnumSiNo
   
    'No borrar es para la consulta 1,12
    Dim dato
    Dim m_Trozo As String
    
    On Error Resume Next
    If InStr(1, p_URL, ":\") <> 0 Then
        dato = Split(p_URL, ":\")
        m_Trozo = dato(1)
        If InStr(1, m_Trozo, "\\") <> 0 Then
            FormatoCorrectoURL = EnumSiNo.No
            Exit Function
        End If
        FormatoCorrectoURL = EnumSiNo.Sí
        Exit Function
    End If
    If InStr(1, p_URL, "\\") = 0 Then
        FormatoCorrectoURL = EnumSiNo.No
        Exit Function
    End If
    dato = Split(p_URL, "\\")
    If UBound(dato) > 1 Then
        FormatoCorrectoURL = EnumSiNo.No
        Exit Function
    End If
    FormatoCorrectoURL = EnumSiNo.Sí
    
    
    Exit Function

    
    
End Function

Public Function getObjetoSeleccionadoDeRiesgos( _
                                                ByVal p_Key As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Object
    
    Dim dato As Variant
    Dim m_Id As String
    Dim m_PrimerTrozo As String
    On Error GoTo errores

'   EDICION1 EDICION|IDEDICION
'       RIESGO1  RIESGO|IDRIESGO1
'           PM1     PM|IDRIESGO1|IDPM1
'               ACCIÓN1 PMACCION|IDPM1|IDACCION1
'               ACCIÓN2 PMACCION|IDPM1|IDACCION2
'           PM2     PM|IDRIESGO1|IDPM2
'               ACCIÓN1 PMACCION|IDPM2|IDACCION1
'       RIESGO2  RIESGO|IDRIESGO2
    If p_Key = "" Then
        Exit Function
    End If
    If InStr(1, p_Key, "|") = 0 Then
        Exit Function
    End If
    dato = Split(p_Key, "|")
    m_PrimerTrozo = dato(0)
    If m_PrimerTrozo = "EDICION" Then
        m_Id = dato(1)
        If m_ObjEdicionActiva.IDEdicion <> m_Id Then
            Set m_ObjEdicionActiva = Constructor.getEdicion(m_Id, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        Set getObjetoSeleccionadoDeRiesgos = m_ObjEdicionActiva
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If m_PrimerTrozo = "RIESGO" Then
        m_Id = dato(1)
        Set getObjetoSeleccionadoDeRiesgos = Constructor.getRiesgo(m_Id, , , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If m_PrimerTrozo = "PM" Then
        m_Id = dato(2)
        Set getObjetoSeleccionadoDeRiesgos = Constructor.getPM(m_Id, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If m_PrimerTrozo = "PC" Then
        m_Id = dato(2)
        Set getObjetoSeleccionadoDeRiesgos = Constructor.getPC(m_Id, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If m_PrimerTrozo = "PMACCION" Then
        'PMACCION|IDPM2|IDACCION1
        m_Id = dato(2)
        Set getObjetoSeleccionadoDeRiesgos = Constructor.getPMAccion(m_Id, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If m_PrimerTrozo = "PCACCION" Then
        m_Id = dato(2)
        Set getObjetoSeleccionadoDeRiesgos = Constructor.getPCAccion(m_Id, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Arbol.getObjetoSeleccionadoDeRiesgos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function CambiaIconoElementoArbolDeRiesgos(p_Arbol As MSComctlLib.TreeView, p_Error As String) As String
    
    
    Dim m_ObjetoSeleccionado As Object
    Dim m_NodoSeleccionado As MSComctlLib.Node
    Dim m_NodoACambiarIcono As MSComctlLib.Node
    Dim m_NombreIcono As String
    
    
    
    'si estamos en un nodo riesgo solo cambia el riesgo
    'si estamos en un nodo plan y este cambia de icono, vamos al riesgo a ver que pasa
    'si estamos en un nodo accion y cambia, vamos al plan y lo cambiamos y si este cambia, vamos al riesgo
    
    Set m_NodoSeleccionado = p_Arbol.SelectedItem
    If m_NodoSeleccionado Is Nothing Then
        Exit Function
    End If
    If m_ObjEdicionActiva Is Nothing Then
        p_Error = "No hay una edición activa"
        Err.Raise 1000
    End If
    
    Set m_ObjetoSeleccionado = getObjetoSeleccionadoDeRiesgos(m_NodoSeleccionado.Key, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjetoSeleccionado Is Nothing Then
        Exit Function
    End If
    If TypeOf m_ObjetoSeleccionado Is riesgo Then
        m_NombreIcono = getNombreIconoRiesgoArbolDeRiesgos(m_ObjetoSeleccionado, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_NombreIcono <> m_NodoSeleccionado.Image Then
            m_NodoSeleccionado.Image = m_NombreIcono
            If m_ObjEdicionActiva.colRiesgosNoRetirados.Exists(m_ObjetoSeleccionado.IDRiesgo) Then
                m_ObjEdicionActiva.colRiesgosNoRetirados.Remove (m_ObjetoSeleccionado.IDRiesgo)
                m_ObjEdicionActiva.colRiesgosNoRetirados.Add m_ObjetoSeleccionado.IDRiesgo, m_ObjetoSeleccionado
            End If
            CachePublicabilidad_RecalcularEdicionYResetear m_ObjEdicionActiva, , p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        
        Exit Function
    End If
    If TypeOf m_ObjetoSeleccionado Is PM Or TypeOf m_ObjetoSeleccionado Is PC Then
        m_NombreIcono = getNombreIconoPlanArbolDeRiesgos(m_ObjetoSeleccionado, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_NombreIcono <> m_NodoSeleccionado.Image Then
            m_NodoSeleccionado.Image = m_NombreIcono
            Set m_ObjetoSeleccionado = m_ObjetoSeleccionado.riesgo
            Set m_NodoACambiarIcono = m_NodoSeleccionado.Parent
            m_NombreIcono = getNombreIconoRiesgoArbolDeRiesgos(m_ObjetoSeleccionado, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            m_NodoSeleccionado.Image = m_NombreIcono
            
            If m_ObjEdicionActiva.colRiesgosNoRetirados.Exists(m_ObjetoSeleccionado.IDRiesgo) Then
                m_ObjEdicionActiva.colRiesgosNoRetirados.Remove (m_ObjetoSeleccionado.IDRiesgo)
                m_ObjEdicionActiva.colRiesgosNoRetirados.Add m_ObjetoSeleccionado.IDRiesgo, m_ObjetoSeleccionado
            End If
            CachePublicabilidad_RecalcularEdicionYResetear m_ObjEdicionActiva, , p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Function
        End If
    End If
    If TypeOf m_ObjetoSeleccionado Is PMAccion Or TypeOf m_ObjetoSeleccionado Is PCAccion Then
        m_NombreIcono = getNombreIconoAccionArbolDeRiesgos(m_ObjetoSeleccionado, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
            m_NodoSeleccionado.Image = m_NombreIcono
            If TypeOf m_ObjetoSeleccionado Is PMAccion Then
                Set m_ObjetoSeleccionado = m_ObjetoSeleccionado.Mitigacion
            Else
                Set m_ObjetoSeleccionado = m_ObjetoSeleccionado.Contingencia
            End If
            Set m_NodoACambiarIcono = m_NodoSeleccionado.Parent
            m_NombreIcono = getNombreIconoPlanArbolDeRiesgos(m_ObjetoSeleccionado, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
            m_NodoACambiarIcono.Image = m_NombreIcono
            Set m_ObjetoSeleccionado = m_ObjetoSeleccionado.riesgo
            Set m_NodoACambiarIcono = m_NodoACambiarIcono.Parent
            m_NombreIcono = getNombreIconoRiesgoArbolDeRiesgos(m_ObjetoSeleccionado, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_NodoACambiarIcono.Image <> m_NombreIcono Then
                m_NodoACambiarIcono.Image = m_NombreIcono
            End If
            If m_ObjEdicionActiva.colRiesgosNoRetirados.Exists(m_ObjetoSeleccionado.IDRiesgo) Then
                m_ObjEdicionActiva.colRiesgosNoRetirados.Remove (m_ObjetoSeleccionado.IDRiesgo)
                m_ObjEdicionActiva.colRiesgosNoRetirados.Add m_ObjetoSeleccionado.IDRiesgo, m_ObjetoSeleccionado
            End If
            CachePublicabilidad_RecalcularEdicionYResetear m_ObjEdicionActiva, , p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
            
        
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo CambiaIconoElementoArbolDeRiesgos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Public Function getNombreIconoRiesgoArbolDeRiesgos( _
                                                    m_Riesgo As riesgo, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    
    
    On Error GoTo errores
    With m_Riesgo
        
        If .CalcularHayErrorEnRiesgo(p_Error:=p_Error) = "Sí" Then
           getNombreIconoRiesgoArbolDeRiesgos = FSO.GetFileName(m_ObjEntorno.URLIconoItemInCompleto32)
        Else
           getNombreIconoRiesgoArbolDeRiesgos = FSO.GetFileName(m_ObjEntorno.URLIconoItemCompleto32)
        End If
        
    End With
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getNombreIconoRiesgoArbolDeRiesgos ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getNombreIconoAccionArbolDeRiesgos(m_Accion As Object, Optional ByRef p_Error As String) As String
    
    Dim m_PMA As PMAccion
    On Error GoTo errores
    
    
    If m_Accion.NecesitaReplanificacion = EnumSiNo.Sí Then
        getNombreIconoAccionArbolDeRiesgos = FSO.GetFileName(m_ObjEntorno.URLIconoItemInCompleto32)
    Else
        getNombreIconoAccionArbolDeRiesgos = FSO.GetFileName(m_ObjEntorno.URLIconoItemCompleto32)
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getNombreIconoAccionArbolDeRiesgos ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getNombreIconoPlanArbolDeRiesgos(m_Plan As Object, Optional ByRef p_Error As String) As String
    
    Dim m_EstadoRiesgo As EnumRiesgoEstado
    Dim m_RiesgoAltoOMuyAlto As EnumSiNo
    Dim m_Riesgo As riesgo
    Dim m_PM As PM
    On Error GoTo errores
    
    If Not TypeOf m_Plan Is PM And Not TypeOf m_Plan Is PC Then
        Exit Function
    End If
    
    
    If m_Plan.EsActivo = EnumSiNo.No Then
        getNombreIconoPlanArbolDeRiesgos = FSO.GetFileName(m_ObjEntorno.URLIconoItemCompleto32)
        Exit Function
    End If

    If m_Plan.TieneAcciones = EnumSiNo.No Then
        getNombreIconoPlanArbolDeRiesgos = FSO.GetFileName(m_ObjEntorno.URLIconoItemInCompleto32)
        Exit Function
    End If
    If m_Plan.AlgunaAccionParaReplanificar = EnumSiNo.Sí Then
        getNombreIconoPlanArbolDeRiesgos = FSO.GetFileName(m_ObjEntorno.URLIconoItemInCompleto32)
        Exit Function
    End If
    Set m_Riesgo = m_Plan.riesgo
    
    With m_Riesgo
        
        m_EstadoRiesgo = .ESTADOCalculado
        m_RiesgoAltoOMuyAlto = .RiesgoAltoOMuyAlto
    End With
    If m_RiesgoAltoOMuyAlto = EnumSiNo.Sí Then
        If TypeOf m_Plan Is PM Then
            If m_Plan.ESTADOCalculado <> EnumPlanEstado.Activo Then
                getNombreIconoPlanArbolDeRiesgos = FSO.GetFileName(m_ObjEntorno.URLIconoItemInCompleto32)
                Exit Function
            End If
        
            
        End If
        
    End If
    If m_EstadoRiesgo = EnumRiesgoEstado.Materializado And TypeOf m_Plan Is PC Then
        If m_Plan.ESTADOCalculado <> EnumPlanEstado.Activo Then
            getNombreIconoPlanArbolDeRiesgos = FSO.GetFileName(m_ObjEntorno.URLIconoItemCompleto32)
            Exit Function
        End If
    End If
    If m_Plan.NecesitaReplanificacion = EnumSiNo.No Then
        getNombreIconoPlanArbolDeRiesgos = FSO.GetFileName(m_ObjEntorno.URLIconoItemCompleto32)
    Else
        getNombreIconoPlanArbolDeRiesgos = FSO.GetFileName(m_ObjEntorno.URLIconoItemInCompleto32)
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getNombreIconoPlanArbolDeRiesgos ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getColorNodo( _
                                p_Riesgo As riesgo, _
                                Optional ByRef p_Error As String _
                                ) As Long
    
    
    On Error GoTo errores
    
    With p_Riesgo
        
        If .ColorIcono = "Negro" Then
            getColorNodo = vbBlack
        ElseIf .ColorIcono = "Rojo" Then
            getColorNodo = vbRed
        Else
            getColorNodo = vbBlack
        End If
        
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getColorNodo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function PriorizacionUsadaEnRiesgo( _
                                            p_Priorizacion As String, _
                                            colRiesgos As Scripting.Dictionary, _
                                            Optional ByRef p_Error As String _
                                            ) As riesgo
    
    Dim m_Id As Variant
    Dim m_Riesgo As riesgo
    On Error GoTo errores
    If Not IsNumeric(p_Priorizacion) Then
        Exit Function
    End If
    If colRiesgos Is Nothing Then
        Exit Function
    End If
    For Each m_Id In colRiesgos
        Set m_Riesgo = colRiesgos(m_Id)
        If p_Priorizacion = m_Riesgo.Priorizacion Then
            Set PriorizacionUsadaEnRiesgo = m_Riesgo
            Exit Function
        End If
        Set m_Riesgo = Nothing
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PriorizacionUsadaEnRiesgo ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function


Public Function getRiesgosEstadoCalculadoTexto( _
                                                p_Estado As EnumRiesgoEstado, _
                                                Optional ByRef p_Error As String _
                                                ) As String

    
    On Error GoTo errores

    

    If p_Estado = EnumRiesgoEstado.Aceptado Then
        getRiesgosEstadoCalculadoTexto = "Aceptado"
    ElseIf p_Estado = EnumRiesgoEstado.AceptadoRechazado Then
        getRiesgosEstadoCalculadoTexto = "AceptadoRechazado"
    ElseIf p_Estado = EnumRiesgoEstado.AceptadoSinJustificar Then
        getRiesgosEstadoCalculadoTexto = "AceptadoSinJustificar"
    ElseIf p_Estado = EnumRiesgoEstado.AceptadoSinVisar Then
        getRiesgosEstadoCalculadoTexto = "AceptadoSinVisar"
    ElseIf p_Estado = EnumRiesgoEstado.Activo Then
        getRiesgosEstadoCalculadoTexto = "Activo"
    ElseIf p_Estado = EnumRiesgoEstado.Cerrado Then
        getRiesgosEstadoCalculadoTexto = "Cerrado"
    ElseIf p_Estado = EnumRiesgoEstado.Detectado Then
        getRiesgosEstadoCalculadoTexto = "Detectado"
    ElseIf p_Estado = EnumRiesgoEstado.Incompleto Then
        getRiesgosEstadoCalculadoTexto = "Incompleto"
    ElseIf p_Estado = EnumRiesgoEstado.Materializado Then
        getRiesgosEstadoCalculadoTexto = "Materializado"
    ElseIf p_Estado = EnumRiesgoEstado.Planificado Then
        getRiesgosEstadoCalculadoTexto = "Planificado"
    ElseIf p_Estado = EnumRiesgoEstado.Retirado Then
        getRiesgosEstadoCalculadoTexto = "Retirado"
    ElseIf p_Estado = EnumRiesgoEstado.RetiradoRechazado Then
        getRiesgosEstadoCalculadoTexto = "RetiradoRechazado"
    ElseIf p_Estado = EnumRiesgoEstado.RetiradoSinJustificar Then
        getRiesgosEstadoCalculadoTexto = "RetiradoSinJustificar"
    ElseIf p_Estado = EnumRiesgoEstado.RetiradoSinVisar Then
        getRiesgosEstadoCalculadoTexto = "RetiradoSinVisar"
    
    End If
    



    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getRiesgosEstadoCalculadoTexto ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function



Public Function RiesgoEnAceptacion( _
                                    p_EstadoRiesgo As EnumRiesgoEstado, _
                                    Optional ByRef p_Error As String) As EnumSiNo
    
    
    On Error GoTo errores
    
    
    If p_EstadoRiesgo = EnumRiesgoEstado.Aceptado Or _
        p_EstadoRiesgo = EnumRiesgoEstado.AceptadoSinJustificar Or _
        p_EstadoRiesgo = EnumRiesgoEstado.AceptadoSinVisar Or _
        p_EstadoRiesgo = EnumRiesgoEstado.AceptadoRechazado Then
        RiesgoEnAceptacion = EnumSiNo.Sí
    Else
        RiesgoEnAceptacion = EnumSiNo.No
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgoEnAceptacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function RiesgoEnRetirada( _
                                    p_EstadoRiesgo As EnumRiesgoEstado, _
                                    Optional ByRef p_Error As String) As EnumSiNo
    
   
    On Error GoTo errores
    
    
    
   
    If p_EstadoRiesgo = EnumRiesgoEstado.Retirado Or _
        p_EstadoRiesgo = EnumRiesgoEstado.RetiradoSinJustificar Or _
        p_EstadoRiesgo = EnumRiesgoEstado.RetiradoSinVisar Or _
        p_EstadoRiesgo = EnumRiesgoEstado.RetiradoRechazado Then
        RiesgoEnRetirada = EnumSiNo.Sí
    Else
        RiesgoEnRetirada = EnumSiNo.No
    End If
    
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgoEnRetirada ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getEstadoRiesgo( _
                                    p_EstadoRiesgo As String, _
                                    Optional ByRef p_Error As String) As EnumRiesgoEstado
    
   
    On Error GoTo errores
    
    If p_EstadoRiesgo = "Aceptado" Then
        getEstadoRiesgo = EnumRiesgoEstado.Aceptado
    ElseIf p_EstadoRiesgo = "AceptadoRechazado" Then
        getEstadoRiesgo = EnumRiesgoEstado.AceptadoRechazado
    ElseIf p_EstadoRiesgo = "AceptadoSinJustificar" Then
        getEstadoRiesgo = EnumRiesgoEstado.AceptadoSinJustificar
    ElseIf p_EstadoRiesgo = "AceptadoSinVisar" Then
        getEstadoRiesgo = EnumRiesgoEstado.AceptadoSinVisar
    ElseIf p_EstadoRiesgo = "Activo" Then
        getEstadoRiesgo = EnumRiesgoEstado.Activo
    ElseIf p_EstadoRiesgo = "Cerrado" Then
        getEstadoRiesgo = EnumRiesgoEstado.Cerrado
    ElseIf p_EstadoRiesgo = "Detectado" Then
        getEstadoRiesgo = EnumRiesgoEstado.Detectado
    ElseIf p_EstadoRiesgo = "Incompleto" Then
        getEstadoRiesgo = EnumRiesgoEstado.Incompleto
    ElseIf p_EstadoRiesgo = "Materializado" Then
        getEstadoRiesgo = EnumRiesgoEstado.Materializado
    ElseIf p_EstadoRiesgo = "Planificado" Then
        getEstadoRiesgo = EnumRiesgoEstado.Planificado
    ElseIf p_EstadoRiesgo = "Retirado" Then
        getEstadoRiesgo = EnumRiesgoEstado.Retirado
    ElseIf p_EstadoRiesgo = "RetiradoRechazado" Then
        getEstadoRiesgo = EnumRiesgoEstado.RetiradoRechazado
    ElseIf p_EstadoRiesgo = "RetiradoSinJustificar" Then
        getEstadoRiesgo = EnumRiesgoEstado.RetiradoSinJustificar
    ElseIf p_EstadoRiesgo = "RetiradoSinVisar" Then
        getEstadoRiesgo = EnumRiesgoEstado.RetiradoSinVisar
    End If
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEstadoRiesgo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getEstadoRiesgoTexto( _
                                    p_EstadoRiesgo As EnumRiesgoEstado, _
                                    Optional ByRef p_Error As String) As String
    
   
    On Error GoTo errores
    
    If p_EstadoRiesgo = EnumRiesgoEstado.Aceptado Then
        getEstadoRiesgoTexto = "Aceptado"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.AceptadoRechazado Then
        getEstadoRiesgoTexto = "AceptadoRechazado"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.AceptadoSinJustificar Then
        getEstadoRiesgoTexto = "AceptadoSinJustificar"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.AceptadoSinVisar Then
        getEstadoRiesgoTexto = "AceptadoSinVisar"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Activo Then
        getEstadoRiesgoTexto = "Activo"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Cerrado Then
        getEstadoRiesgoTexto = "Cerrado"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Detectado Then
        getEstadoRiesgoTexto = "Detectado"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Incompleto Then
        getEstadoRiesgoTexto = "Incompleto"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Materializado Then
        getEstadoRiesgoTexto = "Materializado"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Planificado Then
        getEstadoRiesgoTexto = "Planificado"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Retirado Then
        getEstadoRiesgoTexto = "Retirado"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.RetiradoRechazado Then
        getEstadoRiesgoTexto = "RetiradoRechazado"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.RetiradoSinJustificar Then
        getEstadoRiesgoTexto = "RetiradoSinJustificar"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.RetiradoSinVisar Then
        getEstadoRiesgoTexto = "RetiradoSinVisar"
    End If
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEstadoRiesgoTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Public Function getEstadoRiesgoTextoEtiqueta( _
                                            p_EstadoRiesgo As EnumRiesgoEstado, _
                                            Optional ByRef p_Error As String) As String
    
   
    On Error GoTo errores
    
    If p_EstadoRiesgo = EnumRiesgoEstado.Aceptado Then
        getEstadoRiesgoTextoEtiqueta = "Riesgo Aceptado"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.AceptadoRechazado Then
        getEstadoRiesgoTextoEtiqueta = "Riesgo Aceptado Rechazado"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.AceptadoSinJustificar Then
        getEstadoRiesgoTextoEtiqueta = "Riesgo Aceptado Sin Justificar"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.AceptadoSinVisar Then
        getEstadoRiesgoTextoEtiqueta = "Riesgo Aceptado Sin Visar"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Activo Then
        getEstadoRiesgoTextoEtiqueta = "Riesgo Activo"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Cerrado Then
        getEstadoRiesgoTextoEtiqueta = "Riesgo Cerrado"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Detectado Then
        getEstadoRiesgoTextoEtiqueta = "Riesgo Detectado"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Incompleto Then
        getEstadoRiesgoTextoEtiqueta = "Riesgo Incompleto"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Materializado Then
        getEstadoRiesgoTextoEtiqueta = "Riesgo Materializado"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Planificado Then
        getEstadoRiesgoTextoEtiqueta = "Riesgo Planificado"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Retirado Then
        getEstadoRiesgoTextoEtiqueta = "Riesgo Retirado"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.RetiradoRechazado Then
        getEstadoRiesgoTextoEtiqueta = "Riesgo Retirado Rechazado"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.RetiradoSinJustificar Then
        getEstadoRiesgoTextoEtiqueta = "Riesgo Retirado Sin Justificar"
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.RetiradoSinVisar Then
        getEstadoRiesgoTextoEtiqueta = "Riesgo Retirado Sin Visar"
    End If
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEstadoRiesgoTextoEtiqueta ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getEstadoRiesgoColorEtiqueta( _
                                            p_EstadoRiesgo As EnumRiesgoEstado, _
                                            Optional ByRef p_Error As String) As Long
    
   
    On Error GoTo errores
    
    If p_EstadoRiesgo = EnumRiesgoEstado.Aceptado Then
        getEstadoRiesgoColorEtiqueta = ETIQUETA_ESTADO_COLOR_VERDE
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.AceptadoRechazado Then
        getEstadoRiesgoColorEtiqueta = ETIQUETA_ESTADO_COLOR_ROJO
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.AceptadoSinJustificar Then
        getEstadoRiesgoColorEtiqueta = ETIQUETA_ESTADO_COLOR_ROJO
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.AceptadoSinVisar Then
        getEstadoRiesgoColorEtiqueta = ETIQUETA_ESTADO_COLOR_ROJO
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Activo Then
        getEstadoRiesgoColorEtiqueta = ETIQUETA_ESTADO_COLOR_VERDE
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Cerrado Then
        getEstadoRiesgoColorEtiqueta = ETIQUETA_ESTADO_COLOR_VERDE
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Detectado Then
        getEstadoRiesgoColorEtiqueta = ETIQUETA_ESTADO_COLOR_VERDE
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Incompleto Then
        getEstadoRiesgoColorEtiqueta = ETIQUETA_ESTADO_COLOR_ROJO
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Materializado Then
        getEstadoRiesgoColorEtiqueta = ETIQUETA_ESTADO_COLOR_ROJO
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Planificado Then
        getEstadoRiesgoColorEtiqueta = ETIQUETA_ESTADO_COLOR_VERDE
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.Retirado Then
        getEstadoRiesgoColorEtiqueta = ETIQUETA_ESTADO_COLOR_VERDE
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.RetiradoRechazado Then
        getEstadoRiesgoColorEtiqueta = ETIQUETA_ESTADO_COLOR_ROJO
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.RetiradoSinJustificar Then
        getEstadoRiesgoColorEtiqueta = ETIQUETA_ESTADO_COLOR_ROJO
    ElseIf p_EstadoRiesgo = EnumRiesgoEstado.RetiradoSinVisar Then
        getEstadoRiesgoColorEtiqueta = ETIQUETA_ESTADO_COLOR_ROJO
    End If
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEstadoRiesgoColorEtiqueta ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function RellenaEtiquetaRiesgo( _
                                        lbl As Label, _
                                        p_EstadoRiesgo As EnumRiesgoEstado, _
                                        Optional ByRef p_Error As String) As String
    
    Dim m_Texto As String
    Dim m_Color As Long
    On Error GoTo errores
    m_Texto = getEstadoRiesgoTextoEtiqueta(p_EstadoRiesgo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
'    m_Color = getEstadoRiesgoColorEtiqueta(p_EstadoRiesgo, p_Error)
'    If p_Error <> "" Then
'        Err.Raise 1000
'    End If
    lbl.Caption = m_Texto
    lbl.ForeColor = 16777215
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenaEtiquetaRiesgo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Public Function EstablecerComboOrigen( _
                                        p_cmb As ComboBox, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    
    Dim m_Origen As Variant
    Dim m_Descripcion As String
    
    On Error GoTo errores
    
    p_Error = ""
    
    p_cmb.RowSource = ""
    If m_ObjEntorno.ColOrigenRiesgosValores Is Nothing Then
        Exit Function
    End If
    For Each m_Origen In m_ObjEntorno.ColOrigenRiesgosValores
        If CStr(m_Origen) = "Oferta" Then
            If EsTecnico = EnumSiNo.Sí Then
                GoTo siguiente
            End If
        End If
        m_Descripcion = m_ObjEntorno.ColOrigenRiesgosValores(m_Origen)
        p_cmb.AddItem m_Origen & ";" & m_Descripcion
siguiente:
    Next
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerComboOrigen ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function RellenarBordeCamposObligatorios( _
                                                p_Form As Form, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    
    Dim ctl As Control
    
    On Error GoTo errores
    
    p_Error = ""
    
    For Each ctl In p_Form.Controls
        If ctl.ControlType = AcControlType.acTextBox Or _
            ctl.ControlType = AcControlType.acComboBox Then
            If InStr(Nz(ctl.Tag, ""), "OBLIGATORIO") <> 0 Then
                If Nz(ctl.Value, "") = "" Then
                    EstablecerControlCombo ctl, EnumSiNo.No
                    
                Else
                    EstablecerControlCombo ctl, EnumSiNo.Sí
                    
                End If
            End If
        End If
    Next
    RellenaBordeListaSinSeleccion p_Form, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarBordeCamposObligatorios ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function EstablecerControlCombo( _
                                        ctl As Control, _
                                        p_Relleno As EnumSiNo, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    
    
    
    On Error GoTo errores
    
    p_Error = ""
    If p_Relleno = EnumSiNo.Sí Then
        ctl.BorderColor = COLOR_BORDE_CAMPO_RELLENO
        ctl.BorderWidth = ANCHO_BORDE_CAMPO_RELLENO
    ElseIf p_Relleno = EnumSiNo.No Then
        ctl.BorderColor = COLOR_BORDE_CAMPO_NORELLENO
        ctl.BorderWidth = ANCHO_BORDE_CAMPO_NORELLENO
    End If
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerControlCombo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function EstablecerEtiquetaEstado( _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim NavSinActivar As Object
    Dim NavActivado As Object
    Dim miItem As Variant
    Dim ColNavSinActivar As Scripting.Dictionary
    Dim ColNavActivado As Scripting.Dictionary
    
    On Error Resume Next
    p_Error = ""
    If FormularioAbierto("FormRiesgo") Then
        Set NavSinActivar = Forms("FormRiesgo").Controls("NavPlazoCosteCalidad")
        Set NavActivado = Forms("FormRiesgo").Controls("NavRetirado")
    End If
    For Each miItem In NavSinActivar.Properties
        If ColNavSinActivar Is Nothing Then
            Set ColNavSinActivar = New Scripting.Dictionary
        End If
        ColNavSinActivar.CompareMode = TextCompare
        ColNavSinActivar.Add miItem.Name, miItem.Value
    Next
    For Each miItem In NavActivado.Properties
        If ColNavActivado Is Nothing Then
            Set ColNavActivado = New Scripting.Dictionary
        End If
        ColNavActivado.CompareMode = TextCompare
        ColNavActivado.Add miItem.Name, miItem.Value
    Next
    MostrarPropiedades ColNavSinActivar, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerEtiquetaEstado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
    
End Function



Public Function MostrarPropiedades( _
                                    ByRef p_Col As Scripting.Dictionary, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim m_Propiedad As Variant
    Dim m_valor As Variant
    
    
    
    On Error Resume Next
    p_Error = ""
    If p_Col Is Nothing Then
        Exit Function
    End If
    If p_Col.Count = 0 Then
        Exit Function
    End If
    For Each m_Propiedad In p_Col
        Debug.Print m_Propiedad, p_Col(m_Propiedad)
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MostrarPropiedades ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
    
End Function
Private Function RellenaBordeListaSinSeleccion( _
                                                p_Form As Form, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    
    Dim lst As ListBox
    Dim ctl As Control
    Dim m_AlgunElementoSeleccionado As EnumSiNo
    On Error GoTo errores
    
    p_Error = ""
    
    
    
    For Each ctl In p_Form.Controls
        If ctl.ControlType = AcControlType.acListBox Then
            If InStr(Nz(ctl.Tag, ""), "OBLIGATORIO") <> 0 Then
                Set lst = p_Form.Controls(ctl.Name)
                
                If Nz(lst.Value, "") <> "" Then
                    EstablecerControlCombo ctl, EnumSiNo.No
                Else
                    EstablecerControlCombo ctl, EnumSiNo.Sí
                End If
                
                
                
            End If
        End If
siguiente:
    Next
   
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenaBordeListaSinSeleccion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function EstablecerColorFechaFinPrevista(frm As Form) As String

    
    If Nz(frm.FechaFinPrevista, "") <> "" Then
        If Nz(frm.FechaInicio, "") = "" Then
            EstablecerControlCombo frm.FechaInicio, EnumSiNo.No
        Else
            EstablecerControlCombo frm.FechaInicio, EnumSiNo.Sí
            EstablecerControlCombo frm.FechaFinPrevista, EnumSiNo.Sí
        End If
        If Nz(frm.FechaFinReal, "") <> "" Then
            EstablecerControlCombo frm.FechaFinReal, EnumSiNo.Sí
        End If
    Else
        If Nz(frm.FechaInicio, "") <> "" Then
            EstablecerControlCombo frm.FechaFinPrevista, EnumSiNo.No
        Else
            If Nz(frm.FechaFinReal, "") = "" Then
                EstablecerControlCombo frm.FechaFinPrevista, EnumSiNo.Sí
            Else
                EstablecerControlCombo frm.FechaFinPrevista, EnumSiNo.No
            End If
            
        End If
        EstablecerControlCombo frm.FechaInicio, EnumSiNo.Sí
        If Nz(frm.FechaFinReal, "") <> "" Then
            EstablecerControlCombo frm.FechaFinReal, EnumSiNo.No
        Else
            EstablecerControlCombo frm.FechaFinReal, EnumSiNo.Sí
        End If
    End If
    
End Function
Public Function EstablecerColorFechaFinReal(frm As Form) As String

    
    If Nz(frm.FechaFinReal, "") <> "" Then
        If Nz(frm.FechaFinPrevista, "") = "" Then
            EstablecerControlCombo frm.FechaFinPrevista, EnumSiNo.No
        Else
            EstablecerControlCombo frm.FechaFinPrevista, EnumSiNo.Sí
        End If
        If Nz(frm.FechaInicio, "") = "" Then
            EstablecerControlCombo frm.FechaInicio, EnumSiNo.No
        Else
            EstablecerControlCombo frm.FechaInicio, EnumSiNo.Sí
        End If
        If Nz(frm.FechaInicio, "") <> "" Or Nz(frm.FechaFinPrevista, "") <> "" Then
            EstablecerControlCombo frm.FechaFinReal, EnumSiNo.No
        End If
    Else
        If Nz(frm.FechaFinPrevista, "") = "" Then
            If Nz(frm.FechaInicio, "") = "" Then
                EstablecerControlCombo frm.FechaFinPrevista, EnumSiNo.Sí
            Else
                EstablecerControlCombo frm.FechaFinPrevista, EnumSiNo.No
            End If
            
            EstablecerControlCombo frm.FechaFinReal, EnumSiNo.Sí
        Else
            If Nz(frm.FechaInicio, "") = "" Then
                EstablecerControlCombo frm.FechaFinPrevista, EnumSiNo.No
                EstablecerControlCombo frm.FechaFinReal, EnumSiNo.No
            Else
                EstablecerControlCombo frm.FechaFinPrevista, EnumSiNo.Sí
                EstablecerControlCombo frm.FechaFinReal, EnumSiNo.Sí
            End If
            
        End If
    
    End If
    
End Function
Public Function EstablecerColorFechaInicio(frm As Form) As String

    
    If Nz(frm.FechaInicio, "") <> "" Then
        If Nz(frm.FechaFinPrevista, "") = "" Then
            EstablecerControlCombo frm.FechaFinPrevista, EnumSiNo.No
        Else
            EstablecerControlCombo frm.FechaFinPrevista, EnumSiNo.Sí
        End If
        If Nz(frm.FechaFinReal, "") = "" Then
            EstablecerControlCombo frm.FechaInicio, EnumSiNo.Sí
        Else
            EstablecerControlCombo frm.FechaInicio, EnumSiNo.No
        End If
    Else
        If Nz(frm.FechaFinPrevista, "") <> "" Then
            EstablecerControlCombo frm.FechaFinPrevista, EnumSiNo.No
        Else
            EstablecerControlCombo frm.FechaFinPrevista, EnumSiNo.Sí
        End If
    End If
    
End Function

Public Function EstablecerColorPestañasRiesgo(Optional ByRef p_PriorizacionCorrecta As EnumSiNo, Optional ByRef p_Error As String) As String
    
    Dim frm As Form
    Dim flagEstados As String
    Dim m_Estado As EnumRiesgoEstado
    Dim m_Edicion As Edicion
    On Error GoTo errores
    If Not FormularioAbierto("FormRiesgo") Then
        Exit Function
    End If
    Set frm = Forms("FormRiesgo")
    
    With m_ObjRiesgoActivo
        Set m_Edicion = .Edicion
        If m_Edicion.EsActivo = EnumSiNo.No Then
            frm.NavGeneral.ForeColor = PESTAÑA_TODOS_DATOS_OK_SI
            frm.NavPlazoCosteCalidad.ForeColor = PESTAÑA_TODOS_DATOS_OK_SI
            frm.NavPlazoCosteCalidad.ForeColor = PESTAÑA_TODOS_DATOS_OK_SI
            frm.NavVulnerabilidad.ForeColor = PESTAÑA_TODOS_DATOS_OK_SI
            frm.NavMitigacion.ForeColor = PESTAÑA_TODOS_DATOS_OK_SI
            Exit Function
        End If
        m_Estado = .EstadoEnum
        If m_Estado = EnumRiesgoEstado.Retirado Then
            frm.NavGeneral.ForeColor = PESTAÑA_TODOS_DATOS_OK_SI
            frm.NavPlazoCosteCalidad.ForeColor = PESTAÑA_TODOS_DATOS_OK_SI
            frm.NavPlazoCosteCalidad.ForeColor = PESTAÑA_TODOS_DATOS_OK_SI
            frm.NavVulnerabilidad.ForeColor = PESTAÑA_TODOS_DATOS_OK_SI
            frm.NavMitigacion.ForeColor = PESTAÑA_TODOS_DATOS_OK_SI
            Exit Function
        End If
        If p_PriorizacionCorrecta = Empty Then
            p_PriorizacionCorrecta = Constructor.EsPriorizacionCorrecta(m_Edicion.IDEdicion, _
                                    .Priorizacion, .IDRiesgo)
        End If
        
        If m_Edicion.Proyecto.RequiereRiesgoDeBiblioteca = "Sí" Then
             
            If p_PriorizacionCorrecta = EnumSiNo.No Then
                frm.NavGeneral.ForeColor = PESTAÑA_TODOS_DATOS_OK_NO
                flagEstados = "0"
            Else
                If .Descripcion = "" Or .Origen = "" Or .CausaRaiz = "" Or .FechaDetectado = "" Or _
                        .EntidadDetecta = "" Or .DetectadoPor = "" Then
                    frm.NavGeneral.ForeColor = PESTAÑA_TODOS_DATOS_OK_NO
                    flagEstados = "0"
                Else
                    
                    frm.NavGeneral.ForeColor = PESTAÑA_TODOS_DATOS_OK_SI
                    flagEstados = "1"
                End If
            End If
            
        Else
            If p_PriorizacionCorrecta = EnumSiNo.No Then
                frm.NavGeneral.ForeColor = PESTAÑA_TODOS_DATOS_OK_NO
                flagEstados = "0"
            Else
                If .Descripcion = "" Or .Origen = "" Or .FechaDetectado = "" Or _
                        .EntidadDetecta = "" Or .DetectadoPor = "" Then
                    frm.NavGeneral.ForeColor = PESTAÑA_TODOS_DATOS_OK_NO
                    flagEstados = "0"
                Else
                    frm.NavGeneral.ForeColor = PESTAÑA_TODOS_DATOS_OK_SI
                    flagEstados = "1"
                End If
            End If
            
        End If
        If .Plazo = "" Or .Coste = "" Or .Calidad = "" Then
            frm.NavPlazoCosteCalidad.ForeColor = PESTAÑA_TODOS_DATOS_OK_NO
            flagEstados = flagEstados & "0"
        Else
            frm.NavPlazoCosteCalidad.ForeColor = PESTAÑA_TODOS_DATOS_OK_SI
            flagEstados = flagEstados & "1"
        End If
        If .Vulnerabilidad = "" Then
            frm.NavVulnerabilidad.ForeColor = PESTAÑA_TODOS_DATOS_OK_NO
            flagEstados = flagEstados & "0"
        Else
            frm.NavVulnerabilidad.ForeColor = PESTAÑA_TODOS_DATOS_OK_SI
            flagEstados = flagEstados & "1"
        End If
        If .Mitigacion = "" Then
            frm.NavMitigacion.ForeColor = PESTAÑA_TODOS_DATOS_OK_NO
            flagEstados = flagEstados & "0"
        Else
            frm.NavMitigacion.ForeColor = PESTAÑA_TODOS_DATOS_OK_SI
            flagEstados = flagEstados & "1"
        End If
    End With
    If InStr(1, flagEstados, "0") = 0 Then
        If frm.lblEstadoRiesgo.ForeColor = ETIQUETA_ESTADO_COLOR_ROJO Then
            ActualizarEtiquetaEstadoRiesgo EnumSiNo.Sí
        End If
    Else
        If frm.lblEstadoRiesgo.ForeColor = ETIQUETA_ESTADO_COLOR_VERDE Then
            ActualizarEtiquetaEstadoRiesgo EnumSiNo.Sí
        End If
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCaEstablecerColorPestañasmbios ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function ActualizarEtiquetaEstadoRiesgo( _
                                                Optional p_Actualizando As EnumSiNo = EnumSiNo.No, _
                                                Optional ByRef p_Error As String) As String
    
    Dim frm As Form
    Dim m_EstadoRiesgo As EnumRiesgoEstado
    On Error GoTo errores
    If Not FormularioAbierto("FormRiesgo") Then
        Exit Function
    End If
    Set frm = Forms("FormRiesgo")
    
    If p_Actualizando = EnumSiNo.Sí Then
        m_EstadoRiesgo = m_ObjRiesgoActivo.ESTADOCalculado
    Else
        m_EstadoRiesgo = m_ObjRiesgoActivo.EstadoEnum
    End If
    
    If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
        RellenaEtiquetaRiesgo frm.lblEstadoRiesgo, m_EstadoRiesgo, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        frm.lblEstadoRiesgo.Visible = False
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizarEtiquetaEstadoRiesgo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Public Function getEstadosDiferentesHastaEdicion( _
                                                p_EdicionMaxima As Edicion, _
                                                p_CodigoRiesgo As String, _
                                                Optional p_FechaPublicacion As String, _
                                                Optional p_FechaCierre As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
                
    Dim m_Edicion As Edicion
    Dim m_Proyecto As Proyecto
    Dim m_Riesgo As riesgo
    Dim m_ColMaterializados As Scripting.Dictionary
    Dim m_IDEdicion As Variant
    
    Dim i As Integer
    Dim m_Id As Variant
    Dim m_ResultadoMaterializado As String
    Dim m_UltimoEstado As String
    
    ' Variables para las fuentes de datos (Corregidas para usar la REALIDAD del dato)
    Dim sFechaDetectado As String
    Dim sEstadoFoto As String
    Dim sFechaFoto As String
    
    On Error GoTo errores
    
    ' -------------------------------------------------------------------------
    ' 1. INICIALIZACIÓN
    ' -------------------------------------------------------------------------
    Set getEstadosDiferentesHastaEdicion = New Scripting.Dictionary
    getEstadosDiferentesHastaEdicion.CompareMode = TextCompare
    
    Set m_Proyecto = p_EdicionMaxima.Proyecto
    If m_Proyecto Is Nothing Then Err.Raise 1000, , "No se ha podido obtener el proyecto"
    
    ' validación de seguridad
    If Not m_Proyecto.EdicionUltima Is Nothing Then
        If CInt(m_Proyecto.EdicionUltima.IDEdicion) < CInt(p_EdicionMaxima.IDEdicion) Then
            Err.Raise 1000, , "La Edición solicitada es posterior a la última existente."
        End If
    End If
    
    i = 1
    m_UltimoEstado = ""
    
    ' -------------------------------------------------------------------------
    ' 2. BUCLE CRONOLÓGICO
    ' -------------------------------------------------------------------------
    For Each m_IDEdicion In m_Proyecto.colEdiciones
        Set m_Edicion = m_Proyecto.colEdiciones(m_IDEdicion)
        'If p_CodigoRiesgo = "R011" And m_edicion.Edicion = "4" Then Stop
        ' Límite solicitado
        If CInt(m_Edicion.IDEdicion) > CInt(p_EdicionMaxima.IDEdicion) Then Exit For
        
        ' Cargar Riesgo en esta Edición (FOTO FIJA)
        Set m_Riesgo = Constructor.getRiesgo(, m_Edicion.IDEdicion, p_CodigoRiesgo)
        
        If m_Riesgo Is Nothing Then GoTo siguienteEdicion
        
        ' --- A. ESTADO DETECTADO (Primera vez que aparece) ---
        If m_UltimoEstado = "" Then
            sFechaDetectado = m_Riesgo.FechaDetectado
            ' Fallback solo si es nula
            If Not IsDate(sFechaDetectado) Then sFechaDetectado = m_Edicion.FechaPublicacion
            
            getEstadosDiferentesHastaEdicion.Add CStr(i), "Detectado" & "|" & sFechaDetectado
            m_UltimoEstado = "Detectado"
            i = i + 1
        End If
        
        ' --- B. MATERIALIZACIONES (Durante la Edición) ---
        ' Usamos tu función auxiliar que filtra por IDEdicion
        Set m_ColMaterializados = getEstadosDiferentesEnEdicionTbMaterializados( _
                                        m_Edicion.IDEdicion, _
                                        p_CodigoRiesgo, _
                                        m_UltimoEstado, _
                                        p_Error)
        
        If Not m_ColMaterializados Is Nothing Then
            For Each m_Id In m_ColMaterializados
                m_ResultadoMaterializado = m_ColMaterializados(m_Id)
                getEstadosDiferentesHastaEdicion.Add CStr(i), m_ResultadoMaterializado
                i = i + 1
                ' Actualizamos puntero
                m_UltimoEstado = Split(m_ResultadoMaterializado, "|")(0)
            Next
        End If
        
        ' --- C. FOTO FIJA DE LA Edición ---
        sEstadoFoto = m_Riesgo.ESTADOCalculadoTexto
        sFechaFoto = getFechaEstadoRiesgoSegunEstado(m_Riesgo, sEstadoFoto, m_Edicion, p_FechaCierre, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
        If Not IsDate(sFechaFoto) Then
            If IsDate(m_Edicion.FechaPublicacion) Then
                sFechaFoto = m_Edicion.FechaPublicacion
            Else
                sFechaFoto = m_Edicion.FechaEdicion
            End If
        End If
        
        ' Si el estado final de la Edición es diferente al último registrado
        If sEstadoFoto <> m_UltimoEstado Then
            getEstadosDiferentesHastaEdicion.Add CStr(i), sEstadoFoto & "|" & sFechaFoto
            m_UltimoEstado = sEstadoFoto
            i = i + 1
        End If
        
siguienteEdicion:
    Next
    
    ' -------------------------------------------------------------------------
    ' 3. CIERRE FORZADO (Solo si el proyecto se cerró administrativamente)
    ' -------------------------------------------------------------------------
    If CInt(m_Proyecto.EdicionUltima.IDEdicion) = CInt(p_EdicionMaxima.IDEdicion) Then
        Dim bCerrar As Boolean
        Dim sFechaCierreFin As String
        
        If IsDate(p_FechaCierre) Then
            bCerrar = True
            sFechaCierreFin = p_FechaCierre
        ElseIf IsDate(m_Proyecto.FechaCierre) Then
            bCerrar = True
            sFechaCierreFin = m_Proyecto.FechaCierre
        End If
        
        If bCerrar Then
            If InStr(1, m_UltimoEstado, "Cerrado", vbTextCompare) = 0 And _
               InStr(1, m_UltimoEstado, "Retirado", vbTextCompare) = 0 Then
               
                getEstadosDiferentesHastaEdicion.Add CStr(i), "Cerrado" & "|" & sFechaCierreFin
                i = i + 1
            End If
        End If
    End If

    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en getEstadosDiferentesHastaEdicion: " & Err.Description
    End If
End Function
Private Function getFechaEstadoRiesgoSegunEstado( _
                                                p_Riesgo As riesgo, _
                                                p_Estado As String, _
                                                p_Edicion As Edicion, _
                                                Optional p_FechaCierre As String, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim m_Estado As String
    Dim m_Fecha As String
    Dim m_Min As String
    Dim m_Max As String
    
    On Error GoTo errores
    
    If p_Riesgo Is Nothing Then
        Exit Function
    End If
    
    m_Estado = Trim(Nz(p_Estado, ""))
    
    If InStr(1, m_Estado, "Aceptado", vbTextCompare) <> 0 Then
        If IsDate(p_Riesgo.FechaAprobacionAceptacionPorCalidad) Then
            m_Fecha = p_Riesgo.FechaAprobacionAceptacionPorCalidad
        ElseIf IsDate(p_Riesgo.FechaJustificacionAceptacionRiesgo) Then
            m_Fecha = p_Riesgo.FechaJustificacionAceptacionRiesgo
        ElseIf IsDate(p_Riesgo.FechaMitigacionAceptar) Then
            m_Fecha = p_Riesgo.FechaMitigacionAceptar
        End If
        getFechaEstadoRiesgoSegunEstado = m_Fecha
        Exit Function
    End If
    
    If InStr(1, m_Estado, "Retirado", vbTextCompare) <> 0 Then
        If IsDate(p_Riesgo.FechaAprobacionRetiroPorCalidad) Then
            m_Fecha = p_Riesgo.FechaAprobacionRetiroPorCalidad
        ElseIf IsDate(p_Riesgo.FechaRetirado) Then
            m_Fecha = p_Riesgo.FechaRetirado
        ElseIf IsDate(p_Riesgo.FechaJustificacionRetiroRiesgo) Then
            m_Fecha = p_Riesgo.FechaJustificacionRetiroRiesgo
        End If
        getFechaEstadoRiesgoSegunEstado = m_Fecha
        Exit Function
    End If
    
    If InStr(1, m_Estado, "Materializado", vbTextCompare) <> 0 Then
        If IsDate(p_Riesgo.FechaMaterializado) Then
            m_Fecha = p_Riesgo.FechaMaterializado
        End If
        getFechaEstadoRiesgoSegunEstado = m_Fecha
        Exit Function
    End If
    
    If InStr(1, m_Estado, "Cerrado", vbTextCompare) <> 0 Then
        If IsDate(p_Riesgo.FechaCerrado) Then
            m_Fecha = p_Riesgo.FechaCerrado
        ElseIf IsDate(p_FechaCierre) Then
            m_Fecha = p_FechaCierre
        ElseIf Not p_Riesgo.Edicion Is Nothing Then
            If IsDate(p_Riesgo.Edicion.Proyecto.FechaCierre) Then
                m_Fecha = p_Riesgo.Edicion.Proyecto.FechaCierre
            End If
        End If
        getFechaEstadoRiesgoSegunEstado = m_Fecha
        Exit Function
    End If
    
    If InStr(1, m_Estado, "Activo", vbTextCompare) <> 0 Then
        m_Min = getMinFechaPMAccion(p_Riesgo.IDRiesgo, "Inicio", p_Error)
        If p_Error <> "" Then Err.Raise 1000
        getFechaEstadoRiesgoSegunEstado = m_Min
        Exit Function
    End If
    
    If InStr(1, m_Estado, "Planificado", vbTextCompare) <> 0 Then
        m_Min = getMinFechaPMAccion(p_Riesgo.IDRiesgo, "FinPrevista", p_Error)
        If p_Error <> "" Then Err.Raise 1000
        getFechaEstadoRiesgoSegunEstado = m_Min
        Exit Function
    End If
    
    If InStr(1, m_Estado, "Detectado", vbTextCompare) <> 0 Then
        If p_Riesgo.TienePMs = EnumSiNo.Sí And p_Riesgo.TodosPMFinalizados = EnumSiNo.Sí Then
            m_Max = getMaxFechaPMAccion(p_Riesgo.IDRiesgo, "FinReal", p_Error)
            If p_Error <> "" Then Err.Raise 1000
            m_Fecha = m_Max
        Else
            m_Fecha = p_Riesgo.FechaDetectado
        End If
        If Not IsDate(m_Fecha) Then
            If Not p_Edicion Is Nothing Then
                If IsDate(p_Edicion.FechaPublicacion) Then
                    m_Fecha = p_Edicion.FechaPublicacion
                Else
                    m_Fecha = p_Edicion.FechaEdicion
                End If
            End If
        End If
        getFechaEstadoRiesgoSegunEstado = m_Fecha
        Exit Function
    End If
    
    If InStr(1, m_Estado, "Incompleto", vbTextCompare) <> 0 Then
        m_Fecha = p_Riesgo.FechaDetectado
        If Not IsDate(m_Fecha) Then
            If Not p_Edicion Is Nothing Then
                If IsDate(p_Edicion.FechaPublicacion) Then
                    m_Fecha = p_Edicion.FechaPublicacion
                Else
                    m_Fecha = p_Edicion.FechaEdicion
                End If
            End If
        End If
        getFechaEstadoRiesgoSegunEstado = m_Fecha
        Exit Function
    End If
    
    getFechaEstadoRiesgoSegunEstado = p_Riesgo.FechaEstado
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getFechaEstadoRiesgoSegunEstado ha devuelto el error: " & Err.Description
    End If
End Function
Private Function getMaxFechaPMAccion( _
                                    p_IDRiesgo As String, _
                                    p_TipoFecha As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As String
    
    On Error GoTo errores
    
    If Trim(Nz(p_IDRiesgo, "")) = "" Then Exit Function
    
    If p_TipoFecha = "Inicio" Then
        m_Campo = "TbRiesgosPlanMitigacionDetalle.FechaInicio"
    ElseIf p_TipoFecha = "FinPrevista" Then
        m_Campo = "TbRiesgosPlanMitigacionDetalle.FechaFinPrevista"
    ElseIf p_TipoFecha = "FinReal" Then
        m_Campo = "TbRiesgosPlanMitigacionDetalle.FechaFinReal"
    Else
        Exit Function
    End If
    
    m_SQL = "SELECT Max(" & m_Campo & ") AS MaxFecha " & _
            "FROM TbRiesgosPlanMitigacionPpal " & _
            "INNER JOIN TbRiesgosPlanMitigacionDetalle " & _
            "ON TbRiesgosPlanMitigacionPpal.IDMitigacion = TbRiesgosPlanMitigacionDetalle.IDMitigacion " & _
            "WHERE TbRiesgosPlanMitigacionPpal.IDRiesgo=" & p_IDRiesgo & " " & _
            "AND Not (" & m_Campo & " Is Null);"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsDate(.Fields("MaxFecha").Value) Then
                getMaxFechaPMAccion = CStr(.Fields("MaxFecha").Value)
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    If Err.Number <> 1000 Then
        p_Error = "El método getMaxFechaPMAccion ha devuelto el error: " & Err.Description
    End If
End Function
Private Function getMinFechaPMAccion( _
                                    p_IDRiesgo As String, _
                                    p_TipoFecha As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As String
    
    On Error GoTo errores
    
    If Trim(Nz(p_IDRiesgo, "")) = "" Then Exit Function
    
    If p_TipoFecha = "Inicio" Then
        m_Campo = "TbRiesgosPlanMitigacionDetalle.FechaInicio"
    ElseIf p_TipoFecha = "FinPrevista" Then
        m_Campo = "TbRiesgosPlanMitigacionDetalle.FechaFinPrevista"
    ElseIf p_TipoFecha = "FinReal" Then
        m_Campo = "TbRiesgosPlanMitigacionDetalle.FechaFinReal"
    Else
        Exit Function
    End If
    
    m_SQL = "SELECT Min(" & m_Campo & ") AS MinFecha " & _
            "FROM TbRiesgosPlanMitigacionPpal " & _
            "INNER JOIN TbRiesgosPlanMitigacionDetalle " & _
            "ON TbRiesgosPlanMitigacionPpal.IDMitigacion = TbRiesgosPlanMitigacionDetalle.IDMitigacion " & _
            "WHERE TbRiesgosPlanMitigacionPpal.IDRiesgo=" & p_IDRiesgo & " " & _
            "AND Not (" & m_Campo & " Is Null);"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsDate(.Fields("MinFecha").Value) Then
                getMinFechaPMAccion = CStr(.Fields("MinFecha").Value)
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    If Err.Number <> 1000 Then
        p_Error = "El método getMinFechaPMAccion ha devuelto el error: " & Err.Description
    End If
End Function
Public Function ResincronizarEstadosRiesgoHastaEdicion( _
                                                        p_EdicionMaxima As Edicion, _
                                                        p_CodigoRiesgo As String, _
                                                        Optional p_FechaCierre As String, _
                                                        Optional p_Forzar As EnumSiNo = EnumSiNo.No, _
                                                        Optional ByRef p_Error As String _
                                                        ) As String
    Dim m_Edicion As Edicion
    Dim m_Proyecto As Proyecto
    Dim m_Riesgo As riesgo
    Dim m_IDEdicion As Variant
    Dim m_EstadoCalculado As String
    Dim m_FechaCalculada As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_FechaActual As String
    Dim m_EstadoActual As String
    Dim bActualizar As Boolean
    
    On Error GoTo errores
    
    If p_EdicionMaxima Is Nothing Then
        p_Error = "Se ha de indicar la Edición"
        Err.Raise 1000
    End If
    If Trim(Nz(p_CodigoRiesgo, "")) = "" Then
        p_Error = "Se ha de indicar el código de riesgo"
        Err.Raise 1000
    End If
    
    Set m_Proyecto = p_EdicionMaxima.Proyecto
    If m_Proyecto Is Nothing Then
        p_Error = "No se ha podido obtener el proyecto"
        Err.Raise 1000
    End If
    
    For Each m_IDEdicion In m_Proyecto.colEdiciones
        Set m_Edicion = m_Proyecto.colEdiciones(m_IDEdicion)
        If CInt(m_Edicion.IDEdicion) > CInt(p_EdicionMaxima.IDEdicion) Then Exit For
        
        Set m_Riesgo = Constructor.getRiesgo(, m_Edicion.IDEdicion, p_CodigoRiesgo)
        If m_Riesgo Is Nothing Then GoTo siguienteEdicion
        
        m_EstadoCalculado = m_Riesgo.ESTADOCalculadoTexto
        m_FechaCalculada = getFechaEstadoRiesgoSegunEstado(m_Riesgo, m_EstadoCalculado, m_Edicion, p_FechaCierre, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
        m_SQL = "SELECT Estado, FechaEstado FROM TbRiesgos WHERE IDRiesgo=" & m_Riesgo.IDRiesgo & ";"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                .Close
                Set rcdDatos = Nothing
                GoTo siguienteEdicion
            End If
            m_EstadoActual = Nz(.Fields("Estado").Value, "")
            m_FechaActual = Nz(.Fields("FechaEstado").Value, "")
            bActualizar = False
            
            If p_Forzar = EnumSiNo.Sí Then
                bActualizar = True
            ElseIf StrComp(Trim(m_EstadoActual), Trim(m_EstadoCalculado), vbTextCompare) <> 0 Then
                bActualizar = True
            ElseIf IsDate(m_FechaCalculada) And IsDate(m_FechaActual) Then
                If CDate(m_FechaActual) <> CDate(m_FechaCalculada) Then
                    bActualizar = True
                End If
            ElseIf (IsDate(m_FechaCalculada) Xor IsDate(m_FechaActual)) Then
                bActualizar = True
            End If
            
            If bActualizar Then
                .Edit
                    .Fields("Estado").Value = m_EstadoCalculado
                    If IsDate(m_FechaCalculada) Then
                        .Fields("FechaEstado").Value = Format(CDate(m_FechaCalculada), "mm/dd/yyyy")
                    Else
                        .Fields("FechaEstado").Value = Null
                    End If
                .Update
            End If
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        
siguienteEdicion:
        Set m_Riesgo = Nothing
    Next
    
    Exit Function
errores:
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    If Err.Number <> 1000 Then
        p_Error = "El método ResincronizarEstadosRiesgoHastaEdicion ha devuelto el error: " & Err.Description
    End If
End Function
Private Function getEstadosDiferentesEnEdicionTbMaterializados( _
                                                                p_IDEdicion As String, _
                                                                p_CodigoRiesgo As String, _
                                                                Optional p_EstadoAnterior As String, _
                                                                Optional ByRef p_Error As String _
                                                                ) As Scripting.Dictionary
                
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_EstadoTabla As String
    Dim m_EsMaterializado As String
    Dim m_FechaEstado As String
    Dim m_Id As String
    Dim i As Integer
    
    On Error GoTo errores
    
    m_SQL = "SELECT * " & _
                "FROM TbRiesgosMaterializaciones " & _
                "WHERE IDEdicion=" & p_IDEdicion & " " & _
                "AND CodigoRiesgo='" & p_CodigoRiesgo & "' " & _
                "ORDER BY Fecha;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        i = 1
        .MoveFirst
        Do While Not .EOF
            m_EstadoTabla = .Fields("Estado")
            m_FechaEstado = .Fields("Fecha")
            m_Id = .Fields("ID")
            If i = 1 And p_EstadoAnterior <> "" Then
                If m_EstadoTabla = p_EstadoAnterior Then
                    GoTo siguiente
                End If
            End If
            If getEstadosDiferentesEnEdicionTbMaterializados Is Nothing Then
                Set getEstadosDiferentesEnEdicionTbMaterializados = New Scripting.Dictionary
                getEstadosDiferentesEnEdicionTbMaterializados.CompareMode = TextCompare
            End If
            If Not getEstadosDiferentesEnEdicionTbMaterializados.Exists(CStr(m_Id)) Then
                getEstadosDiferentesEnEdicionTbMaterializados.Add CStr(m_Id), m_EstadoTabla & "|" & m_FechaEstado
            End If
            
siguiente:
            i = i + 1
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEstadosDiferentesEnEdicionTbMaterializados ha devuelto el error: " & Err.Description
    End If
End Function
Private Function getUltimoElementoDeCole(ByRef p_Col As Scripting.Dictionary, Optional ByRef p_Error As String) As String
    
    Dim m_Id As Variant
    Dim m_Resultado As String
    
    On Error GoTo errores
    If p_Col Is Nothing Then
        Exit Function
    End If
    For Each m_Id In p_Col
        m_Resultado = p_Col(m_Id)
    Next
    getUltimoElementoDeCole = m_Resultado
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUltimoElementoDeCole ha devuelto el error: " & Err.Description
    End If
End Function
Private Function getUltimoEstadoDeCole(ByRef p_Col As Scripting.Dictionary, Optional ByRef p_Error As String) As String
    
    Dim m_Id As Variant
    Dim m_Resultado As String
    Dim dato As Variant
    On Error GoTo errores
    m_Resultado = getUltimoElementoDeCole(p_Col)
    If InStr(1, m_Resultado, "|") <> 0 Then
        dato = Split(m_Resultado, "|")
        getUltimoEstadoDeCole = dato(0)
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUltimoElementoDeCole ha devuelto el error: " & Err.Description
    End If
End Function
Public Function CopiarRiesgoCompleto( _
                                        p_RiesgoOrigen As riesgo, _
                                        Optional p_Error As String _
                                        ) As riesgo
    Dim m_objEdicionOrigen As Edicion
    Dim m_Edicion As Edicion
    Dim m_ObjProyecto As Proyecto
    Dim m_Riesgo As riesgo
    
    
    
    Dim m_IDPX As Variant
    Dim m_ObjPM As PM
    Dim m_ObjPC As PC
    Dim m_IdAccion As Variant
    Dim m_ObjPMAccion As PMAccion
    Dim m_ObjPCAccion As PCAccion
    Dim m_IDEdicionNueva As String
    Dim m_ObjRiesgoNuevo As riesgo
    Dim m_IDRiesgoNuevo As String
    Dim m_ObjPMNuevo As PM
    Dim m_IDPMNuevo As String
    Dim m_ObjPMAccionNueva As PMAccion
    Dim m_IDPMAccionNueva As String
    Dim m_ObjPCNuevo As PC
    Dim m_IDPCNuevo As String
    Dim m_ObjPCAccionNueva As PCAccion
    Dim m_IDPCAccionNueva As String
    
    
    On Error GoTo errores
    If m_ObjEntorno Is Nothing Then
        EVE , p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Set m_objEdicionOrigen = p_RiesgoOrigen.Edicion
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjProyecto = m_objEdicionOrigen.Proyecto
    Set m_Edicion = m_objEdicionOrigen.EdicionSiguiente
    If m_Edicion Is Nothing Then
        Exit Function
    End If
    Set m_Riesgo = Constructor.getRiesgo(, m_Edicion.IDEdicion, p_RiesgoOrigen.CodigoRiesgo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Riesgo Is Nothing Then
        Exit Function
    End If
    Set m_Riesgo = CopiarRiesgo(p_RiesgoOrigen.IDRiesgo, m_Edicion.IDEdicion, EnumSiNo.No, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_IDRiesgoNuevo = m_Riesgo.IDRiesgo
    If m_Riesgo.TienePMs = EnumSiNo.Sí Then
        For Each m_IDPX In m_Riesgo.ColPMs.keys
            Set m_ObjPM = m_Riesgo.ColPMs(m_IDPX)
            Set m_ObjPMNuevo = CopiarPM(m_ObjPM.IDMitigacion, m_IDRiesgoNuevo, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Not m_ObjPMNuevo Is Nothing Then
                m_IDPMNuevo = m_ObjPMNuevo.IDMitigacion
                If m_ObjPM.TieneAcciones = EnumSiNo.Sí Then
                    For Each m_IdAccion In m_ObjPM.colAcciones.keys
                        
                        Set m_ObjPMAccion = m_ObjPM.colAcciones(m_IdAccion)
                        Set m_ObjPMAccionNueva = CopiarPMAccion(CStr(m_IdAccion), m_IDPMNuevo, p_Error)
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                        Set m_ObjPMAccion = Nothing
                        Set m_ObjPMAccionNueva = Nothing

                    Next
                End If
            End If
            Set m_ObjPM = Nothing
            Set m_ObjPMNuevo = Nothing

        Next
    End If
    If m_Riesgo.TienePCs = EnumSiNo.Sí Then
        For Each m_IDPX In m_Riesgo.ColPCs.keys
            Set m_ObjPC = m_Riesgo.ColPCs(m_IDPX)
            Set m_ObjPCNuevo = CopiarPC(CStr(m_IDPX), m_IDRiesgoNuevo, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            m_IDPCNuevo = m_ObjPCNuevo.IDContingencia
            If Not m_ObjPC.colAcciones Is Nothing Then
                For Each m_IdAccion In m_ObjPC.colAcciones.keys
                    Set m_ObjPCAccion = m_ObjPC.colAcciones(m_IdAccion)
                    Set m_ObjPCAccionNueva = CopiarPCAccion(CStr(m_IdAccion), m_IDPCNuevo, p_Error)
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    Set m_ObjPCAccion = Nothing
                    Set m_ObjPCAccionNueva = Nothing
                Next
            End If
            Set m_ObjPC = Nothing
            Set m_ObjPCNuevo = Nothing
        Next
    End If
    Set CopiarRiesgoCompleto = m_Riesgo

    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CopiarRiesgoCompleto ha devuelto el error: " & Err.Description
    End If
    
End Function
Public Function getHora() As String
    Dim xmlHttp As Object
    Dim url As String
    Dim response As String
    Dim json As Object
    Dim hora As String
    Dim dato As Variant
    Dim flag As String
    
    ' URL de la API para obtener la hora en Madrid, España
    url = "http://worldtimeapi.org/api/timezone/Europe/Madrid"
    
    ' Crear el objeto XMLHTTP
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Hacer la solicitud a la API
    xmlHttp.Open "GET", url, False
    xmlHttp.send
    
    ' Obtener la respuesta
    response = xmlHttp.responseText
    
    ' Analizar la respuesta JSON
    Set json = JsonConverter.ParseJson(response)
    
    ' Extraer la hora de la respuesta JSON
    hora = json("datetime")
    dato = Split(hora, "T")
    flag = dato(1)
    dato = Split(flag, ".")
    hora = dato(0)
    ' Retornar la hora
    getHora = hora
End Function
Public Function EstablecerlblRechazado(p_Form As Form, Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    With m_ObjEdicionActiva
        If .PublicacionRechazada = EnumSiNo.Sí Then
            p_Form.lblRechazoCalidad.Visible = True
            p_Form.lblRechazoCalidad.Caption = "PUBLICACIÓN RECHAZADA POR CALIDAD EL " & .PropuestaRechazadaPorCalidadFecha
        Else
            p_Form.lblRechazoCalidad.Visible = False
        End If
    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerlblRechazado ha producido el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function GetNombreParaNodo( _
                                    p_Riesgo As riesgo, _
                                    Optional ByRef p_VerDescripcion As EnumSiNo = EnumSiNo.Sí, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim m_Priorizacion As String
    Dim m_TituloNodo As String
    Dim m_VerDescripcion As EnumSiNo
    On Error GoTo errores
    
    If p_VerDescripcion = 0 Then
        If m_ObjEntorno.VerRiesgosDescripcion = 0 Then
            m_ObjEntorno.VerRiesgosDescripcion = EnumSiNo.Sí
        End If
        m_VerDescripcion = m_ObjEntorno.VerRiesgosDescripcion
    Else
        m_VerDescripcion = p_VerDescripcion
    End If
    
    With p_Riesgo
        If .Priorizacion = "" Then
            m_Priorizacion = "(--)"
        Else
            m_Priorizacion = "(" & Format(.Priorizacion, "00") & ")"
        End If
        
        If m_VerDescripcion = EnumSiNo.Sí Then
            If .CodRiesgoBiblioteca <> "" Then
                m_TituloNodo = m_Priorizacion & " " & .CodigoRiesgo & " (" & Left(.CausaRaiz, 100) & "... )"
            Else
                m_TituloNodo = m_Priorizacion & " " & .CodigoRiesgo & " (" & Left(.Descripcion, 100) & "... )"
            End If
        Else
            m_TituloNodo = m_Priorizacion & " " & .CodigoRiesgo & " (" & Trim(.ESTADOCalculadoTexto) & ")"
        End If
            
    End With
    GetNombreParaNodo = m_TituloNodo
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GetNombreParaNodo ha producido el error: " & vbNewLine & Err.Description
    End If
End Function


Public Function setTitulosNodosRiesgos( _
                                        tv As MSComctlLib.TreeView, _
                                        Optional ByRef p_VerDescripcion As EnumSiNo = EnumSiNo.Sí, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    Dim nodoRaiz As MSComctlLib.Node
    Dim nodoHijo As MSComctlLib.Node
    Dim m_Riesgo As riesgo
    Dim m_TituloNodo As String
    Dim dato As Variant
    Dim m_IdRiesgo As String
    Dim m_Tipo As String
    Dim m_ColorEstado As Long
    
    On Error GoTo errores
    
    If tv.Nodes.Count = 0 Then
        Exit Function
    End If
    Set nodoRaiz = tv.Nodes(1)
    
    Set nodoHijo = nodoRaiz.Child
    If nodoHijo Is Nothing Then
        Exit Function
    End If
    Do While Not nodoHijo Is Nothing
        If InStr(1, nodoHijo.Key, "|") = 0 Then
            GoTo siguiente
        End If
        dato = Split(nodoHijo.Key, "|")
        m_Tipo = dato(0)
        If m_Tipo <> "RIESGO" Then
            GoTo siguiente
        End If
        m_IdRiesgo = dato(1)
        Set m_Riesgo = Constructor.getRiesgo(p_IDRiesgo:=m_IdRiesgo)
        
        m_TituloNodo = GetNombreParaNodo(p_Riesgo:=m_Riesgo, p_VerDescripcion:=p_VerDescripcion)
        If m_TituloNodo <> "" Then
            nodoHijo.Text = m_TituloNodo
            m_ColorEstado = -1
            m_ColorEstado = getColorNodo(p_Riesgo:=m_Riesgo, p_Error:=p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ColorEstado <> -1 Then
                nodoHijo.ForeColor = m_ColorEstado
            Else
                nodoHijo.ForeColor = vbBlack
            End If
            
            
        End If
siguiente:
        ' Pasar al siguiente hermano del hijo actual
        Set nodoHijo = nodoHijo.Next
    Loop
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método setTitulosNodosRiesgos ha producido el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function setModoVerArbolRiesgos( _
                                        p_VerDescripcionActual As String, _
                                        p_VerSoloNoRetirados As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    On Error GoTo errores
    'si p_VerDescripcionActual="Sí"--->grabo en la tabla ese valor y pongo m_ObjEntorno.VerRiesgosDescripcion
    
    If p_VerDescripcionActual <> "Sí" And p_VerDescripcionActual <> "No" Then
        Exit Function
    End If
    If p_VerSoloNoRetirados <> "Sí" And p_VerSoloNoRetirados <> "No" Then
        Exit Function
    End If
    If p_VerDescripcionActual = "Sí" Then
        m_ObjEntorno.VerRiesgosDescripcion = EnumSiNo.Sí
    Else
        m_ObjEntorno.VerRiesgosDescripcion = EnumSiNo.No
    End If
    If p_VerSoloNoRetirados = "Sí" Then
        m_ObjEntorno.VerSoloRiesgosNoRetirados = EnumSiNo.No
    Else
        m_ObjEntorno.VerSoloRiesgosNoRetirados = EnumSiNo.Sí
    End If
    m_SQL = "SELECT * FROM TbConfiguracionVisionRiesgos " & _
            "WHERE Usuario='" & m_ObjUsuarioConectado.UsuarioRed & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            .AddNew
                .Fields("IDConfiguracion") = DameID("TbConfiguracionVisionRiesgos", "IDConfiguracion", getdb())
                .Fields("Usuario") = m_ObjUsuarioConectado.UsuarioRed
                
                If m_ObjEntorno.VerSoloRiesgosNoRetirados = EnumSiNo.Sí Then
                    .Fields("VerSoloNoRetirados") = "Sí"
                Else
                    .Fields("VerSoloNoRetirados") = "No"
                End If
                If m_ObjEntorno.VerRiesgosDescripcion = EnumSiNo.Sí Then
                    .Fields("VerDescripcion") = "Sí"
                Else
                    .Fields("VerDescripcion") = "No"
                End If
                
            .Update
           
        Else
            .Edit
                
                
                If m_ObjEntorno.VerSoloRiesgosNoRetirados = EnumSiNo.Sí Then
                    .Fields("VerSoloNoRetirados") = "Sí"
                Else
                    .Fields("VerSoloNoRetirados") = "No"
                End If
                If m_ObjEntorno.VerRiesgosDescripcion = EnumSiNo.Sí Then
                    .Fields("VerDescripcion") = "Sí"
                Else
                    .Fields("VerDescripcion") = "No"
                End If
            .Update
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método setModoVerArbolRiesgos ha producido el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function setVerSoloRiesgosNoRetirados( _
                                                p_VerSoloRiesgosNoRetiradosActual As String, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    On Error GoTo errores
    'si p_VerSoloRiesgosNoRetiradosActual="Sí"--->grabo en la tabla ese valor y pongo m_ObjEntorno.VerSoloRiesgosNoRetirados
    
    If p_VerSoloRiesgosNoRetiradosActual <> "Sí" And p_VerSoloRiesgosNoRetiradosActual <> "No" Then
        Exit Function
    End If
    If p_VerSoloRiesgosNoRetiradosActual = "Sí" Then
        m_ObjEntorno.VerSoloRiesgosNoRetirados = EnumSiNo.Sí
    Else
        m_ObjEntorno.VerSoloRiesgosNoRetirados = EnumSiNo.No
    End If
    m_SQL = "SELECT * FROM TbConfiguracionVisionRiesgos " & _
            "WHERE Usuario='" & m_ObjUsuarioConectado.UsuarioRed & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            .AddNew
                .Fields("IDConfiguracion") = DameID("TbConfiguracionVisionRiesgos", "IDConfiguracion", getdb())
                .Fields("Usuario") = m_ObjUsuarioConectado.UsuarioRed
                If m_ObjEntorno.VerRiesgosDescripcion = 0 Then
                    m_ObjEntorno.VerRiesgosDescripcion = EnumSiNo.Sí
                End If
                If m_ObjEntorno.VerRiesgosDescripcion = EnumSiNo.Sí Then
                    .Fields("VerDescripcion") = "Sí"
                Else
                    .Fields("VerDescripcion") = "No"
                End If
                
                .Fields("VerSoloNoRetirados") = p_VerSoloRiesgosNoRetiradosActual
            .Update
           
        Else
            .Edit
                
                .Fields("VerSoloNoRetirados") = p_VerSoloRiesgosNoRetiradosActual
            .Update
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método setVerSoloRiesgosNoRetirados ha producido el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getDirectorioOneDrive(Optional ByRef p_Error As String) As String
    Dim FSO As Object
    Dim carpetaRaiz As Object
    Dim subCarpeta As Object
    Dim rutaEncontrada As String
    Dim encontrado As Boolean
    Dim dato As Variant
    On Error GoTo errores
    ' Crear objeto FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Obtener la carpeta raíz de C:\
    Set carpetaRaiz = FSO.GetFolder("C:\")
    
    ' Inicializar variables
    encontrado = False
    rutaEncontrada = ""
    
    ' Recorrer las subcarpetas en la raíz de C:\
    For Each subCarpeta In carpetaRaiz.SubFolders
        If InStr(1, subCarpeta.Name, "OneDrive", vbTextCompare) > 0 Then
            rutaEncontrada = subCarpeta.Path
            encontrado = True
            Exit For
        
        End If
    Next subCarpeta
    
    ' Mostrar el resultado
    If encontrado Then
        getDirectorioOneDrive = rutaEncontrada
    Else
        rutaEncontrada = Environ("OneDrive")
        If InStr(1, rutaEncontrada, "OneDrive") <> 0 Then
            dato = Split(rutaEncontrada, "OneDrive")
            getDirectorioOneDrive = dato(0) & "OneDrive"
        End If
    End If
    
    ' Liberar objetos
    Set subCarpeta = Nothing
    Set carpetaRaiz = Nothing
    Set FSO = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getDirectorioOneDrive ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function

Private Function getDirectorioOneDriveTelefonicaApps(Optional ByRef p_Error As String) As String
    Dim FSO As Object
    Dim carpeta As String
    Dim m_RutaOneDrive As String
    
    On Error GoTo errores
    ' Crear objeto FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    m_RutaOneDrive = getDirectorioOneDrive(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RutaOneDrive = "" Then
        Exit Function
    End If
    carpeta = m_RutaOneDrive & "\Telefonica\Aplicaciones_dys.TMETF - Aplicaciones PpD\"
    If Not FSO.FolderExists(carpeta) Then
        Exit Function
    End If
    getDirectorioOneDriveTelefonicaApps = carpeta
    
    ' Liberar objetos
   
    
    Set FSO = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getDirectorioOneDriveTelefonicaApps ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function
Private Function getDirectorioOneDriveApps(Optional ByRef p_Error As String) As String
    Dim FSO As Object
    Dim carpeta As String
    Dim m_RutaOneDrive As String
    
    On Error GoTo errores
    ' Crear objeto FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    m_RutaOneDrive = getDirectorioOneDrive(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RutaOneDrive = "" Then
        Exit Function
    End If
    'C:\OneDrive\OneDrive - Telefonica\00LABORAL\Aplicaciones PpD
    carpeta = m_RutaOneDrive & "\OneDrive - Telefonica\00LABORAL\Aplicaciones PpD\"
    If Not FSO.FolderExists(carpeta) Then
        Exit Function
    End If
    getDirectorioOneDriveApps = carpeta
    
    ' Liberar objetos
   
    
    Set FSO = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getDirectorioOneDriveApps ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function


Public Function getRutaAplicacionesLocal(Optional ByRef p_Error As String) As String
    Dim FSO As Object
    Dim m_RutaOneDrive As String
    Dim m_RutaOneDriveTelefonica As String
    
    On Error GoTo errores
    ' Crear objeto FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    m_RutaOneDriveTelefonica = getDirectorioOneDriveTelefonicaApps(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If FSO.FolderExists(m_RutaOneDriveTelefonica) Then
        Set FSO = Nothing
        getRutaAplicacionesLocal = m_RutaOneDriveTelefonica
        Exit Function
    End If
    
    m_RutaOneDrive = getDirectorioOneDrive(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RutaOneDrive = "" Then
        Set FSO = Nothing
        Exit Function
    End If
   
    getRutaAplicacionesLocal = m_RutaOneDrive
    
    ' Liberar objetos
   
    
    Set FSO = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getRutaAplicacionesLocal ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function ObtenerFechaMenor(Fechas As String) As String
    Dim FechaArray() As String
    Dim FechaActual As Date
    Dim FechaMinima As Date
    Dim i As Integer
    
    ' Dividir el string en un array usando la coma como separador
    FechaArray = Split(Fechas, ",")
    
    ' Inicializar la fecha mínima con la primera fecha del array
    FechaMinima = CDate(Trim(FechaArray(0)))
    
    ' Recorrer el array para encontrar la fecha más pequeña
    For i = 1 To UBound(FechaArray)
        If IsDate(Trim(FechaArray(i))) Then
            FechaActual = CDate(Trim(FechaArray(i)))
            If FechaActual < FechaMinima Then
                FechaMinima = FechaActual
            End If
        End If
        
    Next i
    
    ' Devolver la fecha mínima encontrada
    ObtenerFechaMenor = CStr(FechaMinima)
End Function

Public Function RegistrarActualizacionEdicion( _
                                                p_Edicion As Edicion, _
                                                Optional p_FechaSiguientePublicacion As String, _
                                                Optional ByVal p_db As DAO.Database, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
       
    Dim m_SQL As String
    Dim db As DAO.Database
    
    On Error GoTo errores
    If p_Edicion Is Nothing Then
        Exit Function
    End If
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then Err.Raise 1000
    Else
        Set db = p_db
    End If
    m_SQL = "UPDATE TbProyectosEdiciones SET FechaUltimoCambio = Now(), " & _
            "UsuarioUltimoCambio = '" & m_ObjUsuarioConectado.UsuarioRed & "' " & _
            "WHERE IDEdicion=" & p_Edicion.IDEdicion & ";"
    db.Execute m_SQL
    
    p_FechaSiguientePublicacion = SetFechaMaximaPublicacion(p_IDProyecto:=p_Edicion.Proyecto.IDProyecto, _
                                    p_FechaSiguientePublicacion:=p_FechaSiguientePublicacion, _
                                    p_db:=db, _
                                    p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If FormularioAbierto("FormRiesgosGestion") Then
        m_ObjEdicionActiva.UsuarioUltimoCambio = m_ObjUsuarioConectado.UsuarioRed
        m_ObjEdicionActiva.FechaUltimoCambio = Now()
        If IsDate(p_FechaSiguientePublicacion) Then
            m_ObjEdicionActiva.FechaMaxProximaPublicacion = p_FechaSiguientePublicacion
            m_ObjProyectoActivo.FechaMaxProximaPublicacion = p_FechaSiguientePublicacion
        End If
        Form_FormRiesgosGestion.ActualizarEtiquetaUltimoCambio
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistrarActualizacionEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function RegistrarReseteoActualizacionEdicion( _
                                                        p_Edicion As Edicion, _
                                                        Optional ByVal p_db As DAO.Database, _
                                                        Optional ByRef p_Error As String _
                                                        ) As String
    
       
    Dim m_SQL As String
    Dim db As DAO.Database
    
    On Error GoTo errores
    If p_Edicion Is Nothing Then
        Exit Function
    End If
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then Err.Raise 1000
    Else
        Set db = p_db
    End If
    m_SQL = "UPDATE TbProyectosEdiciones SET FechaUltimoCambio =Null, " & _
            "UsuarioUltimoCambio = Null " & _
            "WHERE IDEdicion=" & p_Edicion.IDEdicion & ";"
    db.Execute m_SQL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistrarReseteoActualizacionEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function RegistrarActualizacionRiesgo( _
                                                Optional p_Riesgo As riesgo, _
                                                Optional p_Edicion As Edicion, _
                                                Optional ByVal p_db As DAO.Database, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
       
   
    Dim m_Col As Scripting.Dictionary
    Dim m_Id As Variant
    Dim m_Riesgo As riesgo
    Dim m_IdPlan As Variant
    Dim m_Plan As Object
    
    On Error GoTo errores
    If p_Riesgo Is Nothing And p_Edicion Is Nothing Then
        Exit Function
    End If
    
    If Not p_Riesgo Is Nothing Then
        'If p_Riesgo.CodigoRiesgo = "R003" Then Stop
        With p_Riesgo
            .RefrescarDerivadosTx p_db:=p_db, p_Error:=p_Error
            If p_Error <> "" Then Err.Raise 1000
            If Not .ColPMs Is Nothing Then
                For Each m_IdPlan In .ColPMs
                    Set m_Plan = .ColPMs(m_IdPlan)
                    
                    Set m_Plan = Nothing
                Next
            End If
            If Not .ColPCs Is Nothing Then
                For Each m_IdPlan In .ColPCs
                    Set m_Plan = .ColPCs(m_IdPlan)
                    Set m_Plan = Nothing
                Next
            End If
        End With
        
        
        Exit Function
    End If
    If Not p_Edicion Is Nothing Then
        Set m_Col = p_Edicion.colRiesgos
        For Each m_Id In m_Col
            Set m_Riesgo = m_Col(m_Id)
            RegistrarActualizacionRiesgo = RegistrarActualizacionRiesgo(p_Riesgo:=m_Riesgo, p_db:=p_db, p_Error:=p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set m_Riesgo = Nothing
        Next
        
        Set m_Col = Nothing
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistrarActualizacionRiesgo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function CalcularFechaMaxProximaPublicacion( _
                                                        Optional p_Proyecto As Proyecto, _
                                                        Optional p_Edicion As Edicion, _
                                                        Optional p_FechaPublicacion As String, _
                                                        Optional p_EsUltimaEdicion As String, _
                                                        Optional ByRef p_Error As String _
                                                        ) As String
    
    Dim m_FechaPrevistaCierreCalculada As String
    Dim m_FechaUltimaPublicacion As String
    Dim m_FechaPrevistaCierre As String
    Dim m_FechaRegistroInicial As String
    Dim m_FechaMaximaPublicacion As String
    Dim m_FechaUltimoCambio As String
    Dim m_EdicionUltima As Edicion
    Dim m_Fecha As String
    Dim m_FechaSemanaDespuesUltimoCambio As String
    Dim m_FechaSemanaAntesDeFPrevistaCierre As String
    Dim m_FechaMenor As String
    Dim m_FechaPrevistaCierreMenos15 As String
    Dim m_sFechaMaxProximaPublicacionCalculada As String
    'función que ha de servir para
    '1) en un momento determinado que se obtenga cuál va a ser la fecha proxima
    '2) Justo antes de publicar (con fecha de publicación dada por p_FechaPublicacion cuál va a ser la fecha proxima
    'si se da p_Edicion se ha de dar p_FechaPublicacion y p_EsUltimaEdicion
    On Error GoTo errores
    If p_Proyecto Is Nothing And p_Edicion Is Nothing Then
        p_Error = "No se conoce el proyecto"
        Err.Raise 1000
    End If
    If Not p_Proyecto Is Nothing Then
        m_FechaPrevistaCierreCalculada = p_Proyecto.FechaPrevistaCierreCalculada
        
    Else
        If IsDate(p_Edicion.Proyecto.FechaCierre) Then
            Exit Function
        End If
    End If
    
    If Not p_Edicion Is Nothing Then
        If Not IsDate(p_FechaPublicacion) Then
            p_Error = "Al indicar la edición se ha de decir la fecha ultima publicacion"
            Err.Raise 1000
        End If
        If IsDate(p_Edicion.FechaPublicacion) Then
            p_Error = "Al indicar la edición No puede estar publicada"
            Err.Raise 1000
        End If
        If p_EsUltimaEdicion <> "Sí" And p_EsUltimaEdicion <> "No" Then
            p_Error = "Se ha de indicar si es la última edición"
            Err.Raise 1000
        End If
        If p_EsUltimaEdicion = "Sí" Then
            Exit Function
        End If
        If CDate(p_FechaPublicacion) > CDate(Date) Then
            p_Error = "La fecha de publicación no puede ser anterior a hoy"
            Err.Raise 1000
        End If
         m_FechaMaximaPublicacion = DateAdd("m", m_ObjEntorno.JPMesesAvisoEntreEdiciones, CDate(p_FechaPublicacion))
        'estamos en la última edición y solo se va a tener en cuenta la fecha fin prevista y si no lo estipulado
        m_FechaPrevistaCierreCalculada = p_Edicion.Proyecto.FechaPrevistaCierreCalculada
        If Not IsDate(m_FechaPrevistaCierreCalculada) Then
            CalcularFechaMaxProximaPublicacion = m_FechaMaximaPublicacion
            Exit Function
        End If
        If CDate(m_FechaPrevistaCierreCalculada) < Date Then
            CalcularFechaMaxProximaPublicacion = m_FechaMaximaPublicacion
            Exit Function
        End If
        m_FechaPrevistaCierreMenos15 = CStr(DateAdd("d", 15, CDate(m_FechaPrevistaCierreCalculada)))
'        If CDate(Date) > CDate(m_FechaPrevistaCierreMenos15) Then
'            CalcularFechaMaxProximaPublicacion = CStr(DateAdd("d", 15, CDate(Date)))
'            Exit Function
'        End If
        
        If CDate(m_FechaMaximaPublicacion) > CDate(m_FechaPrevistaCierreMenos15) Then
            CalcularFechaMaxProximaPublicacion = m_FechaPrevistaCierreMenos15
            Exit Function
        End If
        CalcularFechaMaxProximaPublicacion = m_FechaMaximaPublicacion
        Exit Function
    End If
    
    'ya estamos en el caso tipico en el que no le hemos puesto nada en p_FechaPublicacion
    Set m_EdicionUltima = p_Proyecto.EdicionUltima
    If m_EdicionUltima Is Nothing Then
        p_Error = "No se conoce la última edición"
        Err.Raise 1000
    End If
    If p_Proyecto.EdicionUltimaPublicada Is Nothing Then
        'se ha de publicar la primera semana de dar de alta
        m_sFechaMaxProximaPublicacionCalculada = DateAdd("d", 7, CDate(p_Proyecto.fechaRegistroInicial))
        If CDate(Date) > CDate(m_sFechaMaxProximaPublicacionCalculada) Then
            CalcularFechaMaxProximaPublicacion = CStr(DateAdd("d", 7, CDate(Date)))
            Exit Function
        End If
        CalcularFechaMaxProximaPublicacion = m_sFechaMaxProximaPublicacionCalculada
        Exit Function
    End If
    m_FechaUltimoCambio = m_EdicionUltima.FechaUltimoCambio
    If IsDate(m_FechaUltimoCambio) Then
        m_FechaSemanaDespuesUltimoCambio = DateAdd("d", 7, CDate(m_FechaUltimoCambio))
    End If
    
    m_FechaPrevistaCierre = m_FechaPrevistaCierreCalculada
    m_FechaRegistroInicial = p_Proyecto.fechaRegistroInicial
    m_FechaUltimaPublicacion = p_Proyecto.EdicionUltimaPublicada.FechaPublicacion
    If Not IsDate(m_FechaUltimaPublicacion) Then
        
        m_Fecha = DateAdd("d", m_ObjEntorno.JPDiasPreviosParaElAviso, CDate(m_FechaRegistroInicial))
        m_FechaMenor = ObtenerFechaMenor(m_Fecha & "," & m_FechaSemanaDespuesUltimoCambio)
        
    Else
        m_Fecha = DateAdd("m", m_ObjEntorno.JPMesesAvisoEntreEdiciones, CDate(m_FechaUltimaPublicacion))
        If IsDate(m_FechaPrevistaCierre) Then
            m_FechaSemanaAntesDeFPrevistaCierre = DateAdd("d", 15, CDate(m_FechaPrevistaCierre))
        End If
        m_FechaMenor = ObtenerFechaMenor(m_Fecha & "," & m_FechaSemanaAntesDeFPrevistaCierre & "," & m_FechaSemanaDespuesUltimoCambio)
       
        
    End If
'    If CDate(Date) > CDate(m_FechaMenor) Then
'        CalcularFechaMaxProximaPublicacion = DateAdd("d", 15, CDate(Date))
'    Else
'        CalcularFechaMaxProximaPublicacion = m_FechaMenor
'    End If
    CalcularFechaMaxProximaPublicacion = m_FechaMenor
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CalcularFechaMaxProximaPublicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function getNodoRiesgoDelSeleccionado( _
                                                ByVal p_Nodo As MSComctlLib.Node, _
                                                Optional ByRef p_Error As String _
                                                ) As MSComctlLib.Node
    Dim m_TipoNodo As EnumTipoNodo
    On Error GoTo errores

    
    If p_Nodo Is Nothing Then
        Exit Function
    End If
    m_TipoNodo = getTipoNodo(p_Nodo:=p_Nodo, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TipoNodo = Empty Then
        Exit Function
    End If
    If m_TipoNodo = EnumTipoNodo.riesgo Then
        Set getNodoRiesgoDelSeleccionado = p_Nodo
        Exit Function
    End If
    If m_TipoNodo = EnumTipoNodo.PM Or m_TipoNodo = EnumTipoNodo.PC Then
        Set getNodoRiesgoDelSeleccionado = p_Nodo.Parent
        Exit Function
    End If
    
    If m_TipoNodo = EnumTipoNodo.PMA Or m_TipoNodo = EnumTipoNodo.PCA Then
        Set getNodoRiesgoDelSeleccionado = p_Nodo.Parent.Parent
        Exit Function
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo getNodoRiesgoDelSeleccionado ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getTipoNodo( _
                                ByVal p_Nodo As MSComctlLib.Node, _
                                Optional ByRef p_Error As String _
                                ) As EnumTipoNodo
    
    On Error GoTo errores

    
    If p_Nodo Is Nothing Then
        Exit Function
    End If
    
    If InStr(1, p_Nodo.Key, "EDICION|") <> 0 Then
        getTipoNodo = EnumTipoNodo.Edicion
        Exit Function
    End If
    If InStr(1, p_Nodo.Key, "RIESGO|") <> 0 Then
        getTipoNodo = EnumTipoNodo.riesgo
        Exit Function
    End If
    If InStr(1, p_Nodo.Key, "PM|") <> 0 Then
        getTipoNodo = EnumTipoNodo.PM
        Exit Function
    End If
    If InStr(1, p_Nodo.Key, "PC|") <> 0 Then
        getTipoNodo = EnumTipoNodo.PC
        Exit Function
    End If
    If InStr(1, p_Nodo.Key, "PMACCION|") <> 0 Then
        getTipoNodo = EnumTipoNodo.PMA
        Exit Function
    End If
    If InStr(1, p_Nodo.Key, "PCACCION|") <> 0 Then
        getTipoNodo = EnumTipoNodo.PCA
        Exit Function
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo getTipoNodo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getIDNodo( _
                            ByVal p_Key As String, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim dato As Variant
    
    On Error GoTo errores

    If InStr(1, p_Key, "|") = 0 Then
        Exit Function
    End If
    dato = Split(p_Key, "|")
    If UBound(dato) < 1 Then
        Exit Function
    End If
    getIDNodo = dato(1)
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo getIDNodo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getObjeto( _
                            p_Id As String, _
                            p_EnumTipoNodo As EnumTipoNodo, _
                            Optional ByRef p_Error As String _
                            ) As Object
    
    
    
    On Error GoTo errores
    
    If p_EnumTipoNodo = EnumTipoNodo.Edicion Then
        Set getObjeto = Constructor.getEdicion(p_IDEdicion:=p_Id, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_EnumTipoNodo = EnumTipoNodo.riesgo Then
        Set getObjeto = Constructor.getRiesgo(p_IDRiesgo:=p_Id, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_EnumTipoNodo = EnumTipoNodo.PM Then
        Set getObjeto = Constructor.getPM(p_IDPM:=p_Id, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_EnumTipoNodo = EnumTipoNodo.PC Then
        Set getObjeto = Constructor.getPC(p_IDPC:=p_Id, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_EnumTipoNodo = EnumTipoNodo.PMA Then
        Set getObjeto = Constructor.getPMAccion(p_IDPMAccion:=p_Id, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_EnumTipoNodo = EnumTipoNodo.PCA Then
        Set getObjeto = Constructor.getPCAccion(p_IDPCAccion:=p_Id, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getObjeto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function
Public Sub ResetAvanceCounter()
    ' Responsabilidad: Pone a cero el contador de pasos del splash.
    ' Debe ser llamado desde el evento Form_Open del frmSplash.
    s_contadorPasos = 0
End Sub



Public Sub GestionarRibbon(ByVal mostrar As Boolean)
    ' RESPONSABILIDAD: Muestra u oculta la cinta de opciones de Access.
    On Error Resume Next ' Si hay algún problema, no debe detener el arranque.
    
    If mostrar Then
        DoCmd.ShowToolbar "Ribbon", acToolbarYes
    Else
        DoCmd.ShowToolbar "Ribbon", acToolbarNo
    End If
End Sub







