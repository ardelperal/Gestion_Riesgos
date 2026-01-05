
Option Compare Database
Option Explicit

Public Function PublicarEdicionTransaccional( _
                            p_Edicion As Edicion, _
                            Optional p_FechaCierre As String, _
                            Optional p_RegenerandoCambiosDesdeElInicio As EnumSiNo, _
                            Optional p_FechaSiguientePublicacion As String, _
                            Optional p_fechaRef As String, _
                            Optional p_CorreoRAC As String, _
                            Optional p_ConCorreoAlRAC As EnumSiNo, _
                            Optional p_EnviarCorreo As EnumSiNo, _
                            Optional p_ActualizarTareasCalidad As EnumSiNo, _
                            Optional ByRef p_Error As String _
                            ) As String
     
    
    Dim m_ObjDocumento As Documento
    
    Dim m_URLInicial As String
    Dim m_URLFinal As String
    Dim m_NombreArchivo As String
    Dim m_DocumentoGenerado As Boolean
    Dim m_DocumentoCopiado As Boolean
    Dim m_ColCambiosEntreEdiciones As Scripting.Dictionary
    Dim m_CodigoDocumento As String
    Dim m_EdicionDocumento As String
    
    Dim m_IDDocumento As String
    Dim m_ObjEdicionAlInicio As Edicion
    Dim m_ObjEdicionGenerada As Edicion
    Dim m_Proyecto As Proyecto
    Dim m_Publicable As EnumSiNo
    Dim m_URLInforme As String
    Dim m_PublicacionLog As PublicacionLog
    
    Dim ws As DAO.Workspace
    Dim m_EnTransaccion As Boolean
    
    On Error GoTo errores
    
    ' Inicializar Workspace
    Set ws = DBEngine.Workspaces(0)
    m_EnTransaccion = False
    
    Set m_ObjEdicionAlInicio = Constructor.getEdicion(p_IDEdicion:=p_Edicion.IDEdicion)
    If Not IsDate(p_fechaRef) Then
        p_fechaRef = Date
    End If
    
    
    Avance "Publicación...Comprobando publicabilidad..."
    If p_Edicion.Publicable = EnumSiNo.No Then
        m_URLInforme = GenerarInformePublicabilidadEdicionInteractivoHTML(p_Edicion:=p_Edicion, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_URLHTMLActivo = m_URLInforme
        p_Error = "La Edición no es publicable"
        Err.Raise 1000
    End If
    
    ' INICIO DE TRANSACCIÓN
    If p_ConCorreoAlRAC = Empty Then
        p_ConCorreoAlRAC = EnumSiNo.No
    End If
    If p_EnviarCorreo = Empty Then
        p_EnviarCorreo = EnumSiNo.Sí
    End If
    If p_ActualizarTareasCalidad = Empty Then
        p_ActualizarTareasCalidad = EnumSiNo.Sí
    End If

    ws.BeginTrans
    m_EnTransaccion = True
    
    If p_FechaCierre <> "" Then
        Avance "Publicación...Cerrando riesgos activos o detectados"
        p_Edicion.CerrarRiesgos p_Error:=p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    
    End If
    Set m_Proyecto = p_Edicion.Proyecto
    If p_CorreoRAC <> "" Then
        If p_CorreoRAC <> m_Proyecto.CorreoRAC Then
            m_Proyecto.RegistrarCorreoRAC p_CorreoRAC, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    
    '------------------
    'Registrando cambios de la edición con la anterior
    '----------------------------------------------------
    
    
    Avance "Publicación...Generando Informe..."
    m_URLInicial = GenerarInforme(p_Edicion:=p_Edicion, p_EnExcel:=EnumSiNo.No, _
                                p_FechaCierre:=p_FechaCierre, p_FechaPublicacion:=p_fechaRef, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_DocumentoGenerado = True ' Marcar que se ha generado para limpieza en caso de error
    
    If p_FechaCierre <> "" Then
        'se pone el estado de cada riesgo como cerrado
        p_Edicion.Cerrar p_Error:=p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Avance "Publicación...Copiando Archivo a repositorio final..."
    With p_Edicion
        .FechaPublicacion = CStr(p_fechaRef)
        If .FechaPreparadaParaPublicar = "" Then
            .FechaPreparadaParaPublicar = .FechaPublicacion
        End If
        .Elaborado = .Proyecto.Elaborado
        .Revisado = .Proyecto.Revisado
        .Aprobado = .Proyecto.Aprobado
        
        If m_Proyecto.Juridica = "TdE" Then
            Set m_ObjDocumento = RegistrarEnAGEDO(p_ObjEdicionAPublicar:=p_Edicion, p_URLInicial:=m_URLInicial, p_Error:=p_Error)
'            If p_Error <> "" Then
'                Err.Raise 1000
'            End If
            p_Error = ""
            If Not m_ObjDocumento Is Nothing Then
                .IDDocumentoAGEDO = m_ObjDocumento.IDDocumento
                m_URLFinal = m_ObjDocumento.URLAdjunto
            End If
            
            
        Else
            m_NombreArchivo = FSO.GetFileName(m_URLInicial)
            .NombreArchivoInforme = m_NombreArchivo
            m_URLFinal = m_ObjEntorno.URLDirectorioDocumentacion & m_NombreArchivo
            FSO.CopyFile m_URLInicial, m_URLFinal
            m_DocumentoCopiado = True ' Marcar que se ha copiado
        End If
        
        
        .PropuestaRechazadaPorCalidadFecha = ""
        .PropuestaRechazadaPorCalidadMotivo = ""
        .Registrar p_ObjEdicionAlInicio:=m_ObjEdicionAlInicio, p_ProvieneDePublicacion:=EnumSiNo.Sí, p_Error:=p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If p_FechaCierre = "" Then
            Avance "Publicación...Generando nueva Edición a partir de actual..."
            Set m_ObjEdicionGenerada = GenerarEdicionNuevaAPartirDeAnterior(p_IDEdicion:=p_Edicion.IDEdicion, _
                                      p_FechaSiguientePublicacion:=p_FechaSiguientePublicacion, p_Error:=p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
        Else
            Avance "Publicación...Cerrando Edición..."
            m_Proyecto.FechaCierre = CStr(p_FechaCierre)
            m_Proyecto.Cerrar p_Error:=p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
        End If
        
    End With
    
    Set m_PublicacionLog = New PublicacionLog
    With m_PublicacionLog
        .IDEdicion = p_Edicion.IDEdicion
        .FechaPublicacion = p_Edicion.FechaPublicacion
        .UsuarioFechaPublicacion = m_ObjUsuarioConectado.Nombre
        .Registrar p_Error:=p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Set m_PublicacionLog = Nothing
    
    If p_ActualizarTareasCalidad = EnumSiNo.Sí Then
        EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If p_EnviarCorreo = EnumSiNo.Sí Then
        If m_Proyecto.ParaInformeAvisos = "S¡" Then
            EnvioCorreoNuevaPublicacion m_Proyecto, m_URLFinal, p_ConCorreoAlRAC, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    
    ' FIN DE TRANSACCIÓN
    ws.CommitTrans
    m_EnTransaccion = False
    
    PublicarEdicionTransaccional = m_URLFinal
    Exit Function
errores:
    If m_EnTransaccion Then
        ws.Rollback
    End If
    
    If Err.Number <> 1000 Then
        p_Error = "El método PublicacionTransaccional.PublicarEdicionTransaccional ha producido el error : " & vbNewLine & Err.Description
    End If
    On Error Resume Next
    If m_DocumentoGenerado = True Then
        ' Si se generó el inicial, intentar borrarlo (aunque GenerarInforme suele sobreescribir)
        ' Pero si falló después, queremos limpiar?
        ' Publicar original borra m_URLFinal si m_DocumentoCopiado es true o si m_DocumentoGenerado es true
        
        If m_URLInicial <> "" Then
            If FSO.FileExists(m_URLInicial) Then FSO.DeleteFile m_URLInicial, True
        End If
        If m_DocumentoCopiado = True Then
            If m_URLFinal <> "" Then
                If FSO.FileExists(m_URLFinal) Then FSO.DeleteFile m_URLFinal, True
            End If
        End If
        
        ' Si era AGEDO, m_ObjDocumento.Borrar
        If Not m_ObjDocumento Is Nothing Then
             m_ObjDocumento.Borrar p_Error
        End If
    End If
    
End Function
