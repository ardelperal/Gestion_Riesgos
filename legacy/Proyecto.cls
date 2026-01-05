

Option Compare Database
Option Explicit

Public Event NoHayCambios()
Public Event HayErrores()



Public IDProyecto As String
Public IDExpediente As String
Public Proyecto As String
Public Juridica As String
Public NombreProyecto As String
Public Cliente As String
Public FechaPrevistaCierre As String
Public FechaCierre As String
Public fechaRegistroInicial As String
Public Elaborado As String
Public Revisado As String
Public Aprobado As String
Public CodigoDocumento As String
Public ParaInformeAvisos As String
Public FechaFirmaContrato As String
Public RiesgosDeLaOferta As String
Public RiesgosDelSubContratista As String
Public NombreUsuarioCalidad As String
Public EnUTE As String
Public FechaMaxProximaPublicacion As String
Public RequiereRiesgoDeBiblioteca As String
Public CorreoRAC As String
Public Ordinal As String
Public CadenaNombreAutorizados As String
Public NombreParaNodo As String

Private m_sIDProyectoCalculado As String
Private m_sProyectoCalculado As String
Private m_sJuridicaCalculada As String
Private m_sNombreProyectoCalculado As String
Private m_sFechaPrevistaCierreCalculada As String
Private m_sClienteCalculado As String
Private m_sCodigoDocumentoCalculado As String
Private m_sFechaFirmaContratoCalculada As String
Private m_sNombreUsuarioCalidadCalculado As String
Private m_sEnUTECalculado As String
Private m_sFechaMaxProximaPublicacionCalculada As String
Private m_sFechaUltimaPublicacion As String
Private m_eRequiereRiesgoDeBibliotecaCalculado As EnumSiNo
Private m_sRequiereRiesgoDeBibliotecaCalculadoTexto As String
Private m_sCorreoRACCalculado As String
Private m_sOrdinalCalculado As String


Private m_objElaboradoObj As Usuario
Private m_objExpediente As Expediente
Private m_objRevisadoObj As Usuario
Private m_objAprobadoObj As Usuario
Private m_objColEdiciones As Scripting.Dictionary
Private m_ObjEdicionActiva As Edicion
Private m_ObjEdicionUltima As Edicion
Private m_ObjEdicionUltimaPublicada As Edicion
Private m_ObjEdicionPrimera As Edicion

Private m_ObjColAutorizados As Scripting.Dictionary


Private m_sUsuarioCalidadRowSource As String

Private m_ObjUsuarioCalidad As Usuario

Private m_objUsuarioElaboradoObj As Usuario
Private m_objUsuarioRevisadoObj As Usuario
Private m_objUsuarioAprobadoObj As Usuario

Private m_ObjColAnexos As Scripting.Dictionary
Private m_sIDRiesgoDELProximoAnexo As String
Private m_eOrdenarColeEdicionesAscendentemente As EnumSiNo

Private m_sDiasParaElSiguienteAvisoAJP As String
Private m_eUsuarioConectadoAutorizado As EnumSiNo
Private m_sURLUltimaPublicacion As String
Private m_ObjEdicionUltimaNoActiva As Edicion
Private m_eEsActivo As EnumSiNo
Private m_eTieneAlMenosUnJP As EnumSiNo
Private m_eTieneAlMenosUnEnvioCorreo As EnumSiNo
Private m_ObjColRiesgosOferta As Scripting.Dictionary
Private m_objColRiesgosSubContratista As Scripting.Dictionary
Private m_objColRiesgosPedidos As Scripting.Dictionary
Private m_objColRiesgosExternos As Scripting.Dictionary
Private m_eTieneEdiciones As EnumSiNo
Private m_eTieneRiesgos As EnumSiNo
Private m_eTieneRiesgosExt As EnumSiNo
Private m_eTieneRiesgosDeOferta As EnumSiNo
Private m_eTieneRiesgosDeSubContratista As EnumSiNo
Private m_eTieneRiesgosDePedidos As EnumSiNo
Private m_ObjColCambios As Scripting.Dictionary
Private m_sDestinatariosCorreo As String
Private m_objColAnexosTotales As Scripting.Dictionary
Private m_objColRiesgosDeJP As Scripting.Dictionary

Private m_eAPuntoDeCaducar As EnumSiNo
Private m_eCaducado As EnumSiNo
Private m_objColEdicionesPublicadas As Scripting.Dictionary
Private m_sCorreoResponsableCalidad As String
Private m_objColSuministradores As Scripting.Dictionary
Private m_sHTMLProyecto As String
Private m_sHTMLTablaProyectoDatos As String
Private m_sHTMLTablaEdiciones As String
Private m_sHTMLTablaProyectoAutorizados As String
Private m_sCadenaCorreoAutorizados As String
Private m_sCadenaNombreAutorizadosCalculados As String
Private m_objColCamposParaCambios As Collection
Private m_eAlmenosUnaEdicionPublicada As EnumSiNo

Private m_sNombreParaNodoCalculado As String
Private m_sNombreUsuarioCalidadParaNodo As String
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get NombreUsuarioCalidadParaNodo() As String
    
    Dim dato As Variant
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sNombreUsuarioCalidadParaNodo <> "" Then
        NombreUsuarioCalidadParaNodo = m_sNombreUsuarioCalidadParaNodo
        Exit Property
    End If
    If Me.NombreUsuarioCalidad = "" Then
        Exit Property
    End If
    dato = Split(Me.NombreUsuarioCalidad, " ")
    m_sNombreUsuarioCalidadParaNodo = dato(0)
    
    
    NombreUsuarioCalidadParaNodo = m_sNombreUsuarioCalidadParaNodo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.NombreUsuarioCalidadParaNodo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get NombreParaNodoCalculado() As String
    
    Dim m_NombreParaNodoCalculadoInicial As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sNombreParaNodoCalculado <> "" Then
        NombreParaNodoCalculado = m_sNombreParaNodoCalculado
        Exit Property
    End If
    m_NombreParaNodoCalculadoInicial = Me.NombreProyecto
    
'    If Me.Expediente Is Nothing Then
'        m_NombreParaNodoCalculadoInicial = Me.NombreParaNodoCalculado
'    Else
'        If Me.Expediente.Nemotecnico <> "" Then
'            m_NombreParaNodoCalculadoInicial = Me.Expediente.Nemotecnico
'        Else
'            If Len(Me.Expediente.Tipo) > 20 Then
'                m_NombreParaNodoCalculadoInicial = Left(Me.Expediente.Titulo, 20) & " ..."
'            Else
'                m_NombreParaNodoCalculadoInicial = Me.Expediente.Titulo
'            End If
'
'        End If
'    End If
    If Me.ParaInformeAvisos = "Sí" Then
        
        m_sNombreParaNodoCalculado = m_NombreParaNodoCalculadoInicial
    Else
        m_sNombreParaNodoCalculado = m_NombreParaNodoCalculadoInicial & " (" & Format(Me.Ordinal, "00") & ")"
    End If
  
    
    NombreParaNodoCalculado = m_sNombreParaNodoCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.NombreParaNodoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get NombreUsuarioCalidadCalculado() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sNombreUsuarioCalidadCalculado <> "" Then
        NombreUsuarioCalidadCalculado = m_sNombreUsuarioCalidadCalculado
        Exit Property
    End If
    If Me.Expediente Is Nothing Then
        Exit Property
    End If
    If Not Me.Expediente.ResponsableCalidad Is Nothing Then
        m_sNombreUsuarioCalidadCalculado = Me.Expediente.ResponsableCalidad.Nombre
    End If
    
    
    NombreUsuarioCalidadCalculado = m_sNombreUsuarioCalidadCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.NombreUsuarioCalidadCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EnUTECalculado() As String
    
    Dim dato As Variant
    Dim m_Juridica As Variant
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sEnUTECalculado <> "" Then
        EnUTECalculado = m_sEnUTECalculado
        Exit Property
    End If
    If Me.Juridica = "" Then
        m_sEnUTECalculado = "No"
        EnUTECalculado = m_sEnUTECalculado
        Exit Property
    End If
    If InStr(1, Me.Juridica, "|") = 0 Then
        m_sEnUTECalculado = "No"
        EnUTECalculado = m_sEnUTECalculado
        Exit Property
    End If
    dato = Split(Me.Juridica, "|")
    For Each m_Juridica In dato
        If CStr(m_Juridica) <> "TSOL" And CStr(m_Juridica) <> "TME" And CStr(m_Juridica) <> "TDE" Then
            m_sEnUTECalculado = "Sí"
            EnUTECalculado = m_sEnUTECalculado
            Exit Property
        End If
    Next
    
    m_sEnUTECalculado = "No"
    
    EnUTECalculado = m_sEnUTECalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.EnUTECalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get FechaPrevistaCierreCalculada() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sFechaPrevistaCierreCalculada <> "" Then
        FechaPrevistaCierreCalculada = m_sFechaPrevistaCierreCalculada
        Exit Property
    End If
    If Me.Expediente Is Nothing Then
        Exit Property
    End If
    m_sFechaPrevistaCierreCalculada = Me.Expediente.FechaFinContrato
    
    FechaPrevistaCierreCalculada = m_sFechaPrevistaCierreCalculada
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.FechaPrevistaCierreCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ClienteCalculado() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sClienteCalculado <> "" Then
        ClienteCalculado = m_sClienteCalculado
        Exit Property
    End If
    If Me.Expediente Is Nothing Then
        Exit Property
    End If
    If Me.Expediente.OrganoContratacion Is Nothing Then
        Exit Property
    End If
    m_sClienteCalculado = Me.Expediente.OrganoContratacion.OrganoContratacion
    
    ClienteCalculado = m_sClienteCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.ClienteCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get NombreProyectoCalculado() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sNombreProyectoCalculado <> "" Then
        NombreProyectoCalculado = m_sNombreProyectoCalculado
        Exit Property
    End If
    If Me.Expediente Is Nothing Then
        Exit Property
    End If
    With Me.Expediente
        If .Nemotecnico <> "" Then
            If IsNumeric(.Ordinal) Then
                If .Ordinal = "1" Then
                    m_sNombreProyectoCalculado = .Nemotecnico
                Else
                    m_sNombreProyectoCalculado = .Nemotecnico & " (GR" & Format(.Ordinal, "00") & ")"
                End If
            Else
                m_sNombreProyectoCalculado = .Nemotecnico
            End If
            
            
        Else
            
            If .Titulo <> "" Then
                If IsNumeric(.Ordinal) Then
                    If .Ordinal = "1" Then
                        m_sNombreProyectoCalculado = .Titulo
                    Else
                       m_sNombreProyectoCalculado = .Titulo & " (GR" & Format(.Ordinal, "00") & ")"
                    End If
                Else
                    m_sNombreProyectoCalculado = .Titulo
                End If
                
            End If
        End If
    End With
    
    
    NombreProyectoCalculado = m_sNombreProyectoCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.NombreProyectoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get JuridicaCalculada() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sJuridicaCalculada <> "" Then
        JuridicaCalculada = m_sJuridicaCalculada
        Exit Property
    End If
    If Me.Expediente Is Nothing Then
        Exit Property
    End If
    m_sJuridicaCalculada = Me.Expediente.CadenaJuridicas
    
    JuridicaCalculada = m_sJuridicaCalculada
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.JuridicaCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get ProyectoCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sProyectoCalculado <> "" Then
        ProyectoCalculado = m_sProyectoCalculado
        Exit Property
    End If
    If Me.Expediente Is Nothing Then
        Exit Property
    End If
    If Me.Expediente.CodExp <> "" Then
        m_sProyectoCalculado = Me.Expediente.CodExp
        ProyectoCalculado = m_sProyectoCalculado
        Exit Property
    End If
    If Me.Expediente.CodigoActividad <> "" Then
        m_sProyectoCalculado = Me.Expediente.CodigoActividad
        ProyectoCalculado = m_sProyectoCalculado
        Exit Property
    End If
    If Me.Expediente.CodExpLargo <> "" Then
        m_sProyectoCalculado = Me.Expediente.CodExpLargo
        ProyectoCalculado = m_sProyectoCalculado
        Exit Property
    End If
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.ProyectoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get FechaFirmaContratoCalculada() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sFechaFirmaContratoCalculada <> "" Then
        FechaFirmaContratoCalculada = m_sFechaFirmaContratoCalculada
        Exit Property
    End If
    If Me.Expediente Is Nothing Then
        Exit Property
    End If
    m_sFechaFirmaContratoCalculada = Me.Expediente.FechaFinContrato
    FechaFirmaContratoCalculada = m_sFechaFirmaContratoCalculada
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.FechaFirmaContratoCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get CorreoRACCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sCorreoRACCalculado <> "" Then
        CorreoRACCalculado = m_sCorreoRACCalculado
        Exit Property
    End If
    If Me.Expediente Is Nothing Then
        Exit Property
    End If
    m_sCorreoRACCalculado = Me.Expediente.CadenaCorreoRACs
    CorreoRACCalculado = m_sCorreoRACCalculado
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.CorreoRACCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get CodigoDocumentoCalculado() As String
    
    Dim m_CodDocPorParteExp As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sCodigoDocumentoCalculado <> "" Then
        CodigoDocumentoCalculado = m_sCodigoDocumentoCalculado
        Exit Property
    End If
    If Me.Expediente Is Nothing Then
        Exit Property
    End If
    m_CodDocPorParteExp = Me.Expediente.CodigoDocParteExpediente
    m_Error = Me.Expediente.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sCodigoDocumentoCalculado = getCodigoDocumento(m_CodDocPorParteExp, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CodigoDocumentoCalculado = m_sCodigoDocumentoCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.CodigoDocumentoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Expediente() As Expediente

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objExpediente Is Nothing Then
        Set Expediente = m_objExpediente
        Exit Property
    End If
    If Me.IDExpediente = "" Then
        Exit Property
    End If
    Set m_objExpediente = Constructor.getExpediente(p_IDExpediente:=Me.IDExpediente, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Expediente = m_objExpediente
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.Expediente ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property
Public Property Set Expediente(p_Valor As Variant)
    Set m_objExpediente = p_Valor
End Property
Public Property Get CorreoResponsableCalidad() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sCorreoResponsableCalidad <> "" Then
        CorreoResponsableCalidad = m_sCorreoResponsableCalidad
        Exit Property
    End If
    If Me.UsuarioCalidad Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_sCorreoResponsableCalidad = Me.UsuarioCalidad.CorreoUsuario
    CorreoResponsableCalidad = m_sCorreoResponsableCalidad
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.CorreoResponsableCalidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property




Public Property Get IDProyectoCalculado() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sIDProyectoCalculado = DameID("TbProyectos", "IDProyecto", , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDProyectoCalculado = m_sIDProyectoCalculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.IDProyectoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get OrdinalCalculado() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sOrdinalCalculado = getOrdinalSiguiente(p_IDExpediente:=Me.IDExpediente, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    OrdinalCalculado = m_sOrdinalCalculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.OrdinalCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property


Public Property Get RequiereRiesgoDeBibliotecaCalculado() As EnumSiNo

    On Error GoTo errores
    
    m_Error = ""
    Me.Error = ""
    
    If Me.RequiereRiesgoDeBiblioteca <> "Sí" And Me.RequiereRiesgoDeBiblioteca <> "No" Then
        Me.RequiereRiesgoDeBiblioteca = "No"
    End If
    If Me.RequiereRiesgoDeBiblioteca = "Sí" Then
        m_eRequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí
    Else
        m_eRequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No
    End If
    RequiereRiesgoDeBibliotecaCalculado = m_eRequiereRiesgoDeBibliotecaCalculado
     Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.RequiereRiesgoDeBibliotecaCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property


Public Property Let RequiereRiesgoDeBibliotecaCalculado(ByVal eNewValue As EnumSiNo)

     
    If eNewValue = EnumSiNo.Sí Then
        m_eRequiereRiesgoDeBibliotecaCalculado = eNewValue
        Me.RequiereRiesgoDeBiblioteca = "Sí"
    ElseIf eNewValue = EnumSiNo.No Then
        m_eRequiereRiesgoDeBibliotecaCalculado = eNewValue
        Me.RequiereRiesgoDeBiblioteca = "No"
    Else
        m_eRequiereRiesgoDeBibliotecaCalculado = 0
        Me.RequiereRiesgoDeBibliotecaCalculado = ""
    End If

End Property
Public Property Get RequiereRiesgoDeBibliotecaCalculadoTexto() As String
    
    On Error GoTo errores
    
    m_Error = ""
    Me.Error = ""
    If m_sRequiereRiesgoDeBibliotecaCalculadoTexto <> "" Then
        RequiereRiesgoDeBibliotecaCalculadoTexto = m_sRequiereRiesgoDeBibliotecaCalculadoTexto
        Exit Property
    End If
    If Me.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
        m_sRequiereRiesgoDeBibliotecaCalculadoTexto = "Sí"
    ElseIf Me.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
        m_sRequiereRiesgoDeBibliotecaCalculadoTexto = "No"
    End If
    
    RequiereRiesgoDeBibliotecaCalculadoTexto = m_sRequiereRiesgoDeBibliotecaCalculadoTexto
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.RequiereRiesgoDeBibliotecaCalculadoTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Function FechaPrevistaCierreGrabar( _
                                                    Optional ByRef p_Error As String _
                                                     ) As String
    
    Dim m_FechaPrevistaCierre As String
    
    On Error GoTo errores
    m_FechaPrevistaCierre = Me.FechaPrevistaCierreCalculada
    
    FechaPrevistaCierreGrabarTransaccional Me, , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    FechaPrevistaCierreGrabar = m_FechaPrevistaCierre
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo FechaPrevistaCierreGrabar ha devuelto el error: " & Err.Description
    End If
End Function

Public Property Get FechaMaxProximaPublicacionCalculada() As String
    
   
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    m_sFechaMaxProximaPublicacionCalculada = CalcularFechaMaxProximaPublicacion(p_Proyecto:=Me, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
   
    
    FechaMaxProximaPublicacionCalculada = m_sFechaMaxProximaPublicacionCalculada
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.FechaMaxProximaPublicacionCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let FechaMaxProximaPublicacionCalculada(ByVal sNewValue As String)

    m_sFechaMaxProximaPublicacionCalculada = sNewValue

End Property
Public Property Get FechaUltimaPublicacion() As String
    
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If Me.EdicionUltimaPublicada Is Nothing Then
        Exit Property
    End If
    
    
    m_sFechaUltimaPublicacion = Me.EdicionUltimaPublicada.FechaPublicacion
    
   
    
    FechaUltimaPublicacion = m_sFechaUltimaPublicacion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.FechaUltimaPublicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let FechaUltimaPublicacion(ByVal sNewValue As String)

    m_sFechaUltimaPublicacion = sNewValue

End Property
Public Property Get APuntoDeCaducar() As EnumSiNo
    
    Dim m_FechaRef As String
    Dim m_FechaMax As String
    Dim m_fechaHoy As String
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_eAPuntoDeCaducar <> 0 Then
        APuntoDeCaducar = m_eAPuntoDeCaducar
        Exit Property
    End If
    If Me.FechaCierre <> "" Then
        m_eAPuntoDeCaducar = EnumSiNo.No
        Exit Property
    End If
    m_fechaHoy = Date
    'm_fechaRef = "12/12/2023"
    
    m_FechaMax = Me.FechaMaxProximaPublicacionCalculada
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If IsDate(m_FechaMax) Then
        m_FechaRef = DateAdd("d", -7, CDate(m_FechaMax))
        If CDate(m_fechaHoy) < CDate(m_FechaMax) And CDate(m_fechaHoy) >= CDate(m_FechaRef) Then
        
            m_eAPuntoDeCaducar = EnumSiNo.Sí
        Else
            m_eAPuntoDeCaducar = EnumSiNo.No
        End If
    End If
    
    APuntoDeCaducar = m_eAPuntoDeCaducar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.APuntoDeCaducar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let APuntoDeCaducar(ByVal eNewValue As EnumSiNo)

    m_eAPuntoDeCaducar = eNewValue

End Property

Public Property Get Caducado() As EnumSiNo
    
    Dim m_FechaRef As Date
    Dim m_FechaMax As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eCaducado <> 0 Then
        Caducado = m_eCaducado
        Exit Property
    End If
    If Me.FechaCierre <> "" Then
        m_eCaducado = EnumSiNo.No
        Exit Property
    End If
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    m_FechaMax = Me.EdicionUltima.FechaMaxProximaPublicacion
    
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If IsDate(m_FechaMax) Then
        If m_FechaRef > CDate(m_FechaMax) Then
            m_eCaducado = EnumSiNo.Sí
        Else
            m_eCaducado = EnumSiNo.No
        End If
    End If
    
    Caducado = m_eCaducado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.Caducado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let Caducado(ByVal eNewValue As EnumSiNo)

    m_eCaducado = eNewValue

End Property

Public Property Get ElaboradoObj() As Usuario
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objElaboradoObj Is Nothing Then
        Set ElaboradoObj = m_objElaboradoObj
        Exit Property
    End If
    If Me.Elaborado = "" Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    Set m_objElaboradoObj = getUsuario(, , Me.Elaborado, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ElaboradoObj = m_objElaboradoObj
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.ElaboradoObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set ElaboradoObj(ByVal objNewValue As Variant)
    
    Set m_objElaboradoObj = objNewValue
    
End Property
Public Property Get RevisadoObj() As Usuario
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objRevisadoObj Is Nothing Then
        Set RevisadoObj = m_objRevisadoObj
        Exit Property
    End If
    If Me.Revisado = "" Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    Set m_objRevisadoObj = getUsuario(, , Me.Revisado, , m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set RevisadoObj = m_objRevisadoObj
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.RevisadoObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set RevisadoObj(ByVal objNewValue As Variant)
    
    Set m_objRevisadoObj = objNewValue
    
End Property
Public Property Get AprobadoObj() As Usuario
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objAprobadoObj Is Nothing Then
        Set AprobadoObj = m_objAprobadoObj
        Exit Property
    End If
    If Me.Aprobado = "" Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    Set m_objAprobadoObj = getUsuario(, , Me.Aprobado, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set AprobadoObj = m_objAprobadoObj
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.AprobadoObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property
Public Property Set AprobadoObj(ByVal objNewValue As Variant)
    
    Set m_objAprobadoObj = objNewValue
    
End Property
Public Property Get colEdiciones() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Set m_objColEdiciones = Constructor.getEdiciones(Me.IDProyecto, Me.OrdenarColeEdicionesAscendentemente, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colEdiciones = m_objColEdiciones
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.colEdiciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set colEdiciones(ByVal objNewValue As Variant)
    
    Set m_objColEdiciones = objNewValue
    
End Property

Public Property Get ColEdicionesPublicadas() As Scripting.Dictionary
    
    Dim m_objEdicion As Edicion
    Dim m_Id As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.colEdiciones Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    For Each m_Id In Me.colEdiciones
        Set m_objEdicion = Me.colEdiciones(m_Id)
        If m_objEdicion.FechaPublicacion <> "" Then
            If m_objColEdicionesPublicadas Is Nothing Then
                Set m_objColEdicionesPublicadas = New Scripting.Dictionary
                m_objColEdiciones.CompareMode = TextCompare
            End If
            If Not m_objColEdicionesPublicadas.Exists(m_objEdicion.IDEdicion) Then
                m_objColEdicionesPublicadas.Add m_objEdicion.IDEdicion, m_objEdicion
            End If
        End If
        Set m_objEdicion = Nothing
    Next
    Set ColEdicionesPublicadas = m_objColEdicionesPublicadas
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.colEdiciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColEdicionesPublicadas(ByVal objNewValue As Variant)

    Set m_objColEdicionesPublicadas = objNewValue

End Property

Public Property Get EdicionActiva() As Edicion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Set m_ObjEdicionActiva = Constructor.getEdicionActiva(Me.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EdicionActiva = m_ObjEdicionActiva
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.EdicionActiva ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set EdicionActiva(ByVal sNewValue As Variant)
    
    
    Set m_ObjEdicionActiva = sNewValue

End Property
Public Property Get EdicionUltima() As Edicion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjEdicionUltima Is Nothing Then
        Set EdicionUltima = m_ObjEdicionUltima
        Exit Property
    End If
    Set m_ObjEdicionUltima = Constructor.getEdicionUltima(Me.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EdicionUltima = m_ObjEdicionUltima
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.EdicionUltima ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set EdicionUltima(ByVal sNewValue As Variant)
    
    
    Set m_ObjEdicionUltima = sNewValue

End Property

Public Property Get EdicionUltimaPublicada() As Edicion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjEdicionUltimaPublicada Is Nothing Then
        Set EdicionUltimaPublicada = m_ObjEdicionUltimaPublicada
        Exit Property
    End If
    Set m_ObjEdicionUltimaPublicada = Constructor.getEdicionUltimaPublicada(Me.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EdicionUltimaPublicada = m_ObjEdicionUltimaPublicada
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.EdicionUltimaPublicada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set EdicionUltimaPublicada(ByVal sNewValue As Variant)
    
    
    Set m_ObjEdicionUltimaPublicada = sNewValue

End Property

Public Property Get EdicionPrimera() As Edicion
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjEdicionPrimera Is Nothing Then
        Set EdicionPrimera = m_ObjEdicionPrimera
        Exit Property
    End If
    If Me.IDProyecto = "" Then
        m_Error = "No se conoce el IDProyecto"
        Err.Raise 1000
    End If
    Set m_ObjEdicionPrimera = Constructor.getEdicionPrimera(Me.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EdicionPrimera = m_ObjEdicionPrimera
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.EdicionPrimera ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set EdicionPrimera(ByVal sNewValue As Variant)
   
    Set m_ObjEdicionPrimera = sNewValue

End Property


Public Property Get ColAutorizados() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not m_ObjColAutorizados Is Nothing Then
        Set ColAutorizados = m_ObjColAutorizados
        Exit Property
    End If
    Set m_ObjColAutorizados = Constructor.getExpedienteResponsables(Me.IDExpediente, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAutorizados = m_ObjColAutorizados
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.ColAutorizados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColAutorizados(ByVal objNewValue As Variant)
    
    
    Set m_ObjColAutorizados = objNewValue

End Property

Public Property Get ColAnexos() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColAnexos Is Nothing Then
        Set ColAnexos = m_ObjColAnexos
        Exit Property
    End If
    Set m_ObjColAnexos = Constructor.getAnexosByProyecto(Me.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAnexos = m_ObjColAnexos
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.ColAnexos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColAnexos(ByVal sNewValue As Variant)
    
    Set m_ObjColAnexos = sNewValue
     
End Property


Public Property Get UsuarioCalidad() As Usuario
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjUsuarioCalidad Is Nothing Then
        Set UsuarioCalidad = m_ObjUsuarioCalidad
        Exit Property
    End If
    If Me.NombreUsuarioCalidad = "" Then
        Exit Property
    End If
    Set m_ObjUsuarioCalidad = Constructor.getUsuario(, , Me.NombreUsuarioCalidad, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioCalidad = m_ObjUsuarioCalidad
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.UsuarioCalidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    
End Property


Public Property Get UsuarioElaboradoObj() As Usuario
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objUsuarioElaboradoObj Is Nothing Then
        Set UsuarioElaboradoObj = m_objUsuarioElaboradoObj
        Exit Property
    End If
    If Me.Elaborado = "" Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    Set m_objUsuarioElaboradoObj = Constructor.getUsuario(, , Me.Elaborado, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioElaboradoObj = m_objUsuarioElaboradoObj
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.UsuarioElaboradoObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    

End Property

Public Property Set UsuarioElaboradoObj(ByVal objNewValue As Variant)

    Set m_objUsuarioElaboradoObj = objNewValue

End Property

Public Property Get UsuarioRevisadoObj() As Usuario

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objUsuarioRevisadoObj Is Nothing Then
        Set UsuarioRevisadoObj = m_objUsuarioRevisadoObj
        Exit Property
    End If
    If Me.Revisado = "" Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    
    Set m_objUsuarioRevisadoObj = Constructor.getUsuario(, , Me.Revisado, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioRevisadoObj = m_objUsuarioRevisadoObj
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.UsuarioRevisadoObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set UsuarioRevisadoObj(ByVal objNewValue As Variant)

    Set m_objUsuarioRevisadoObj = objNewValue

End Property

Public Property Get UsuarioAprobadoObj() As Usuario
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objUsuarioAprobadoObj Is Nothing Then
        Set UsuarioAprobadoObj = m_objUsuarioAprobadoObj
        Exit Property
    End If
    If Me.Aprobado = "" Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    
    Set m_objUsuarioAprobadoObj = Constructor.getUsuario(, , Me.Aprobado, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioAprobadoObj = m_objUsuarioAprobadoObj
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.UsuarioAprobadoObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Set UsuarioAprobadoObj(ByVal objNewValue As Variant)

    Set m_objUsuarioAprobadoObj = objNewValue

End Property

Public Property Get OrdenarColeEdicionesAscendentemente() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_eOrdenarColeEdicionesAscendentemente <> 0 Then
        OrdenarColeEdicionesAscendentemente = m_eOrdenarColeEdicionesAscendentemente
        Exit Property
    End If
    m_eOrdenarColeEdicionesAscendentemente = EnumSiNo.Sí
    OrdenarColeEdicionesAscendentemente = m_eOrdenarColeEdicionesAscendentemente
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.OrdenarColeEdicionesAscendentemente ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Let OrdenarColeEdicionesAscendentemente(ByVal eNewValue As EnumSiNo)

    m_eOrdenarColeEdicionesAscendentemente = eNewValue

End Property


Public Property Get DiasParaElSiguienteAvisoAJP() As String
    
    Dim m_FechaSiguienteAvisoJP As String
    Dim intDiasParaSiguienteAvisoJP As Integer
    Dim m_DiasPreviosAviso As Integer
    Dim m_FechaAviso As String
    Dim m_FechaRef As Date
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sDiasParaElSiguienteAvisoAJP) Then
        DiasParaElSiguienteAvisoAJP = m_sDiasParaElSiguienteAvisoAJP
        Exit Property
    End If
    If Me.FechaCierre = "" Then
        Exit Property
    End If
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    m_DiasPreviosAviso = m_ObjEntorno.JPDiasPreviosParaElAviso
    m_Error = m_ObjEntorno.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not IsNumeric(m_DiasPreviosAviso) Then
        m_Error = "No se ha podido obtener los días previos para el aviso a los JP"
        Err.Raise 1000
    End If
    If Me.EdicionActiva Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    If Not IsDate(Me.EdicionActiva.FechaMaxProximaPublicacionCalculado) Then
        m_Error = Me.EdicionActiva.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se ha podido obtener FechaMaxProximaPublicacionCalculado"
        Err.Raise 1000
    End If
    m_FechaAviso = DateAdd(CDate(Me.EdicionActiva.FechaMaxProximaPublicacionCalculado), -m_DiasPreviosAviso, m_FechaRef)
    m_sDiasParaElSiguienteAvisoAJP = DateDiff("d", m_FechaRef, CDate(m_FechaAviso))
    DiasParaElSiguienteAvisoAJP = m_sDiasParaElSiguienteAvisoAJP
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.DiasParaElSiguienteAvisoAJP ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Function Abrir(p_Error As String) As String
    
    On Error GoTo errores
    
    Avance "Abrir Proyecto .... Datos iniciales"
    
    If Not IsDate(Me.FechaCierre) Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "El Proyecto no aparece como cerrado"
        Err.Raise 1000
    End If
    
    AbrirProyectoTransaccional Me, , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If FormularioAbierto("FormProyectosGestion") Then
        Form_FormProyectosGestion.ActualizarProyecto Me, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Form_FormProyectosGestion.ListaFiltrados_Click
    End If
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.Abrir ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Function

Public Property Get UsuarioConectadoAutorizado() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eUsuarioConectadoAutorizado <> 0 Then
        UsuarioConectadoAutorizado = m_eUsuarioConectadoAutorizado
        Exit Property
    End If
    m_eUsuarioConectadoAutorizado = UsuarioAutorizado(, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    UsuarioConectadoAutorizado = m_eUsuarioConectadoAutorizado
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.UsuarioConectadoAutorizado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get URLUltimaPublicacion() As String
    
    Dim m_objEdicion As Edicion
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLUltimaPublicacion <> "" Then
        URLUltimaPublicacion = m_sURLUltimaPublicacion
        Exit Property
    End If
    If Me.EdicionUltima Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se ha podido obtener la Última Edición"
        Err.Raise 1000
    End If
    Set m_objEdicion = Me.EdicionUltima
    If m_objEdicion.EsActivo = EnumSiNo.Sí Then
        If m_objEdicion.EsPrimeraEdicion Then
            m_Error = m_objEdicion.Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Property
        End If
        Set m_objEdicion = Me.EdicionUltima.EdicionAnterior
        m_Error = Me.EdicionUltima.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_objEdicion.EsActivo = EnumSiNo.Sí Then
            m_Error = "No se ha podido obtener la última Edición NO activa"
            Err.Raise 1000
        End If
    End If
    m_sURLUltimaPublicacion = m_objEdicion.URLArchivoInforme
    m_Error = m_objEdicion.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    URLUltimaPublicacion = m_sURLUltimaPublicacion
    Exit Function
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.URLUltimaPublicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EdicionUltimaNoActiva() As Edicion
    
    Dim m_objEdicion As Edicion
    Dim m_NumeroEdiciones As Integer
    Dim i As Integer
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjEdicionUltimaNoActiva Is Nothing Then
        Set EdicionUltimaNoActiva = m_ObjEdicionUltimaNoActiva
        Exit Property
    End If
    Set m_objEdicion = Me.EdicionUltima
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_objEdicion Is Nothing Then
        m_Error = "No se ha podido obtener la Edición Última"
        Err.Raise 1000
    End If
    If m_objEdicion.FechaPublicacion <> "" Then
        Set m_ObjEdicionUltimaNoActiva = m_objEdicion
        Set EdicionUltimaNoActiva = m_ObjEdicionUltimaNoActiva
        Exit Property
    End If
    m_NumeroEdiciones = m_objEdicion.Proyecto.colEdiciones.Count
    m_Error = m_objEdicion.Proyecto.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    For i = 1 To m_NumeroEdiciones - 1
        If m_objEdicion.EsPrimeraEdicion = EnumSiNo.No Then
            Set m_objEdicion = m_objEdicion.EdicionAnterior
            m_Error = Me.Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            If m_objEdicion.FechaPublicacion <> "" Then
                Set m_ObjEdicionUltimaNoActiva = m_objEdicion
                Set EdicionUltimaNoActiva = m_ObjEdicionUltimaNoActiva
                Exit Property
            End If
        End If
    Next
    Set EdicionUltimaNoActiva = m_ObjEdicionUltimaNoActiva
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.EdicionUltimaNoActiva ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set EdicionUltimaNoActiva(ByVal objNewValue As Variant)

    Set m_ObjEdicionUltimaNoActiva = objNewValue

End Property

Public Property Get EsActivo() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eEsActivo <> 0 Then
        EsActivo = m_eEsActivo
        Exit Property
    End If
    If Me.FechaCierre = "" Then
        m_eEsActivo = EnumSiNo.Sí
    Else
        m_eEsActivo = EnumSiNo.No
    End If
    EsActivo = m_eEsActivo
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.EsActivo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColRiesgosExternos() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRiesgosExternos Is Nothing Then
        Set ColRiesgosExternos = m_objColRiesgosExternos
        Exit Property
    End If
    Set m_objColRiesgosExternos = Constructor.getRiesgosExternos(Me.IDProyecto, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColRiesgosExternos = m_objColRiesgosExternos
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.ColRiesgosExternos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRiesgosExternos(ByVal objNewValue As Variant)
    
    Set m_objColRiesgosExternos = objNewValue
    If objNewValue Is Nothing Then
        Set Me.ColRiesgosOferta = Nothing
        Set Me.ColRiesgosPedidos = Nothing
        Set Me.ColRiesgosSubContratista = Nothing
    End If
End Property

Public Property Get ColRiesgosOferta() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColRiesgosOferta Is Nothing Then
        Set ColRiesgosOferta = m_ObjColRiesgosOferta
        Exit Property
    End If
    Set m_ObjColRiesgosOferta = Constructor.getRiesgosExternos(Me.IDProyecto, EnumOrigenRiesgoExterno.Oferta, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColRiesgosOferta = m_ObjColRiesgosOferta
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.ColRiesgosOferta ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRiesgosOferta(ByVal objNewValue As Variant)
    
    Set m_ObjColRiesgosOferta = objNewValue
End Property

Public Property Get ColRiesgosSubContratista() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRiesgosSubContratista Is Nothing Then
        Set ColRiesgosSubContratista = m_objColRiesgosSubContratista
        Exit Property
    End If
    Set m_objColRiesgosSubContratista = Constructor.getRiesgosExternos(Me.IDProyecto, EnumOrigenRiesgoExterno.Subcontratista, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColRiesgosSubContratista = m_objColRiesgosSubContratista
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.ColRiesgosSubContratista ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRiesgosSubContratista(ByVal objNewValue As Variant)

    Set m_objColRiesgosSubContratista = objNewValue
    
End Property

Public Property Get ColRiesgosPedidos() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRiesgosPedidos Is Nothing Then
        Set ColRiesgosPedidos = m_objColRiesgosPedidos
        Exit Property
    End If
    Set m_objColRiesgosPedidos = Constructor.getRiesgosExternos(Me.IDProyecto, EnumOrigenRiesgoExterno.Pedido, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColRiesgosPedidos = m_objColRiesgosPedidos
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.ColRiesgosPedidos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRiesgosPedidos(ByVal objNewValue As Variant)

    Set m_objColRiesgosPedidos = objNewValue
    
End Property

Public Property Get TieneEdiciones() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.colEdiciones Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneEdiciones = EnumSiNo.No
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneEdiciones = EnumSiNo.Sí
    End If
    TieneEdiciones = m_eTieneEdiciones
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Riesgo.TieneEdiciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get TieneRiesgosExt() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.ColRiesgosExternos Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneRiesgosExt = EnumSiNo.No
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneRiesgosExt = EnumSiNo.Sí
    End If
    TieneRiesgosExt = m_eTieneRiesgosExt
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.TieneRiesgosExt ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get TieneRiesgosDeOferta() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.ColRiesgosOferta Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneRiesgosDeOferta = EnumSiNo.No
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneRiesgosDeOferta = EnumSiNo.Sí
    End If
    TieneRiesgosDeOferta = m_eTieneRiesgosDeOferta
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.TieneRiesgosDeOferta ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get TieneRiesgosDeSubContratista() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.ColRiesgosSubContratista Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneRiesgosDeSubContratista = EnumSiNo.No
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneRiesgosDeSubContratista = EnumSiNo.Sí
    End If
    TieneRiesgosDeSubContratista = m_eTieneRiesgosDeSubContratista
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.TieneRiesgosDeSubContratista ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get TieneRiesgosDePedidos() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.ColRiesgosPedidos Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneRiesgosDePedidos = EnumSiNo.No
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneRiesgosDePedidos = EnumSiNo.Sí
    End If
    TieneRiesgosDePedidos = m_eTieneRiesgosDePedidos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.TieneRiesgosDePedidos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get ColAnexosTotales() As Scripting.Dictionary
    
    Dim m_ObjAnexo As Anexo
    Dim m_IDAnexo As Variant
    Dim m_IDEdicion As Variant
    Dim m_objEdicion As Edicion
    Dim m_IdRiesgo As Variant
    Dim m_ObjRiesgo As riesgo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Set m_objColAnexosTotales = Nothing
    Set Me.ColAnexos = Nothing
    If Not Me.ColAnexos Is Nothing Then
        For Each m_IDAnexo In Me.ColAnexos.keys
            If Nz(m_IDAnexo, "") <> "" Then
                Set m_ObjAnexo = Me.ColAnexos(m_IDAnexo)
                If m_objColAnexosTotales Is Nothing Then
                    Set m_objColAnexosTotales = New Scripting.Dictionary
                    m_objColAnexosTotales.CompareMode = TextCompare
                End If
                Set m_objColAnexosTotales(m_IDAnexo) = m_ObjAnexo
                Set m_ObjAnexo = Nothing
            End If

        Next
    End If
    If Me.TieneEdiciones = EnumSiNo.Sí Then
        For Each m_IDEdicion In Me.colEdiciones.keys
            If Nz(m_IDEdicion, "") <> "" Then
                Set m_objEdicion = Me.colEdiciones(m_IDEdicion)
                If Not m_objEdicion.ColAnexos Is Nothing Then
                    For Each m_IDAnexo In m_objEdicion.ColAnexos.keys
                        If Nz(m_IDAnexo, "") <> "" Then
                            Set m_ObjAnexo = m_objEdicion.ColAnexos(m_IDAnexo)
                            If m_objColAnexosTotales Is Nothing Then
                                Set m_objColAnexosTotales = New Scripting.Dictionary
                                m_objColAnexosTotales.CompareMode = TextCompare
                            End If
                            Set m_objColAnexosTotales(m_IDAnexo) = m_ObjAnexo
                            Set m_ObjAnexo = Nothing
                        End If

                    Next
                End If
                If m_objEdicion.TieneRiesgos = EnumSiNo.Sí Then
                    For Each m_IdRiesgo In m_objEdicion.colRiesgos.keys
                        If Nz(m_IdRiesgo, "") <> "" Then
                            Set m_ObjRiesgo = m_objEdicion.colRiesgos(m_IdRiesgo)
                            If Not m_ObjRiesgo.ColAnexos Is Nothing Then
                                For Each m_IDAnexo In m_ObjRiesgo.ColAnexos.keys
                                    If Nz(m_IDAnexo, "") <> "" Then
                                        Set m_ObjAnexo = m_ObjRiesgo.ColAnexos(m_IDAnexo)
                                        If m_objColAnexosTotales Is Nothing Then
                                            Set m_objColAnexosTotales = New Scripting.Dictionary
                                            m_objColAnexosTotales.CompareMode = TextCompare
                                        End If
                                        Set m_objColAnexosTotales(m_IDAnexo) = m_ObjAnexo
                                        Set m_ObjAnexo = Nothing
                                    End If

                                Next
                            End If
                            Set m_ObjRiesgo = Nothing
                        End If

                    Next
                End If

                Set m_objEdicion = Nothing
            End If
        Next
    End If
    Set ColAnexosTotales = m_objColAnexosTotales
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.ColAnexosTotales ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColAnexosTotales(ByVal objNewValue As Variant)

    Set m_objColAnexosTotales = objNewValue

End Property

Public Property Get ColRiesgosDeJP() As Scripting.Dictionary
    
    Dim m_ObjUltimaEdicion As Edicion
    Dim m_ObjRiesgo As riesgo
    Dim m_IdRiesgo As Variant
    Dim m_objColRiesgos As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Set m_ObjUltimaEdicion = Me.EdicionUltima
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjUltimaEdicion Is Nothing Then
        m_Error = "No se ha podido obtener la última edición"
        Err.Raise 1000
    End If
    Set m_objColRiesgos = m_ObjUltimaEdicion.colRiesgos
    m_Error = m_ObjUltimaEdicion.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_objColRiesgos Is Nothing Then
        Exit Property
    End If
    For Each m_IdRiesgo In m_objColRiesgos.keys
        Set m_ObjRiesgo = m_objColRiesgos(m_IdRiesgo)
        If m_ObjRiesgo.RiesgoExterno Is Nothing Then
            m_Error = m_ObjRiesgo.Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            If m_objColRiesgosDeJP Is Nothing Then
                Set m_objColRiesgosDeJP = New Scripting.Dictionary
                m_objColRiesgosDeJP.CompareMode = TextCompare
            End If
            If Not m_objColRiesgosDeJP.Exists(m_IdRiesgo) Then
                m_objColRiesgosDeJP.Add m_IdRiesgo, m_ObjRiesgo
            End If
        End If
        Set m_ObjRiesgo = Nothing
    Next
    Set ColRiesgosDeJP = m_objColRiesgosDeJP
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.ColRiesgosDeJP ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRiesgosDeJP(ByVal objNewValue As Variant)

    Set m_objColRiesgosDeJP = objNewValue

End Property

Public Property Get AlmenosUnaEdicionPublicada() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eAlmenosUnaEdicionPublicada <> Empty Then
        AlmenosUnaEdicionPublicada = m_eAlmenosUnaEdicionPublicada
        Exit Property
    End If
    If Me.EdicionPrimera Is Nothing Then
        m_eAlmenosUnaEdicionPublicada = EnumSiNo.No
        AlmenosUnaEdicionPublicada = m_eAlmenosUnaEdicionPublicada
        Exit Property
    End If
    If Me.EdicionPrimera.FechaPublicacion = "" Then
        m_eAlmenosUnaEdicionPublicada = EnumSiNo.No
    Else
        m_eAlmenosUnaEdicionPublicada = EnumSiNo.Sí
    End If
    AlmenosUnaEdicionPublicada = m_eAlmenosUnaEdicionPublicada
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.AlmenosUnaEdicionPublicada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Let AlmenosUnaEdicionPublicada(p_Valor As EnumSiNo)
    
    m_eAlmenosUnaEdicionPublicada = p_Valor

End Property



Public Property Get ColSuministradores() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColSuministradores Is Nothing Then
        Set ColSuministradores = m_objColSuministradores
        Exit Property
    End If
    Set m_objColSuministradores = Constructor.getTodosProyectosSuministradoresPorCriterio(, Me.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColSuministradores = m_objColSuministradores
    
     Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.ColSuministradores ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColSuministradores(ByVal objNewValue As Variant)

    Set m_objColSuministradores = objNewValue

End Property

Public Property Get colCambios() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColCambios Is Nothing Then
        Set colCambios = m_ObjColCambios
        Exit Property
    End If
    Set m_ObjColCambios = Constructor.getCambiosPorProyecto(Me.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colCambios = m_ObjColCambios
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.colCambios ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set colCambios(ByVal objNewValue As Variant)
    
    Set m_ObjColCambios = objNewValue
    
End Property

Public Property Get HTMLProyecto() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sHTMLProyecto <> "" Then
        HTMLProyecto = m_sHTMLProyecto
        Exit Property
    End If
    m_sHTMLProyecto = getHTMLAltaProyecto(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    HTMLProyecto = m_sHTMLProyecto
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.HTMLProyecto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get HTMLTablaProyectoDatos() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sHTMLTablaProyectoDatos <> "" Then
        HTMLTablaProyectoDatos = m_sHTMLTablaProyectoDatos
        Exit Property
    End If
    m_sHTMLTablaProyectoDatos = getHTMLTablaProyectoDatos(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    HTMLTablaProyectoDatos = m_sHTMLTablaProyectoDatos
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.HTMLTablaProyectoDatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get HTMLTablaEdiciones() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sHTMLTablaEdiciones <> "" Then
        HTMLTablaEdiciones = m_sHTMLTablaEdiciones
        Exit Property
    End If
    m_sHTMLTablaEdiciones = getHTMLTablaEdiciones(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    HTMLTablaEdiciones = m_sHTMLTablaEdiciones
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.HTMLTablaEdiciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get HTMLTablaProyectoAutorizados() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sHTMLTablaProyectoAutorizados <> "" Then
        HTMLTablaProyectoAutorizados = m_sHTMLTablaProyectoAutorizados
        Exit Property
    End If
    m_sHTMLTablaProyectoAutorizados = getHTMLTablaProyectoAutorizados(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    HTMLTablaProyectoAutorizados = m_sHTMLTablaProyectoAutorizados
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.HTMLTablaProyectoAutorizados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get CadenaCorreoAutorizados() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sCadenaCorreoAutorizados <> "" Then
        CadenaCorreoAutorizados = m_sCadenaCorreoAutorizados
        Exit Property
    End If
    m_sCadenaCorreoAutorizados = getCadenaCorreoAutorizados(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    CadenaCorreoAutorizados = m_sCadenaCorreoAutorizados
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.CadenaCorreoAutorizados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get CadenaNombreAutorizadosCalculados() As String
    
    Dim m_Id As Variant
    Dim m_Responsable As ExpedienteResponsable
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sCadenaNombreAutorizadosCalculados <> "" Then
        CadenaNombreAutorizadosCalculados = m_sCadenaNombreAutorizadosCalculados
        Exit Property
    End If
    If Me.ColAutorizados Is Nothing Then
        Exit Property
    End If
    For Each m_Id In Me.ColAutorizados
        Set m_Responsable = Me.ColAutorizados(m_Id)
        If m_sCadenaNombreAutorizadosCalculados = "" Then
            m_sCadenaNombreAutorizadosCalculados = m_Responsable.Usuario.Nombre
        Else
            m_sCadenaNombreAutorizadosCalculados = m_sCadenaNombreAutorizadosCalculados & ";" & m_Responsable.Usuario.Nombre
        End If
        Set m_Responsable = Nothing
    Next
    CadenaNombreAutorizadosCalculados = m_sCadenaNombreAutorizadosCalculados
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.CadenaNombreAutorizadosCalculados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Private Function getHTMLTablaProyectoDatos( _
                                            Optional ByRef p_Error As String _
                                            ) As String
    Dim m_mensaje As String
    
    
    Dim m_NombreRevisado As String
    Dim m_NombreAprobado As String
    Dim m_NombreElaborado As String
    
    
    On Error GoTo errores
    
    
    With Me
        m_NombreRevisado = .Revisado
        m_NombreAprobado = .Aprobado
        m_NombreElaborado = .Elaborado
        
    
    End With
    m_mensaje = m_mensaje & "<table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""ColespanArriba"" colspan='2'>gestión de riesgos DE " & UCase(Me.Proyecto) & "</td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Jurídica</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & Me.Juridica & " </td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Proyecto</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & Me.Proyecto & " </td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Nemotécnico</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & Me.NombreProyecto & " </td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Cliente</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & Me.Cliente & " </td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        
        m_mensaje = m_mensaje & "<tr>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"">F.Prev Cierre</td>" & vbNewLine
            Me.FechaPrevistaCierre = Me.FechaPrevistaCierreCalculada
            If IsDate(Me.FechaPrevistaCierre) Then
                m_mensaje = m_mensaje & "<td> " & Me.FechaPrevistaCierre & " </td>" & vbNewLine
            Else
                m_mensaje = m_mensaje & "<td> &nbsp;  </td>" & vbNewLine
            End If
            
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>"
        
        m_mensaje = m_mensaje & "<tr>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Código Documentación</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & Me.CodigoDocumento & " </td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>"
        
        m_mensaje = m_mensaje & "<tr>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Proyecto En UTE</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & Me.EnUTE & " </td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>"
        m_mensaje = m_mensaje & "<tr>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"">F.Cierre</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & Me.FechaCierre & " </td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>"
        
        
        m_mensaje = m_mensaje & "<tr>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Elaborado</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & IIf(m_NombreElaborado <> "", m_NombreElaborado, "&nbsp;") & " </td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Revisado</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & IIf(m_NombreRevisado <> "", m_NombreRevisado, "&nbsp;") & " </td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Aprobado</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & IIf(m_NombreAprobado <> "", m_NombreAprobado, "&nbsp;") & " </td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Seguimiento Calidad</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & Me.NombreUsuarioCalidad & " </td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
    m_mensaje = m_mensaje & "</table>" & vbNewLine
    
    getHTMLTablaProyectoDatos = m_mensaje
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo getHTMLTablaProyecto ha devuelto el error: " & Err.Description
    End If
End Function

Private Function getHTMLTablaProyectoRiesgosOferta( _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    Dim m_mensaje As String
    Dim m_objColRiesgosExternos As Scripting.Dictionary
    Dim m_IDRiesgoExt As Variant
    Dim m_ObjRiesgoExterno As RiesgoExterno
    Dim m_CodRiesgo As String
    Dim m_Descripcion As String
    Dim m_FechaDetectado As String
    Dim m_Trasladar As String
    Dim m_MotivoNoIntegrado As String
    Dim m_ObjEdicionPrimera As Edicion
    On Error GoTo errores
    If Me Is Nothing Then
        p_Error = "No se ha incluido datos del proyecto"
        Err.Raise 1000
    End If
    Set m_ObjEdicionPrimera = Me.EdicionPrimera
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_objColRiesgosExternos = Me.ColRiesgosOferta
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_mensaje = m_mensaje & "<table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""ColespanArriba"" colspan='6'>Riesgos de la Oferta </td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Cód</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Edición Detectado</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Descripción</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Fecha Detectado</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Trasladado</td>" & vbNewLine
            
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Motivo No Traslado</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        If m_objColRiesgosExternos Is Nothing Then
            m_mensaje = m_mensaje & "</table>" & vbNewLine
            getHTMLTablaProyectoRiesgosOferta = m_mensaje
            Exit Function
        End If
        If m_objColRiesgosExternos.Count = 0 Then
            m_mensaje = m_mensaje & "</table>" & vbNewLine
            getHTMLTablaProyectoRiesgosOferta = m_mensaje
            Exit Function
        End If
        For Each m_IDRiesgoExt In m_objColRiesgosExternos.keys
            If Nz(m_IDRiesgoExt, "") = "" Then
                GoTo SiguienteRiesgoExt
            End If
            Set m_ObjRiesgoExterno = m_objColRiesgosExternos(m_IDRiesgoExt)
            If Not m_ObjRiesgoExterno Is Nothing Then
                With m_ObjRiesgoExterno
                    m_CodRiesgo = .CodRiesgo
                    m_Descripcion = .Descripcion
                    m_FechaDetectado = .FechaDetectado
                    m_Trasladar = .Trasladar
                    m_MotivoNoIntegrado = .MotivoNoIntegrado
                End With
                
                m_mensaje = m_mensaje & "<tr>" & vbNewLine
                    m_mensaje = m_mensaje & "<td> " & IIf(m_CodRiesgo <> "", m_CodRiesgo, "&nbsp;") & " </td>" & vbNewLine
                    m_mensaje = m_mensaje & "<td> " & m_ObjRiesgoExterno.Edicion.Edicion & " </td>" & vbNewLine
                    m_mensaje = m_mensaje & "<td> " & IIf(m_Descripcion <> "", m_Descripcion, "&nbsp;") & " </td>" & vbNewLine
                    m_mensaje = m_mensaje & "<td> " & IIf(m_FechaDetectado <> "", m_FechaDetectado, "&nbsp;") & " </td>" & vbNewLine
                    m_mensaje = m_mensaje & "<td> " & m_Trasladar & " </td>" & vbNewLine
                    If m_Trasladar = "No" Then
                        m_mensaje = m_mensaje & "<td> " & IIf(m_MotivoNoIntegrado <> "", m_MotivoNoIntegrado, "&nbsp;") & " </td>" & vbNewLine
                    Else
                        m_mensaje = m_mensaje & "<td> &nbsp </td>" & vbNewLine
                    End If
                m_mensaje = m_mensaje & "</tr>" & vbNewLine
            End If
            Set m_ObjRiesgoExterno = Nothing
SiguienteRiesgoExt:
        Next
    m_mensaje = m_mensaje & "</table>" & vbNewLine
    
    getHTMLTablaProyectoRiesgosOferta = m_mensaje
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo getHTMLTablaProyectoRiesgosOferta ha devuelto el error: " & Err.Description
    End If
End Function
Private Function getHTMLAltaProyecto( _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                                
    Dim m_HTMLCabecera As String
    Dim m_HTMLUsuariosConectados As String
    Dim m_HTMLTablaProyectoDatos As String
    Dim m_HTMLTablaProyectoAutorizados As String
    Dim m_HTMLTablaProyectoRiesgosOferta As String
    Dim m_HTMLTablaSuministradoresConGestionRiesgos As String
    
    
    Dim m_mensaje As String
    
    On Error GoTo errores
    
    m_HTMLCabecera = getHTMCabecera("Alta de gestión de riesgos", p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLUsuariosConectados = getHTMLUsuariosConectados(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTablaProyectoDatos = getHTMLTablaProyectoDatos(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTablaProyectoAutorizados = Me.HTMLTablaEdiciones
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTablaProyectoRiesgosOferta = getHTMLTablaProyectoRiesgosOferta(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTablaSuministradoresConGestionRiesgos = getHTMLTablaSuministradoresConGestionRiesgos(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_mensaje = m_HTMLCabecera & vbNewLine
    m_mensaje = m_mensaje & m_HTMLUsuariosConectados & vbNewLine
    m_mensaje = m_mensaje & "<br /><br />" & vbNewLine
    m_mensaje = m_mensaje & "<STRONG> ALTA DE gestión de riesgos</STRONG>" & vbNewLine
    m_mensaje = m_mensaje & "<br /><br />" & vbNewLine
    m_mensaje = m_mensaje & m_HTMLTablaProyectoDatos & vbNewLine
    m_mensaje = m_mensaje & "<br /><br />" & vbNewLine
    m_mensaje = m_mensaje & m_HTMLTablaProyectoAutorizados & vbNewLine
    m_mensaje = m_mensaje & "<br /><br />" & vbNewLine
    m_mensaje = m_mensaje & m_HTMLTablaProyectoRiesgosOferta & vbNewLine
    m_mensaje = m_mensaje & "<br /><br />" & vbNewLine
    m_mensaje = m_mensaje & m_HTMLTablaSuministradoresConGestionRiesgos & vbNewLine
    
    If Me.TieneRiesgosDeOferta = EnumSiNo.Sí Then
        m_mensaje = m_mensaje & "</BR>" & vbNewLine
        m_mensaje = m_mensaje & "<p><strong> Recuerde de terminar de rellenar los datos de los riesgo trasladados de la Oferta </strong></p>" & vbNewLine
    End If
    m_mensaje = m_mensaje & "<p><strong>¡¡¡NO RESPONDA A ESTE CORREO (es un mensaje automático)!!!</strong></p>" & vbNewLine
    m_mensaje = m_mensaje & m_ObjEntorno.FinHTML & vbNewLine
    
    getHTMLAltaProyecto = m_mensaje
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo getHTMLAltaProyecto ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getHTMLTablaProyectoAutorizados( _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    Dim m_mensaje As String
    Dim m_ObjColAutorizados As Scripting.Dictionary
    Dim m_UsuarioRed As Variant
    Dim m_ObjProyectoResponsable As ExpedienteResponsable
    Dim m_ObjUsuario As Usuario
    Dim m_NombreResponsable As String
    Dim m_EsJefeProyecto As String
    Dim m_EnvioCorreos As String
    
    On Error GoTo errores
    If Me Is Nothing Then
        p_Error = "No se ha incluido datos del proyecto"
        Err.Raise 1000
    End If
    Set m_ObjColAutorizados = Me.ColAutorizados
    If m_ObjColAutorizados Is Nothing Then
        Exit Function
    End If
    If m_ObjColAutorizados.Count = 0 Then
        Exit Function
    End If
    m_mensaje = m_mensaje & "<table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""ColespanArriba"" colspan='3'>Usuarios Autorizados </td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Autorizado</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">JP</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Correo</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        For Each m_UsuarioRed In m_ObjColAutorizados.keys
            If Nz(m_UsuarioRed, "") <> "" Then
            
                Set m_ObjProyectoResponsable = m_ObjColAutorizados(m_UsuarioRed)
                If Not m_ObjProyectoResponsable Is Nothing Then
                    Set m_ObjUsuario = m_ObjProyectoResponsable.Usuario
                    If Not m_ObjUsuario Is Nothing Then
                        m_NombreResponsable = m_ObjUsuario.Nombre
                        m_EsJefeProyecto = m_ObjProyectoResponsable.EsJefeProyecto
                        m_EnvioCorreos = m_ObjProyectoResponsable.CorreoSiempre
                        m_mensaje = m_mensaje & "<tr>" & vbNewLine
                            m_mensaje = m_mensaje & "<td> " & IIf(m_NombreResponsable <> "", m_NombreResponsable, "&nbsp;") & " </td>" & vbNewLine
                            m_mensaje = m_mensaje & "<td> " & IIf(m_EsJefeProyecto <> "", m_EsJefeProyecto, "&nbsp;") & " </td>" & vbNewLine
                            m_mensaje = m_mensaje & "<td> " & IIf(m_EnvioCorreos <> "", m_EnvioCorreos, "&nbsp;") & " </td>" & vbNewLine
                        m_mensaje = m_mensaje & "</tr>" & vbNewLine
                    End If
                End If
            End If
        Next
    m_mensaje = m_mensaje & "</table>" & vbNewLine
    getHTMLTablaProyectoAutorizados = m_mensaje
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo getHTMLTablaProyectoAutorizados ha devuelto el error: " & Err.Description
    End If
End Function

Private Function getHTMLTablaEdiciones( _
                                            Optional ByRef p_Error As String _
                                            ) As String
    Dim m_mensaje As String
    Dim m_IDEdicion As Variant
    Dim m_objEdicion As Edicion
    Dim m_TextoFechaProximaPublicacion
    Dim m_ObjUsuarioElabora As Usuario
    Dim m_ObjUsuarioAprueba As Usuario
    Dim m_ObjUsuarioRevisa As Usuario
    Dim m_NombreElabora As String
    Dim m_NombreAprueba As String
    Dim m_NombreRevisa As String
    
    
    On Error GoTo errores
    If Me Is Nothing Then
        p_Error = "No se ha incluido datos del proyecto"
        Err.Raise 1000
    End If
    If Me.TieneEdiciones = EnumSiNo.No Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "El Proyecto no tiene ninguna Edición Registrada"
        Err.Raise 1000
    Else
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_mensaje = m_mensaje & "<table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""ColespanArriba"" colspan='7'>EDICIONES DE " & UCase(Me.NombreProyecto) & "</td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"">ED.</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">F.Propuesta Publicación</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">F.Publicación</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Elaborado</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Revisado</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Aprobado</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Próxima Ed.</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
    For Each m_IDEdicion In Me.colEdiciones.keys
        If Nz(m_IDEdicion, "") <> "" Then
            Set m_objEdicion = Me.colEdiciones(m_IDEdicion)
            With m_objEdicion
                If IsDate(.FechaPublicacion) Then
                    m_TextoFechaProximaPublicacion = "-------"
                Else
                    m_TextoFechaProximaPublicacion = .FechaMaxProximaPublicacionCalculado
                End If
                Set m_ObjUsuarioElabora = .UsuarioElabora
                p_Error = .Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If Not m_ObjUsuarioElabora Is Nothing Then
                    m_NombreElabora = m_ObjUsuarioElabora.Nombre
                Else
                    m_NombreElabora = ""
                End If
                Set m_ObjUsuarioRevisa = .UsuarioRevisa
                p_Error = .Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If Not m_ObjUsuarioRevisa Is Nothing Then
                    m_NombreRevisa = m_ObjUsuarioRevisa.Nombre
                Else
                    m_NombreRevisa = ""
                End If
                Set m_ObjUsuarioAprueba = .UsuarioAprueba
                p_Error = .Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If Not m_ObjUsuarioAprueba Is Nothing Then
                    m_NombreAprueba = m_ObjUsuarioAprueba.Nombre
                Else
                    m_NombreAprueba = ""
                End If
                m_mensaje = m_mensaje & "<tr>" & vbNewLine
                    m_mensaje = m_mensaje & "<td>" & .Edicion & "</td>" & vbNewLine
                    If IsDate(.FechaPreparadaParaPublicar) Then
                        m_mensaje = m_mensaje & "<td>" & .FechaPreparadaParaPublicar & "</td>" & vbNewLine
                    Else
                        m_mensaje = m_mensaje & "<td> --------</td>" & vbNewLine
                    End If
                    m_mensaje = m_mensaje & "<td>" & .FechaPublicacion & "</td>" & vbNewLine
                    m_mensaje = m_mensaje & "<td>" & m_NombreElabora & "</td>" & vbNewLine
                    m_mensaje = m_mensaje & "<td>" & m_NombreRevisa & "</td>" & vbNewLine
                    m_mensaje = m_mensaje & "<td>" & m_NombreAprueba & "</td>" & vbNewLine
                    m_mensaje = m_mensaje & "<td>" & m_TextoFechaProximaPublicacion & "</td>" & vbNewLine
                    
                m_mensaje = m_mensaje & "</tr>" & vbNewLine
            End With
            Set m_objEdicion = Nothing
        End If
    Next
    m_mensaje = m_mensaje & "</table>" & vbNewLine
    
    getHTMLTablaEdiciones = m_mensaje
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo getHTMLTablaEdiciones ha devuelto el error: " & Err.Description
    End If
End Function

Private Function getCadenaCorreoAutorizados( _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim m_Col As Scripting.Dictionary
    Dim m_Id As Variant
    Dim m_Responsable As ExpedienteResponsable
    
    
    
    
    On Error GoTo errores
    If Me Is Nothing Then
        p_Error = "No se ha incluido datos del proyecto"
        Err.Raise 1000
    End If
    Set m_Col = Me.ColAutorizados
    If m_Col Is Nothing Then
        Exit Function
    End If
    If m_Col.Count = 0 Then
        Exit Function
    End If
   
    For Each m_Id In m_Col
        Set m_Responsable = m_Col(m_Id)
        If m_Responsable.CorreoSiempre = "Sí" Then
            If getCadenaCorreoAutorizados = "" Then
                getCadenaCorreoAutorizados = m_Responsable.Usuario.CorreoUsuario
            Else
                getCadenaCorreoAutorizados = getCadenaCorreoAutorizados & ";" & m_Responsable.Usuario.CorreoUsuario
            End If
           
        End If
        Set m_Responsable = Nothing
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo getCadenaCorreoAutorizados ha devuelto el error: " & Err.Description
    End If
End Function


Public Function TieneAlMenosUnJP( _
                                    p_ObjColAutorizados As Scripting.Dictionary, _
                                    Optional ByRef p_Error As String _
                                    ) As EnumSiNo
    
    Dim m_ObjAutorizado As ExpedienteResponsable
    Dim m_UsuarioRed As Variant
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    For Each m_UsuarioRed In p_ObjColAutorizados.keys
        If Nz(m_UsuarioRed, "") = "" Then
            GoTo SiguienteAutorizado
        End If
        Set m_ObjAutorizado = p_ObjColAutorizados(m_UsuarioRed)
        If m_ObjAutorizado.EsJefeProyecto = "Sí" Then
            TieneAlMenosUnJP = EnumSiNo.Sí
            Set m_ObjAutorizado = Nothing
            Exit Function
        End If
        Set m_ObjAutorizado = Nothing
SiguienteAutorizado:
    Next
    TieneAlMenosUnJP = EnumSiNo.No
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo Proyecto.TieneAlMenosUnJP ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Function

Public Function TieneAlMenosUnEnvioCorreos( _
                                            p_ObjColAutorizados As Scripting.Dictionary, _
                                            Optional ByRef p_Error As String _
                                            ) As EnumSiNo
    
    Dim m_ObjAutorizado As ExpedienteResponsable
    Dim m_UsuarioRed As Variant
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    For Each m_UsuarioRed In p_ObjColAutorizados.keys
        If Nz(m_UsuarioRed, "") = "" Then
            GoTo SiguienteAutorizado
        End If
        Set m_ObjAutorizado = p_ObjColAutorizados(m_UsuarioRed)
        If m_ObjAutorizado.CorreoSiempre = "Sí" Then
            TieneAlMenosUnEnvioCorreos = EnumSiNo.Sí
            Set m_ObjAutorizado = Nothing
            Exit Function
        End If
        Set m_ObjAutorizado = Nothing
SiguienteAutorizado:
    Next
    TieneAlMenosUnEnvioCorreos = EnumSiNo.No
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo Proyecto.TieneAlMenosUnEnvioCorreos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Function

Public Function ObtenerGraficasRiesgos(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    ObtenerGraficasRiesgos = GenerarGraficosEvolucionRiesgos(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo Proyecto.ObtenerGraficasRiesgos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Private Function GenerarGraficosEvolucionRiesgos( _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    
    'getColEvolucionRiesgos("Priorización")=m_ColEvolucionPriorizacion
    'getColEvolucionRiesgos("Valoración")=m_ColEvolucionValoracion
    'getColEvolucionRiesgos("Estado")=m_ColEvolucionEstado
    
    'm_ColEvolucionPriorizacion("Edicion")=m_ColRiesgosValoresPriorizacion
    'm_ColRiesgosValoresPriorizacion("R001")=1
    
    'm_ColEvolucionValoracion("Edicion")=m_ColRiesgosValoresValoracion
    'm_ColRiesgosValoresValoracion("R001")="Alto"
    
    'm_ColEvolucionEstado("Edicion")=m_ColRiesgosValoresEstado
    'm_ColRiesgosValoresEstado("R001")="Detectado"
    
    Dim m_ColEvolucionRiesgos As Scripting.Dictionary
    
    Dim m_ColEvolucionPriorizacion As Scripting.Dictionary
    Dim m_ColEvolucionValoracion As Scripting.Dictionary
    Dim m_ColEvolucionEstado As Scripting.Dictionary
    
    Dim m_ColRiesgosValoresPriorizacion As Scripting.Dictionary
    Dim m_ColRiesgosValoresValoracion As Scripting.Dictionary
    Dim m_ColRiesgosValoresEstado As Scripting.Dictionary
    
    Dim m_URLExcel As String
    Dim wbLibro As Excel.Workbook
    Dim wbHoja As Excel.Worksheet
    Dim appExcel As Excel.Application
    Dim NumeroRiesgos As Long
    Dim NumeroEdiciones As Long
    Dim MiRango As Excel.Range
    Dim MiRangoTabla As Excel.Range
    Dim MiRangoTablaEquivalente As Excel.Range
    Dim MiRangoPosicionGrafica As Excel.Range
    
    On Error GoTo errores
    Avance "Obtención de datos para gráficas....."
    Set m_ColEvolucionRiesgos = getColEvolucionRiesgos(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColEvolucionRiesgos Is Nothing Then
        p_Error = "No se ha podido obtener la colección de las evoluciones de los riesgos"
        Err.Raise 1000
    End If
    Avance "Abriendo Excel para hacer los gráficos....."
    m_URLExcel = m_ObjEntorno.URLDirectorioLocal & "evolución Riesgos.xlsx"
    If FSO.FileExists(m_URLExcel) Then
        If FicheroAbierto(m_URLExcel) Then
            p_Error = "Parece que tiene abierto el documento"
            Err.Raise 1000
        End If
        FSO.DeleteFile m_URLExcel, True
    End If
    Set appExcel = New Excel.Application
    With appExcel
        .Visible = False
        .DisplayAlerts = False
        Set wbLibro = .Workbooks.Add
    End With
    
    Set wbHoja = wbLibro.Worksheets(1)
    AñadirCuadroTexto wbHoja, "evolución DE LOS RIESGOS " & Me.Proyecto
    '-------------------------------------------------------------------
    '  evolución DE LA PRIORIDAD
    '-------------------------------------------------------------------
    Avance "Datos de evolución de la prioridad....."
    'NumeroRiesgos = CLng(Me.EdicionUltimaNoActiva.colRiesgos.Count)
    NumeroRiesgos = CLng(Me.EdicionUltima.colRiesgos.Count)
    NumeroEdiciones = CLng(Me.ColEdicionesPublicadas.Count)
    Set m_ColEvolucionPriorizacion = m_ColEvolucionRiesgos("Priorización")
    Set MiRango = wbHoja.Cells(7, 3)
    Set MiRangoTabla = GenerarTabla(wbHoja, m_ColEvolucionPriorizacion, NumeroRiesgos, _
                        NumeroEdiciones, MiRango, "evolución de la Prioridad. Riesgos no Aceptados No Retirados", , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Avance "gráfico de evolución de la prioridad....."
    Set MiRangoPosicionGrafica = GenerarGrafico(MiRangoTabla, "evolución de la Prioridad. Riesgos no Aceptados No Retirados", , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '  evolución DEL IMPACTO
    '-------------------------------------------------------------------
    Avance "Datos de evolución de la valoración....."
    Set m_ColEvolucionValoracion = m_ColEvolucionRiesgos("Valoración")
    Set MiRango = wbHoja.Cells(MiRangoPosicionGrafica.Row + 2, 3)
    Set MiRangoTabla = GenerarTabla(wbHoja, m_ColEvolucionValoracion, NumeroRiesgos, _
                        NumeroEdiciones, MiRango, "evolución de la Valoración. Riesgos no Aceptados No Retirados", _
                        m_ObjEntorno.ColImpactoNumericoParaGraficas, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set MiRangoTablaEquivalente = GenerarTablaEquivalencias(wbHoja, m_ColEvolucionValoracion, NumeroRiesgos, _
                                NumeroEdiciones, MiRangoTabla, EnumSiNo.Sí, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Avance "gráfico de evolución de la Valoración....."
    Set MiRangoPosicionGrafica = GenerarGrafico(MiRangoTabla, "evolución de la Valoración. Riesgos no Aceptados No Retirados", MiRangoTablaEquivalente, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    '-------------------------------------------------------------------
    '  evolución DEL ESTADO
    '-------------------------------------------------------------------
    Avance "Datos de evolución del Estado....."
    Set m_ColEvolucionEstado = m_ColEvolucionRiesgos("Estado")
    Set MiRango = wbHoja.Cells(MiRangoPosicionGrafica.Row + 2, 3)
    Set MiRangoTabla = GenerarTabla(wbHoja, m_ColEvolucionEstado, NumeroRiesgos, _
                        NumeroEdiciones, MiRango, "evolución del Estado. Riesgos no Aceptados No Retirados", _
                        m_ObjEntorno.ColEstadoNumericoParaGraficas, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set MiRangoTablaEquivalente = GenerarTablaEquivalencias(wbHoja, m_ColEvolucionEstado, NumeroRiesgos, _
                        NumeroEdiciones, MiRangoTabla, EnumSiNo.No, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set MiRango = wbHoja.Cells(MiRangoTablaEquivalente.Row, MiRangoTablaEquivalente.Column)
    wbHoja.Columns(MiRangoTablaEquivalente.Column).EntireColumn.AutoFit
    Avance "gráfico de evolución del Estado....."
    Set MiRangoPosicionGrafica = GenerarGrafico(MiRangoTabla, "evolución del Estado. Riesgos no Aceptados No Retirados", MiRangoTablaEquivalente, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    wbLibro.SaveAs m_URLExcel
    wbLibro.Close
    Set wbLibro = Nothing
    appExcel.Quit
    Set appExcel = Nothing
    Avance "Mostrando el excel....."
    GenerarGraficosEvolucionRiesgos = m_URLExcel
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo GenerarGraficosEvolucionRiesgos ha devuelto el error: " & Err.Description
    End If
    If Not wbLibro Is Nothing Then
        wbLibro.Close False
        Set wbLibro = Nothing
    End If
    If Not appExcel Is Nothing Then
        appExcel.Quit
        Set appExcel = Nothing
    End If
End Function

Private Function getColEvolucionRiesgos( _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
                                        
    
    
    'getColEvolucionRiesgos("Priorización")=m_ColEvolucionPriorizacion
    'getColEvolucionRiesgos("Valoración")=m_ColEvolucionValoracion
    'getColEvolucionRiesgos("Estado")=m_ColEvolucionEstado
    
    'm_ColEvolucionPriorizacion("Edicion")=m_ColRiesgosValoresPriorizacion
    'm_ColRiesgosValoresPriorizacion("R001")=1
    
    'm_ColEvolucionValoracion("Edicion")=m_ColRiesgosValoresValoracion
    'm_ColRiesgosValoresValoracion("R001")="Alto"
    
    'm_ColEvolucionEstado("Edicion")=m_ColRiesgosValoresEstado
    'm_ColRiesgosValoresEstado("R001")="Detectado"
    
    Dim m_ColEvolucionPriorizacion As Scripting.Dictionary
    Dim m_ColEvolucionValoracion As Scripting.Dictionary
    Dim m_ColEvolucionEstado As Scripting.Dictionary
    
    Dim m_ColRiesgosValoresPriorizacion As Scripting.Dictionary
    Dim m_ColRiesgosValoresValoracion As Scripting.Dictionary
    Dim m_ColRiesgosValoresEstado As Scripting.Dictionary
    
    Dim m_objEdicion As Edicion
    Dim m_IDEd As Variant
    Dim m_ObjRiesgo As riesgo
    Dim m_Id As Variant
    
    On Error GoTo errores
    
    If Me.TieneEdiciones = EnumSiNo.No Then
        Exit Function
    End If
    If Me.ColEdicionesPublicadas Is Nothing Then
        p_Error = "No hay ediciones publicadas"
        Err.Raise 1000
    End If
    Set Me.ColEdicionesPublicadas = Nothing
    
    For Each m_IDEd In Me.ColEdicionesPublicadas
        Set m_objEdicion = Me.ColEdicionesPublicadas(m_IDEd)
        If m_objEdicion.TieneRiesgos = EnumSiNo.No Then
            GoTo siguienteEdicion
        End If
        
        For Each m_Id In m_objEdicion.colRiesgos
            Set m_ObjRiesgo = m_objEdicion.colRiesgos(m_Id)
            'If m_ObjRiesgo.CodigoRiesgo = "R004" Then Stop
            
            If m_ColRiesgosValoresPriorizacion Is Nothing Then
                Set m_ColRiesgosValoresPriorizacion = New Scripting.Dictionary
                m_ColRiesgosValoresPriorizacion.CompareMode = TextCompare
            End If
            If Not m_ColRiesgosValoresPriorizacion.Exists(m_ObjRiesgo.CodigoRiesgo) Then
                m_ColRiesgosValoresPriorizacion.Add m_ObjRiesgo.CodigoRiesgo, m_ObjRiesgo.Priorizacion
            End If
            If m_ColRiesgosValoresValoracion Is Nothing Then
                Set m_ColRiesgosValoresValoracion = New Scripting.Dictionary
                m_ColRiesgosValoresValoracion.CompareMode = TextCompare
            End If
            If Not m_ColRiesgosValoresValoracion.Exists(m_ObjRiesgo.CodigoRiesgo) Then
                If m_ObjRiesgo.Estado = "Aceptado" Or m_ObjRiesgo.Estado = "Retirado" Then
                    m_ColRiesgosValoresValoracion.Add m_ObjRiesgo.CodigoRiesgo, 0
                Else
                    m_ColRiesgosValoresValoracion.Add m_ObjRiesgo.CodigoRiesgo, m_ObjRiesgo.Valoracion
                End If
                
            End If
            If m_ColRiesgosValoresEstado Is Nothing Then
                Set m_ColRiesgosValoresEstado = New Scripting.Dictionary
                m_ColRiesgosValoresEstado.CompareMode = TextCompare
            End If
            If Not m_ColRiesgosValoresEstado.Exists(m_ObjRiesgo.CodigoRiesgo) Then
                m_ColRiesgosValoresEstado.Add m_ObjRiesgo.CodigoRiesgo, m_ObjRiesgo.Estado
            End If
            Set m_ObjRiesgo = Nothing
        Next
        
        If Not m_ColRiesgosValoresPriorizacion Is Nothing Then
            If m_ColEvolucionPriorizacion Is Nothing Then
                Set m_ColEvolucionPriorizacion = New Scripting.Dictionary
                m_ColEvolucionPriorizacion.CompareMode = TextCompare
            End If
            If Not m_ColEvolucionPriorizacion.Exists(m_objEdicion.Edicion) Then
                m_ColEvolucionPriorizacion.Add m_objEdicion.Edicion, m_ColRiesgosValoresPriorizacion
            End If
        End If
        If Not m_ColRiesgosValoresValoracion Is Nothing Then
            If m_ColEvolucionValoracion Is Nothing Then
                Set m_ColEvolucionValoracion = New Scripting.Dictionary
                m_ColEvolucionValoracion.CompareMode = TextCompare
            End If
            If Not m_ColEvolucionValoracion.Exists(m_objEdicion.Edicion) Then
                m_ColEvolucionValoracion.Add m_objEdicion.Edicion, m_ColRiesgosValoresValoracion
            End If
        End If
        If Not m_ColRiesgosValoresEstado Is Nothing Then
            If m_ColEvolucionEstado Is Nothing Then
                Set m_ColEvolucionEstado = New Scripting.Dictionary
                m_ColEvolucionEstado.CompareMode = TextCompare
            End If
            If Not m_ColEvolucionEstado.Exists(m_objEdicion.Edicion) Then
                m_ColEvolucionEstado.Add m_objEdicion.Edicion, m_ColRiesgosValoresEstado
            End If
        End If
siguienteEdicion:
        Set m_ColRiesgosValoresPriorizacion = Nothing
        Set m_ColRiesgosValoresValoracion = Nothing
        Set m_ColRiesgosValoresEstado = Nothing
        Set m_objEdicion = Nothing
    Next
    If Not m_ColEvolucionPriorizacion Is Nothing Then
        If getColEvolucionRiesgos Is Nothing Then
            Set getColEvolucionRiesgos = New Scripting.Dictionary
            getColEvolucionRiesgos.CompareMode = TextCompare
        End If
        If Not getColEvolucionRiesgos.Exists("Priorización") Then
            getColEvolucionRiesgos.Add "Priorización", m_ColEvolucionPriorizacion
        End If
    End If
    
    If Not m_ColEvolucionValoracion Is Nothing Then
        If getColEvolucionRiesgos Is Nothing Then
            Set getColEvolucionRiesgos = New Scripting.Dictionary
            getColEvolucionRiesgos.CompareMode = TextCompare
        End If
        If Not getColEvolucionRiesgos.Exists("Valoración") Then
            getColEvolucionRiesgos.Add "Valoración", m_ColEvolucionValoracion
        End If
    End If
    
    If Not m_ColEvolucionEstado Is Nothing Then
        If getColEvolucionRiesgos Is Nothing Then
            Set getColEvolucionRiesgos = New Scripting.Dictionary
            getColEvolucionRiesgos.CompareMode = TextCompare
        End If
        If Not getColEvolucionRiesgos.Exists("Estado") Then
            getColEvolucionRiesgos.Add "Estado", m_ColEvolucionEstado
        End If
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo getColEvolucionRiesgos ha devuelto el error: " & Err.Description
    End If
                                        
End Function

Private Function GenerarGrafico( _
                                    ByRef MiRangoGrafica As Excel.Range, _
                                    p_Titulo As String, _
                                    Optional ByRef RangoTablaEquivalente As Excel.Range, _
                                    Optional ByRef p_Error As String) As Excel.Range
    
    Dim MiChartObject As Excel.ChartObject
    Dim wbHoja As Excel.Worksheet
    Dim MiChart As Excel.Chart
    Dim i As Long
    Dim j As Long
    Dim MiRango As Excel.Range
    Dim MiRangoCabecera As Excel.Range
    Dim MiRangoEdiciones As Excel.Range
    Dim MiRangoDatos As Excel.Range
    Dim FilaCabecera As Long
    Dim FilaInicial As Long
    Dim FilaFinal As Long
    Dim columnaInicial As Long
    Dim columnaFinal As Long
    Dim ColumnaEdiciones As Long
    On Error GoTo errores
    
    Set wbHoja = MiRangoGrafica.Worksheet
    FilaCabecera = MiRangoGrafica.Row
    columnaInicial = MiRangoGrafica.Column + 1
    columnaFinal = columnaInicial + MiRangoGrafica.Columns.Count - 2
    Set MiRangoCabecera = wbHoja.Range(wbHoja.Cells(FilaCabecera, columnaInicial), _
                        wbHoja.Cells(FilaCabecera, columnaFinal))
    
    FilaInicial = FilaCabecera + 1
    FilaFinal = FilaCabecera + MiRangoGrafica.Rows.Count - 1
    ColumnaEdiciones = 2
    Set MiRangoEdiciones = wbHoja.Range(wbHoja.Cells(FilaInicial, ColumnaEdiciones), _
                wbHoja.Cells(FilaFinal, ColumnaEdiciones))
                
    wbHoja.Shapes.AddChart2(332, xlLineMarkers).Select
    Set MiChart = wbHoja.Parent.ActiveChart
    Set MiChartObject = getChartObject(wbHoja, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    i = 1
    For Each MiRango In MiRangoCabecera.Cells
        Set MiRangoDatos = wbHoja.Range(wbHoja.Cells(FilaInicial, MiRango.Column), _
                        wbHoja.Cells(FilaFinal, MiRango.Column))
        MiChart.SeriesCollection.NewSeries
        MiChart.FullSeriesCollection(i).Name = "=" & wbHoja.Name & "!" & MiRango.Address
        MiChart.FullSeriesCollection(i).Values = "=" & wbHoja.Name & "!" & MiRangoDatos.Address
        MiChart.FullSeriesCollection(i).XValues = "=" & wbHoja.Name & "!" & MiRangoEdiciones.Address
        i = i + 1
    Next
    With MiChart
        .HasTitle = True
        .ChartTitle.Caption = p_Titulo
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
    End With
    Set GenerarGrafico = ColocarGrafica(MiChartObject, MiRangoGrafica, RangoTablaEquivalente, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo GenerarGrafico ha devuelto el error: " & Err.Description
    End If
End Function
    
Private Function getChartObject( _
                                ByRef wbHoja As Excel.Worksheet, _
                                Optional ByRef p_Error As String _
                                ) As Excel.ChartObject
    
    On Error GoTo errores
    
    If wbHoja.ChartObjects.Count = 0 Then
        Exit Function
    End If
    Set getChartObject = wbHoja.ChartObjects(wbHoja.ChartObjects.Count)
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo getChartObject ha devuelto el error: " & Err.Description
    End If
    
End Function

Private Function getShape( _
                            ByRef wbHoja As Excel.Worksheet, _
                            Optional ByRef p_Error As String _
                            ) As Excel.Shape
    
    On Error GoTo errores
    
    If wbHoja.Shapes.Count = 0 Then
        Exit Function
    End If
    Set getShape = wbHoja.Shapes(wbHoja.Shapes.Count)
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo getShape ha devuelto el error: " & Err.Description
    End If
    
End Function
Private Function ColocarGrafica( _
                                ByRef p_ChartObject As Excel.ChartObject, _
                                ByRef RangoTabla As Excel.Range, _
                                Optional ByRef RangoTablaEquivalente As Excel.Range, _
                                Optional ByRef p_Error As String _
                                ) As Excel.Range
    
    Dim SuperioRango As Long
    Dim IzquierdaRango As Long
    Dim altoRango As Long
    Dim SuperiorTabla As Long
    Dim AltoTabla As Long
    Dim SuperiorTablaEquivalente As Long
    Dim AltoTablaEquivalente As Long
    Dim FinalTabla As Long
    Dim FinalTablaEquivalente As Long
    Dim TopGrafica As Long
    On Error GoTo errores
    
    SuperiorTabla = RangoTabla.Top
    AltoTabla = RangoTabla.Height
    FinalTabla = SuperiorTabla + AltoTabla
    
    If Not RangoTablaEquivalente Is Nothing Then
        SuperiorTablaEquivalente = RangoTablaEquivalente.Top
        AltoTablaEquivalente = RangoTablaEquivalente.Height
        FinalTablaEquivalente = SuperiorTablaEquivalente + AltoTablaEquivalente
    End If
    If FinalTabla >= FinalTablaEquivalente Then
        TopGrafica = FinalTabla
    Else
        TopGrafica = FinalTablaEquivalente
    End If
    IzquierdaRango = RangoTabla.Left
    altoRango = RangoTabla.Height
    p_ChartObject.Top = TopGrafica + 20
    p_ChartObject.Left = IzquierdaRango
    Set ColocarGrafica = p_ChartObject.BottomRightCell
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo ColocarGrafica ha devuelto el error: " & Err.Description
    End If
End Function

Private Function AñadirCuadroTexto(ByRef p_WbHoja As Excel.Worksheet, p_Texto As String, Optional ByRef p_Error As String) As String
    
    Dim MiShape As Excel.Shape
    Dim MiRango As Excel.Range
    On Error GoTo errores
    
    Set MiRango = p_WbHoja.Cells(2, 2)
    Set MiShape = p_WbHoja.Shapes.AddTextEffect(31, p_Texto, "+mn-lt", 54, 0, 0, MiRango.Left, MiRango.Top)
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo AñadirCuadroTexto ha devuelto el error: " & Err.Description
    End If
End Function

Private Function Recuadrar(ByRef p_Rango As Excel.Range, Optional ByRef p_Error As String) As String
    
    
    On Error GoTo errores
    p_Rango.Borders(xlDiagonalDown).LineStyle = xlNone
    p_Rango.Borders(xlDiagonalUp).LineStyle = xlNone
    With p_Rango.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With p_Rango.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With p_Rango.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With p_Rango.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With p_Rango.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With p_Rango.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo Recuadrar ha devuelto el error: " & Err.Description
    End If
End Function

Private Function RecuadrarEsquina(ByRef p_Rango As Excel.Range, Optional ByRef p_Error As String) As String
    
    
    On Error GoTo errores
    With p_Rango
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
    End With
    
    With p_Rango.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With p_Rango.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With p_Rango
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    

    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo RecuadrarEsquina ha devuelto el error: " & Err.Description
    End If
End Function
 
Private Function GenerarTabla( _
                            ByRef wbHoja As Excel.Worksheet, _
                            p_ColDatos As Scripting.Dictionary, _
                            NumeroRiesgos As Long, _
                            NumeroEdiciones As Long, _
                            RangoRef As Excel.Range, _
                            p_Titulo As String, _
                            Optional ByRef p_ColEquivalencia As Scripting.Dictionary, _
                            Optional ByRef p_Error As String _
                            ) As Excel.Range
    
    Dim i As Long
    Dim j As Long
    Dim m_ColRiesgosDatos As Scripting.Dictionary
    Dim MiRango As Excel.Range
    Dim MiRangoTitulo As Excel.Range
    Dim MiRangoEsquina As Excel.Range
    Dim MiRangoTabla As Excel.Range
    Dim valor As String
        
    With wbHoja
        Set MiRango = .Range(.Cells(RangoRef.Row + 2, 3), .Cells(RangoRef.Row + 2, 3))
        
        Set MiRangoEsquina = .Range(.Cells(MiRango.Row + 2, MiRango.Column - 1), .Cells(MiRango.Row + 2, MiRango.Column - 1))
        Set MiRangoTabla = .Range(.Cells(MiRangoEsquina.Row, MiRangoEsquina.Column), .Cells(MiRangoEsquina.Row + NumeroEdiciones, MiRangoEsquina.Column + NumeroRiesgos))
        With MiRango
            .Value = p_Titulo
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark2
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        Set MiRangoTitulo = .Range(.Cells(MiRango.Row, MiRango.Column), .Cells(MiRango.Row + 1, MiRango.Column + NumeroRiesgos - 1))
        MiRangoTitulo.MergeCells = True
        Set MiRango = MiRango.Offset(1, 0)
        For i = 1 To NumeroRiesgos
            With MiRango
                .Value = "R" & Format(i, "000")
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            Set MiRango = MiRango.Offset(0, 1)
        Next
        Set MiRango = MiRango.Offset(1, -NumeroRiesgos - 1)
        For j = 1 To NumeroEdiciones
            With MiRango
                MiRango.Value = "Ed" & Format(j, "00")
                MiRango.HorizontalAlignment = xlCenter
                MiRango.VerticalAlignment = xlCenter
            End With
            Set MiRango = MiRango.Offset(0, 1)
            If p_ColDatos.Exists(CStr(j)) Then
                Set m_ColRiesgosDatos = p_ColDatos(CStr(j))
                For i = 1 To NumeroRiesgos
                    
                    With MiRango
                        valor = m_ColRiesgosDatos("R" & Format(i, "000"))
                        If Not p_ColEquivalencia Is Nothing Then
                            If p_ColEquivalencia.Exists(valor) Then
                                valor = p_ColEquivalencia(valor)
                            End If
                        End If
                        MiRango.Value = valor
                        MiRango.HorizontalAlignment = xlCenter
                        MiRango.VerticalAlignment = xlCenter
                    End With
                    Set MiRango = MiRango.Offset(0, 1)
                Next
                Set MiRango = MiRango.Offset(1, -NumeroRiesgos - 1)
            Else
               ' Stop
            End If
            
            
        Next
    End With
    Recuadrar MiRangoTabla, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RecuadrarEsquina MiRangoEsquina, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set GenerarTabla = MiRangoTabla
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo GenerarTabla ha devuelto el error: " & Err.Description
    End If
End Function

Private Function GenerarTablaEquivalencias( _
                                            ByRef wbHoja As Excel.Worksheet, _
                                            p_ColDatos As Scripting.Dictionary, _
                                            NumeroRiesgos As Long, _
                                            NumeroEdiciones As Long, _
                                            RangoRef As Excel.Range, _
                                            p_EsValoracion As EnumSiNo, _
                                            Optional ByRef p_Error As String _
                                            ) As Excel.Range
    
    
    Dim m_ColEquivalencia As Scripting.Dictionary
    Dim MiRango As Excel.Range
    Dim MiRangoTabla As Excel.Range
    Dim valor As Variant
    Dim ValorEquivalente As String
    Dim Elementos As Long
    
    On Error GoTo errores
    
    If p_EsValoracion = 0 Then
        p_Error = "Se ha de incidar si es de valoración o estado"
        Err.Raise 1000
    End If
    If p_EsValoracion = EnumSiNo.Sí Then
        Set m_ColEquivalencia = m_ObjEntorno.ColImpactoNumericoParaGraficas
    Else
        Set m_ColEquivalencia = m_ObjEntorno.ColEstadoNumericoParaGraficas
    End If
    Elementos = m_ColEquivalencia.Count
    With wbHoja
        Set MiRango = .Cells(RangoRef.Row - 2, RangoRef.Columns.Count + 4)
        
        For Each valor In m_ColEquivalencia
            ValorEquivalente = m_ColEquivalencia(valor)
            With MiRango
                .Value = valor
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            Set MiRango = MiRango.Offset(0, 1)
            With MiRango
                .Value = ValorEquivalente
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            Set MiRango = MiRango.Offset(1, -1)
        Next
        Set MiRango = .Cells(RangoRef.Row - 2, RangoRef.Columns.Count + 4)
        Set MiRangoTabla = MiRango.CurrentRegion
    End With
    Recuadrar MiRangoTabla, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set GenerarTablaEquivalencias = MiRangoTabla
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo GenerarTablaEquivalencias ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Cerrar( _
                        Optional ByVal p_db As DAO.Database, _
                        Optional ByRef p_Error As String _
                        ) As String
    
    On Error GoTo errores
    
    m_Error = ""
    Me.Error = ""
    p_Error = ""
    
    CerrarProyectoTransaccional Me, p_db, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo Proyecto.Cerrar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function


Public Function Borrar( _
                        Optional ByVal p_db As DAO.Database, _
                        Optional ByRef p_Error As String _
                        ) As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    ValidacionEliminar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    BorrarProyectoTransaccional Me, p_db, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set m_ObjProyectoActivo = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo Borrar ha devuelto el error: " & Err.Description
    End If
End Function





Public Function UsuarioAutorizado( _
                                Optional p_Usuario As Usuario, _
                                Optional ByRef p_Error As String _
                                ) As EnumSiNo
    On Error GoTo errores
    If p_Usuario Is Nothing Then
        Set p_Usuario = m_ObjUsuarioConectado
    End If
    If Not m_ObjEntorno.ColUsuariosAdministradores Is Nothing Then
        If m_ObjEntorno.ColUsuariosAdministradores.Exists(p_Usuario.UsuarioRed) Then
            UsuarioAutorizado = EnumSiNo.Sí
            Exit Function
        End If
    End If
    If Not m_ObjEntorno.ColUsuariosCalidad Is Nothing Then
        If m_ObjEntorno.ColUsuariosCalidad.Exists(p_Usuario.UsuarioRed) Then
            UsuarioAutorizado = EnumSiNo.Sí
            Exit Function
        End If
    End If
    If Me.CadenaNombreAutorizados = "" Then
        UsuarioAutorizado = EnumSiNo.No
        Exit Function
    End If
    If InStr(1, Me.CadenaNombreAutorizados, p_Usuario.Nombre) = 0 Then
        UsuarioAutorizado = EnumSiNo.No
    Else
        UsuarioAutorizado = EnumSiNo.Sí
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo Proyecto.UsuarioAutorizado ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Registrar( _
                            Optional ByRef p_ObjProyectoAlInicio As Proyecto, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    p_Error = ""
    
    If p_ObjProyectoAlInicio Is Nothing Then
        CrearProyectoTransaccional Me, , , p_Error
    Else
        EditarProyectoTransaccional Me, p_ObjProyectoAlInicio, , p_Error
    End If
    
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo Registrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function RegistrarCorreoRAC( _
                                    p_CorreoRAC As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    p_Error = ""
    
    RegistrarCorreoRACTransaccional Me, p_CorreoRAC, , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo RegistrarCorreoRAC ha devuelto el error: " & Err.Description
    End If
    
End Function
Public Function EstablecerFechaMaxPublicacion( _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    p_Error = ""
    
    EstablecerFechaMaxPublicacionTransaccional Me, , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo EstablecerFechaMaxPublicacion ha devuelto el error: " & Err.Description
    End If
    
End Function

Public Function GrabarCambiosEnProyecto( _
                                            Optional ByVal p_DesdeEdicion As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    On Error GoTo errores
    
    GrabarCambiosEnProyectoTransaccional Me, p_DesdeEdicion, , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo GrabarCambiosEnProyecto ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function BorrarCambiosProyecto( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    On Error GoTo errores
    
    BorrarCambiosProyectoTransaccional Me, , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set Me.colCambios = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo BorrarCambiosProyecto ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function FechaMaxProximaPublicacionGrabar( _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    On Error GoTo errores
    
    Me.FechaMaxProximaPublicacion = Me.FechaMaxProximaPublicacionCalculada
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    FechaMaxProximaPublicacionGrabarTransaccional Me, , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo FechaMaxProximaPublicacionGrabar ha devuelto el error: " & Err.Description
    End If
End Function
Private Function getHTMLTablaSuministradoresConGestionRiesgos( _
                                                                Optional ByRef p_Error As String _
                                                                ) As String
    Dim m_mensaje As String
    Dim m_objColSuministradores As Scripting.Dictionary
    Dim m_Id As Variant
    Dim m_ProyectoSuministrador As ProyectoSuministrador
    Dim m_Suministrador As Suministrador
    Dim m_NombreSuministrador As String
    Dim m_GestionCalidad As String
    
    
    On Error GoTo errores
    
    
    Set m_objColSuministradores = Me.ColSuministradores
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_mensaje = m_mensaje & "<table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""ColespanArriba"" colspan='2'>Suministradores con gestión de riesgos compartida </td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Suministrador</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Gestión de la Calidad</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        If m_objColSuministradores Is Nothing Then
            m_mensaje = m_mensaje & "</table>" & vbNewLine
            getHTMLTablaSuministradoresConGestionRiesgos = m_mensaje
            Exit Function
        End If
        If m_objColSuministradores.Count = 0 Then
            m_mensaje = m_mensaje & "</table>" & vbNewLine
            getHTMLTablaSuministradoresConGestionRiesgos = m_mensaje
            Exit Function
        End If
        For Each m_Id In m_objColSuministradores
            Set m_ProyectoSuministrador = m_objColSuministradores(m_Id)
            m_GestionCalidad = m_ProyectoSuministrador.GestionCalidad
            Set m_Suministrador = m_ProyectoSuministrador.Suministrador
            If Not m_Suministrador Is Nothing Then
                m_NombreSuministrador = m_Suministrador.Nombre
            End If
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
                m_mensaje = m_mensaje & "<td> " & m_NombreSuministrador & " </td>" & vbNewLine
                m_mensaje = m_mensaje & "<td> " & m_GestionCalidad & " </td>" & vbNewLine
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
            Set m_Suministrador = Nothing
            Set m_ProyectoSuministrador = Nothing
        Next
    m_mensaje = m_mensaje & "</table>" & vbNewLine
    
    getHTMLTablaSuministradoresConGestionRiesgos = m_mensaje
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo getHTMLTablaSuministradoresConGestionRiesgos ha devuelto el error: " & Err.Description
    End If
End Function

Public Function RegistrarSuministrador( _
                                        p_IDSuministrador As String, _
                                        p_GestionCalidad As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    p_Error = ""
    
    RegistrarSuministradorTransaccional Me, p_IDSuministrador, p_GestionCalidad, , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo RegistrarSuministrador ha devuelto el error: " & Err.Description
    End If
    
End Function
    

    
Public Function MotivoNoOK( _
                                Optional ByRef p_ObjProyectoAlInicio As Proyecto, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_CampoRepetido As EnumSiNo
    Dim blnComprobarCampo As Boolean
    Dim m_AlmenosUnaEdicionPublicada As EnumSiNo
    Dim m_Expediente As Expediente
    Dim m_ProyectoConEseExpediente As Proyecto
    On Error GoTo errores
    
    With Me
        If .UsuarioAutorizado = EnumSiNo.No Then
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 100
            End If
            MotivoNoOK = "El perfil del usuario no está autorizado a realizar la operación"
            Exit Function
        End If
        Set m_Expediente = .Expediente
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_Expediente Is Nothing Then
            MotivoNoOK = "No se han vinculado con los datos del Expediente"
            Exit Function
        End If
        If .Proyecto = "" Then
            MotivoNoOK = "No se ha podido obtener el nombre del proyecto"
            Exit Function
        End If
        If .Juridica = "" Then
            MotivoNoOK = "No se ha podido obtener la jurídica"
            Exit Function
        End If
        If .NombreProyecto = "" Then
            MotivoNoOK = "No se ha podido obtener el nemotécnico del proyecto"
            Exit Function
        End If
        
        If .Elaborado = "" Then
            MotivoNoOK = "Falta Elaborado"
            Exit Function
        End If
        If .Revisado = "" Then
            MotivoNoOK = "Falta Revisado"
            Exit Function
        End If
        If .Aprobado = "" Then
            MotivoNoOK = "Falta Aprobado"
            Exit Function
        End If
        If .Elaborado = .Revisado Then
            MotivoNoOK = "El Usuario que Elabora no puede ser el mismo que revisa"
            Exit Function
        End If
        If .CodigoDocumento = "" Then
            MotivoNoOK = "No se ha podido obtener el código de documento"
            Exit Function
        End If
        If .ParaInformeAvisos <> "Sí" And .ParaInformeAvisos <> "No" Then
            MotivoNoOK = "Se ha de indicar si es para informes"
            Exit Function
        End If
        If .NombreUsuarioCalidad = "" Then
            MotivoNoOK = "Se ha de indicar el Nombre del Responsable de Calidad"
            Exit Function
        End If
        If .EnUTE <> "Sí" And .EnUTE <> "No" Then
            MotivoNoOK = "Se ha de indicar si se trabaja en UTE o no"
            Exit Function
        End If
        If .CadenaNombreAutorizados = "" Then
            .CadenaNombreAutorizados = .CadenaNombreAutorizadosCalculados
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If .CadenaNombreAutorizados = "" Then
                MotivoNoOK = "El expediente del que depende no tiene registrados ningún usuario autorizado. Hable con el gestor"
                Exit Function
            End If
        End If
        
        blnComprobarCampo = False
        If .ParaInformeAvisos <> "No" Then
            If p_ObjProyectoAlInicio Is Nothing Then
                blnComprobarCampo = True
            Else
                If p_ObjProyectoAlInicio.IDExpediente <> .IDExpediente Then
                    blnComprobarCampo = True
                End If
            End If
        End If
        If blnComprobarCampo = True Then
            Set m_ProyectoConEseExpediente = Constructor.getProyectoPorExpediente(p_IDExpediente:=.IDExpediente, p_Error:=p_Error)
            If p_Error <> "" Then
                Err.Raise 100
            End If
            If Not m_ProyectoConEseExpediente Is Nothing Then
                MotivoNoOK = "El proyecto es para Informe y ese expediente ya ha sido usado en otro"
                Exit Function
            End If
        End If
        
        
    End With
    
    
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El mëtodo MotivoNoOK ha devuelto el error: " & Err.Description
     End If
    
End Function

Private Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String

    
    On Error GoTo errores
    
    If Not Me.EdicionUltima Is Nothing Then
        If Me.EdicionUltima.IDEdicion = Me.EdicionPrimera.IDEdicion Then
            If Me.EdicionPrimera.FechaPublicacion = "" Then
                If Me.UsuarioConectadoAutorizado = EnumSiNo.No Then
                    p_Error = Me.Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    p_Error = "El perfil de su usuario no tiene autorización para realizar esta operación"
                    Err.Raise 1000
                End If
            Else
                If EsTecnico = EnumSiNo.Sí Then
                    p_Error = "El perfil de su usuario no tiene autorización para realizar esta operación"
                    Err.Raise 1000
                End If
            End If
        Else
            If EsTecnico = EnumSiNo.Sí Then
                p_Error = "El perfil de su usuario no tiene autorización para realizar esta operación"
                Err.Raise 1000
            End If
        End If
    Else
        If Me.UsuarioConectadoAutorizado = EnumSiNo.No Then
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            p_Error = "El perfil de su usuario no tiene autorización para realizar esta operación"
            Err.Raise 1000
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo ValidacionEliminar ha devuelto el error: " & Err.Description
    End If
End Function


Public Function DescripcionEnRiesgoNoTrasladadoRepetida( _
                                                            p_Descripcion As String, _
                                                            Optional ByRef p_Error As String _
                                                            ) As EnumSiNo
                               
    Dim m_Id As Variant
    Dim m_ObjRiesgoExt As RiesgoExterno
    On Error GoTo errores
    
    
    If Me.TieneRiesgosDeOferta = EnumSiNo.No Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        DescripcionEnRiesgoNoTrasladadoRepetida = EnumSiNo.No
        Exit Function
    End If
    For Each m_Id In Me.ColRiesgosOferta.keys
        If Nz(m_Id, "") <> "" Then
            Set m_ObjRiesgoExt = Me.ColRiesgosOferta(m_Id)
            If m_ObjRiesgoExt.riesgo Is Nothing Then
                If p_Descripcion = m_ObjRiesgoExt.Descripcion Then
                    Set m_ObjRiesgoExt = Nothing
                    DescripcionEnRiesgoNoTrasladadoRepetida = EnumSiNo.Sí
                    Exit Function
                End If
            End If
            Set m_ObjRiesgoExt = Nothing
        End If

    Next
    DescripcionEnRiesgoNoTrasladadoRepetida = EnumSiNo.No
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo DescripcionEnRiesgoNoTrasladadoRepetida ha devuelto el error: " & Err.Description
    End If
End Function

Public Function DescripcionYCausaRaizEnRiesgoNoTrasladadoRepetida( _
                                                                    p_Descripcion As String, _
                                                                    p_CausaRaiz As String, _
                                                                    Optional ByRef p_RiesgoAEvitar As riesgo, _
                                                                    Optional ByRef p_Error As String _
                                                                    ) As EnumSiNo
                               
    Dim m_Id As Variant
    Dim m_ObjRiesgoExt As RiesgoExterno
    On Error GoTo errores
    
    
    If Me.TieneRiesgosDeOferta = EnumSiNo.No Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        DescripcionYCausaRaizEnRiesgoNoTrasladadoRepetida = EnumSiNo.No
        Exit Function
    End If
    For Each m_Id In Me.ColRiesgosOferta.keys
        If Nz(m_Id, "") <> "" Then
            Set m_ObjRiesgoExt = Me.ColRiesgosOferta(m_Id)
            If m_ObjRiesgoExt.riesgo Is Nothing Then
                If p_Descripcion = m_ObjRiesgoExt.Descripcion And p_CausaRaiz = m_ObjRiesgoExt.CausaRaiz Then
                    Set m_ObjRiesgoExt = Nothing
                    DescripcionYCausaRaizEnRiesgoNoTrasladadoRepetida = EnumSiNo.Sí
                    Exit Function
                End If
            End If
siguiente:
            Set m_ObjRiesgoExt = Nothing
        End If

    Next
    DescripcionYCausaRaizEnRiesgoNoTrasladadoRepetida = EnumSiNo.No
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo DescripcionYCausaRaizEnRiesgoNoTrasladadoRepetida ha devuelto el error: " & Err.Description
    End If
End Function

Private Function getCodigoDocumento( _
                                    p_CodigoDocPorExpediente As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                     
    
    Dim m_ObjDocumento As Documento
    Dim m_ObjColDocumentos As Scripting.Dictionary
    Dim m_IDDocumento As Variant
    Dim m_EdicionMaxima As Integer
    
    
    On Error GoTo errores
    
    If p_CodigoDocPorExpediente = "" Then
        Exit Function
    End If
    Set m_ObjColDocumentos = Constructor.getDocumentosbyCodigo(p_CodigoDocPorExpediente, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjColDocumentos Is Nothing Then
        For Each m_IDDocumento In m_ObjColDocumentos.keys
            If IsObject(m_ObjColDocumentos(m_IDDocumento)) Then
                Set m_ObjDocumento = m_ObjColDocumentos(m_IDDocumento)
                If IsNumeric(m_ObjDocumento.Edicion) Then
                    If CInt(m_ObjDocumento.Edicion) > m_EdicionMaxima Then
                        m_EdicionMaxima = CInt(m_ObjDocumento.Edicion)
                    End If
                End If
                Set m_ObjDocumento = Nothing
            End If
        Next
    End If
    getCodigoDocumento = p_CodigoDocPorExpediente & Format(m_EdicionMaxima + 1, "00")
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo getCodigoDocumento ha devuelto el error: " & Err.Description
    End If
End Function


Public Property Get ColCampos() As Collection

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColCampos Is Nothing Then
        Set ColCampos = m_objColCampos
        Exit Property
    End If
    Set m_objColCampos = New Collection
    With m_objColCampos
        .Add "IDProyecto"
        .Add "IDExpediente"
        .Add "Proyecto"
        .Add "Juridica"
        .Add "NombreProyecto"
        .Add "Cliente"
        .Add "FechaPrevistaCierre"
        .Add "FechaCierre"
        .Add "FechaRegistroInicial"
        .Add "Elaborado"
        .Add "Revisado"
        .Add "Aprobado"
        .Add "CodigoDocumento"
        .Add "ParaInformeAvisos"
        .Add "FechaFirmaContrato"
        .Add "RiesgosDeLaOferta"
        .Add "RiesgosDelSubContratista"
        .Add "NombreUsuarioCalidad"
        .Add "EnUTE"
        .Add "FechaMaxProximaPublicacion"
        .Add "RequiereRiesgoDeBiblioteca"
        .Add "CorreoRAC"
        .Add "Ordinal"
        .Add "CadenaNombreAutorizados"
        .Add "NombreParaNodo"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Get ColCamposParaCambios() As Collection

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColCamposParaCambios Is Nothing Then
        Set ColCamposParaCambios = m_objColCamposParaCambios
        Exit Property
    End If
    Set m_objColCamposParaCambios = New Collection
    With m_objColCamposParaCambios
        .Add "Elaborado"
        .Add "Revisado"
        .Add "Aprobado"
        .Add "CodigoDocumento"
        .Add "EnUTE"
        .Add "CorreoRAC"
        .Add "FechaCierre"
        
    End With
    Set ColCamposParaCambios = m_objColCamposParaCambios
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El mëtodo Proyecto.ColCamposParaCambios ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDProyecto" Then
        getPropiedad = Me.IDProyecto
        Exit Function
    End If
    If m_NombrePropiedad = "IDExpediente" Then
        getPropiedad = Me.IDExpediente
        Exit Function
    End If
    If m_NombrePropiedad = "Juridica" Then
        getPropiedad = Me.Juridica
        Exit Function
    End If
    If m_NombrePropiedad = "NombreProyecto" Then
        getPropiedad = Me.NombreProyecto
        Exit Function
    End If
    If m_NombrePropiedad = "Cliente" Then
        getPropiedad = Me.Cliente
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrevistaCierre" Then
        getPropiedad = Me.FechaPrevistaCierre
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCierre" Then
        getPropiedad = Me.FechaCierre
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRegistroInicial" Then
        getPropiedad = Me.fechaRegistroInicial
        Exit Function
    End If
    If m_NombrePropiedad = "Elaborado" Then
        getPropiedad = Me.Elaborado
        Exit Function
    End If
    If m_NombrePropiedad = "Revisado" Then
        getPropiedad = Me.Revisado
        Exit Function
    End If
    If m_NombrePropiedad = "Aprobado" Then
        getPropiedad = Me.Aprobado
        Exit Function
    End If
    If m_NombrePropiedad = "CodigoDocumento" Then
        getPropiedad = Me.CodigoDocumento
        Exit Function
    End If
    If m_NombrePropiedad = "ParaInformeAvisos" Then
        getPropiedad = Me.ParaInformeAvisos
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFirmaContrato" Then
        getPropiedad = Me.FechaFirmaContrato
        Exit Function
    End If
    If m_NombrePropiedad = "RiesgosDeLaOferta" Then
        getPropiedad = Me.RiesgosDeLaOferta
        Exit Function
    End If
    If m_NombrePropiedad = "RiesgosDeLaOferta" Then
        getPropiedad = Me.RiesgosDeLaOferta
        Exit Function
    End If
    If m_NombrePropiedad = "NombreUsuarioCalidad" Then
        getPropiedad = Me.NombreUsuarioCalidad
        Exit Function
    End If
    If m_NombrePropiedad = "EnUTE" Then
        getPropiedad = Me.EnUTE
        Exit Function
    End If
    If m_NombrePropiedad = "FechaMaxProximaPublicacion" Then
        getPropiedad = Me.FechaMaxProximaPublicacion
        Exit Function
    End If
    If m_NombrePropiedad = "RequiereRiesgoDeBiblioteca" Then
        getPropiedad = Me.RequiereRiesgoDeBiblioteca
        Exit Function
    End If
    If m_NombrePropiedad = "CorreoRAC" Then
        getPropiedad = Me.CorreoRAC
        Exit Function
    End If
    If m_NombrePropiedad = "Ordinal" Then
        getPropiedad = Me.Ordinal
        Exit Function
    End If
    If m_NombrePropiedad = "CadenaNombreAutorizados" Then
        getPropiedad = Me.CadenaNombreAutorizados
        Exit Function
    End If
    If m_NombrePropiedad = "NombreParaNodo" Then
        getPropiedad = Me.NombreParaNodo
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Proyecto.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo Proyecto.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDProyecto" Then
        Me.IDProyecto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDExpediente" Then
        Me.IDExpediente = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Proyecto" Then
        Me.Proyecto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Juridica" Then
        Me.Juridica = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreProyecto" Then
        Me.NombreProyecto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Cliente" Then
        Me.Cliente = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrevistaCierre" Then
        Me.FechaPrevistaCierre = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRegistroInicial" Then
        Me.fechaRegistroInicial = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCierre" Then
        Me.FechaCierre = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "Elaborado" Then
        Me.Elaborado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Revisado" Then
        Me.Revisado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Aprobado" Then
        Me.Aprobado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodigoDocumento" Then
        Me.CodigoDocumento = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ParaInformeAvisos" Then
        Me.ParaInformeAvisos = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFirmaContrato" Then
        Me.FechaFirmaContrato = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "RiesgosDeLaOferta" Then
        Me.RiesgosDeLaOferta = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "RiesgosDelSubContratista" Then
        Me.RiesgosDelSubContratista = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EnUTE" Then
        Me.EnUTE = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreUsuarioCalidad" Then
        Me.NombreUsuarioCalidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaMaxProximaPublicacion" Then
        Me.FechaMaxProximaPublicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "RequiereRiesgoDeBiblioteca" Then
        Me.RequiereRiesgoDeBiblioteca = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CorreoRAC" Then
        Me.CorreoRAC = m_ValorPropiedad
        Exit Function
    End If
    
    
    If m_NombrePropiedad = "Ordinal" Then
        Me.Ordinal = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CadenaNombreAutorizados" Then
        Me.CadenaNombreAutorizados = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreParaNodo" Then
        Me.NombreParaNodo = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Proyecto.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo Proyecto.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function






