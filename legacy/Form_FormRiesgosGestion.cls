VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormRiesgosGestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'````````````````````````````````````````````````````````````````````````````
'`                    Code from Tassos Filoxenidis                          `
'````````````````````````````````````````````````````````````````````````````

' ===============================
'  Separador (layout)
' ===============================

' Valores por defecto / límites
Private Const SepWMin As Long = 60      ' ancho mínimo del árbol
Private Const SepW As Long = 60         ' ancho inicial por defecto del árbol

' Estado del drag
Private m_sngLeftResizePos As Single

' Geometría del formulario
Private MeWidth As Long
Private MeWidthHalf As Long
Private ArbolLeft As Long
Private SubFormLeft As Long

Private Const C_APP As String = "GESTION_RIESGOS"
Private Const C_SECTION As String = "FormRiesgosGestion"
Private Const C_KEY_ARbolW As String = "ArbolWidth"

Private WithEvents m_FormRiesgo As Form_FormRiesgo



Public m_NodoSeleccionado As MSComctlLib.Node
Private m_URLInformeRiesgo As String
Public m_ObjSeleccionado As Object
Public m_RiesgoSeleccionado As riesgo
Public m_EstadoRiesgoSeleccionado As EnumRiesgoEstado
Public m_PMSeleccionado As PM
Public m_PMAccionSeleccionado As PMAccion
Public m_PCSeleccionado As PC
Public m_PCAccionSeleccionado As PCAccion
Public m_EsMitigacion As EnumSiNo
Public m_EsEdicionActiva As EnumSiNo
Public m_NodoEdicion As MSComctlLib.Node
Public m_Arbol As MSComctlLib.TreeView
Private m_ListImages As Object
Private m_SinCargarFormulario As EnumSiNo
Public m_ColRiesgosAplicados As Scripting.Dictionary
Public m_RiesgoUltimoSeleccionado As riesgo
Public blnPermitidoEditar As Boolean

Public m_EdicionAlAbrir As Edicion

Private m_Error As String

Private Sub ComandoPublicacion_Click()
    
    Dim m_NombreFormulario As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If EsTecnico = EnumSiNo.Sí Then
        m_NombreFormulario = "FormPublicacionTecnico"
    Else
        m_NombreFormulario = "FormPublicacionCalidad"
    End If
    If FormularioAbierto(m_NombreFormulario) Then
        DoCmd.Close acForm, m_NombreFormulario, acSaveNo
    End If
    DoCmd.OpenForm m_NombreFormulario
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoPublicacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
Private Sub ComandoRiesgoPriorizacion_Click()
    On Error GoTo errores
    
    
    
    
    m_Error = ""
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If FormularioAbierto("FormRiesgosEstablecerPrioridades") Then
        DoCmd.Close acForm, "FormRiesgosEstablecerPrioridades", acSaveNo
    End If
    DoCmd.OpenForm "FormRiesgosEstablecerPrioridades"
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRiesgoPriorizacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim m_ColEdiciones As Scripting.Dictionary
    Dim m_IDEdicion As Variant
    Dim m_Edicion As Edicion
    Dim cmb As ComboBox
    
    On Error GoTo errores
    m_Error = ""
    Set cmb = Me.ComboEdicion
    cmb.RowSource = ""
    Me.ComandoAnexos.Enabled = False
    Me.AllowEdits = False
    On Error GoTo errores
    
    If Not m_EdicionAlAbrir Is Nothing Then
        Set m_ObjProyectoActivo = m_EdicionAlAbrir.Proyecto
        Set m_ObjEdicionActiva = m_EdicionAlAbrir
    Else
        If m_ObjProyectoActivo Is Nothing Then
            If m_ObjEdicionActiva Is Nothing Then
                p_Error = "No se sabe la gestión de riesgos activa"
                Err.Raise 1000
            End If
            Set m_ObjProyectoActivo = m_ObjEdicionActiva.Proyecto
            If m_ObjEdicionActiva Is Nothing Then
                p_Error = "No se sabe la gestión de riesgos activa"
                Err.Raise 1000
            End If
        Else
            Set m_ObjEdicionActiva = m_ObjProyectoActivo.EdicionUltima
        End If
    End If
    
    
    Set m_ColEdiciones = m_ObjProyectoActivo.colEdiciones
    If Not m_ColEdiciones Is Nothing Then
        For Each m_IDEdicion In m_ColEdiciones
            Set m_Edicion = m_ColEdiciones(m_IDEdicion)
            cmb.AddItem m_Edicion.Edicion
            Set m_Edicion = Nothing
        Next
    End If
   
    
    m_EsEdicionActiva = m_ObjEdicionActiva.EsActivo
    Me.ComboEdicion = m_ObjEdicionActiva.Edicion
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado & " " & "riesgos de la edición " & m_ObjEdicionActiva.Edicion & " de " & m_ObjEdicionActiva.Proyecto.Proyecto
    
    
    If m_ObjEntorno.VerSoloRiesgosNoRetirados = EnumSiNo.Sí Then
        Me.ComboVerRetirados = "No"
    Else
        Me.ComboVerRetirados = "Sí"
    End If
    If m_ObjEntorno.VerRiesgosDescripcion = EnumSiNo.Sí Then
        Me.ComboVerDescripcion = "Sí"
    Else
        Me.ComboVerDescripcion = "No"
    End If
    
    
    
    
    
    Set m_Arbol = Me.Arbol.Object
    Set m_ListImages = Me.ListaImagenes.Object
    RellenarListaImagenes m_ListImages, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RegistrarUltimoProyecto m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CargarArbol p_Refrescando:=EnumSiNo.No, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If EsTecnico = EnumSiNo.Sí Then
        Me.ComandoEditarProyecto.Visible = False
    Else
        Me.ComandoEditarProyecto.Visible = True
        Me.ComandoEditarProyecto.Enabled = True
    End If
    If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
        If m_ObjEdicionActiva.UsuarioConectadoAutorizado = EnumSiNo.Sí Then
            blnPermitidoEditar = True
        Else
            blnPermitidoEditar = False
        End If
    Else
        blnPermitidoEditar = False
    End If
    If blnPermitidoEditar = True Then
        Me.ComandoRiesgoPriorizacion.Enabled = True
    Else
        Me.ComandoRiesgoPriorizacion.Enabled = False
    End If
    Me.ComandoPublicacion.Enabled = True
    
    ActualizarEtiquetaUltimoCambio
    
    Me.AllowEdits = True
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error: " & vbNewLine & Err.Description
    End If
    Me.AllowEdits = True
End Function
Public Function ActualizarEtiquetaUltimoCambio(Optional ByRef p_Error As String) As String
    If m_ObjEdicionActiva.UsuarioUltimoCambio <> "" Then
        Me.lblUltimoCambio.Visible = True
        If Not m_ObjEdicionActiva.UsuarioUltimoCambioObj Is Nothing Then
            Me.lblUltimoCambio.Caption = "Último cambio " & m_ObjEdicionActiva.FechaUltimoCambio & " por " & m_ObjEdicionActiva.UsuarioUltimoCambioObj.Nombre
        Else
            Me.lblUltimoCambio.Caption = "Último cambio " & m_ObjEdicionActiva.FechaUltimoCambio & " por desconocido"
        End If
        
    Else
        Me.lblUltimoCambio.Visible = False
    End If
End Function

Public Sub Arbol_NodeClick(ByVal Node As Object)
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    Me.AllowEdits = False
    Me.ComandoAnexos.Enabled = False
    m_URLInformeRiesgo = ""
    Set m_NodoSeleccionado = Node
    Set m_ObjSeleccionado = getObjetoSeleccionadoDeRiesgos(Node.Key, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If TypeOf m_ObjSeleccionado Is Edicion Then
        If Me.FormDetalle.SourceObject <> "" Then
            Me.FormDetalle.SourceObject = ""
        End If
        Set m_RiesgoUltimoSeleccionado = Nothing
        Me.lblEstadoRiesgo.Visible = False
        Me.ComandoAnexos.Enabled = True
        Set m_ObjEdicionActiva = m_ObjSeleccionado
        If m_ObjEdicionActiva.EsActivo = EnumSiNo.No Then
            blnPermitidoEditar = False
        Else
            If m_ObjEdicionActiva.Proyecto.UsuarioAutorizado = EnumSiNo.Sí Then
                blnPermitidoEditar = True
            
            Else
                blnPermitidoEditar = True
            End If
        End If
        
        
        If m_SinCargarFormulario <> EnumSiNo.Sí Then
            Me.FormDetalle.SourceObject = "FormRiesgosGestionEdicion"
        End If
        
        AvanceCerrar
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    If TypeOf m_ObjSeleccionado Is riesgo Then
        Set m_RiesgoSeleccionado = m_ObjSeleccionado
        m_EstadoRiesgoSeleccionado = m_RiesgoSeleccionado.EstadoEnum
        Set m_ObjRiesgoActivo = m_RiesgoSeleccionado
        m_EstadoRiesgoActivo = m_EstadoRiesgoSeleccionado
        If m_SinCargarFormulario <> EnumSiNo.Sí Then
            Me.FormDetalle.SourceObject = "FormRiesgosGestionRiesgo"
        End If
        If m_RiesgoUltimoSeleccionado Is Nothing Then
            If m_ObjEdicionActiva.EsActivo Then
                RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_EstadoRiesgoSeleccionado, m_Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                Me.lblEstadoRiesgo.Visible = True
            Else
                Me.lblEstadoRiesgo.Visible = False
            End If
            Set m_RiesgoUltimoSeleccionado = Constructor.getRiesgo(m_RiesgoSeleccionado.IDRiesgo)
            
        Else
            If m_RiesgoSeleccionado.IDRiesgo <> m_RiesgoUltimoSeleccionado.IDRiesgo Then
                If m_ObjEdicionActiva.EsActivo Then
                    RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_EstadoRiesgoSeleccionado, m_Error
                    If m_Error <> "" Then
                        Err.Raise 1000
                    End If
                    Me.lblEstadoRiesgo.Visible = True
                Else
                    Me.lblEstadoRiesgo.Visible = False
                End If
                Set m_RiesgoUltimoSeleccionado = Constructor.getRiesgo(m_RiesgoSeleccionado.IDRiesgo)
            End If
        End If
        AvanceCerrar
        Me.ComandoAnexos.Enabled = True
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    If TypeOf m_ObjSeleccionado Is PM Then
        Set m_PMSeleccionado = m_ObjSeleccionado
        Me.m_EsMitigacion = EnumSiNo.Sí
        Set m_PMSeleccionado = m_ObjSeleccionado
        Set m_ObjPMActivo = m_PMSeleccionado
        
        Set m_ObjRiesgoActivo = m_PMSeleccionado.riesgo
        If m_SinCargarFormulario <> EnumSiNo.Sí Then
            Me.FormDetalle.SourceObject = "FormRiesgosGestionPlanPrincipal"
        End If
        If m_RiesgoUltimoSeleccionado Is Nothing Then
            If m_ObjEdicionActiva.EsActivo Then
                RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_ObjRiesgoActivo.EstadoEnum, m_Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                Me.lblEstadoRiesgo.Visible = True
            Else
                Me.lblEstadoRiesgo.Visible = False
            End If
           
            Set m_RiesgoUltimoSeleccionado = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo)
        Else
            If m_ObjRiesgoActivo.IDRiesgo <> m_RiesgoUltimoSeleccionado.IDRiesgo Then
                If m_ObjEdicionActiva.EsActivo Then
                    RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_ObjRiesgoActivo.EstadoEnum, m_Error
                    If m_Error <> "" Then
                        Err.Raise 1000
                    End If
                    Me.lblEstadoRiesgo.Visible = True
                Else
                    Me.lblEstadoRiesgo.Visible = False
                End If
                Set m_RiesgoUltimoSeleccionado = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo)
            End If
        End If
        
        
        AvanceCerrar
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    If TypeOf m_ObjSeleccionado Is PC Then
        Set m_PCSeleccionado = m_ObjSeleccionado
        Set m_ObjPCActivo = m_PCSeleccionado
        Set m_ObjRiesgoActivo = m_PCSeleccionado.riesgo
        Me.m_EsMitigacion = EnumSiNo.No
        
        If m_SinCargarFormulario <> EnumSiNo.Sí Then
            Me.FormDetalle.SourceObject = "FormRiesgosGestionPlanPrincipal"
        End If
        If m_RiesgoUltimoSeleccionado Is Nothing Then
            If m_ObjEdicionActiva.EsActivo Then
                RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_ObjRiesgoActivo.EstadoEnum, m_Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                Me.lblEstadoRiesgo.Visible = True
            Else
                Me.lblEstadoRiesgo.Visible = False
            End If
           
            Set m_RiesgoUltimoSeleccionado = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo)
        Else
            If m_ObjRiesgoActivo.IDRiesgo <> m_RiesgoUltimoSeleccionado.IDRiesgo Then
                If m_ObjEdicionActiva.EsActivo Then
                    RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_ObjRiesgoActivo.EstadoEnum, m_Error
                    If m_Error <> "" Then
                        Err.Raise 1000
                    End If
                    Me.lblEstadoRiesgo.Visible = True
                Else
                    Me.lblEstadoRiesgo.Visible = False
                End If
                Set m_RiesgoUltimoSeleccionado = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo)
            End If
        End If
        AvanceCerrar
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    If TypeOf m_ObjSeleccionado Is PMAccion Then
        Set m_PMAccionSeleccionado = m_ObjSeleccionado
        Set m_ObjPMAccionActiva = m_PMAccionSeleccionado
        Set m_ObjRiesgoActivo = m_PMAccionSeleccionado.Mitigacion.riesgo
        Me.m_EsMitigacion = EnumSiNo.Sí
        
        If m_SinCargarFormulario <> EnumSiNo.Sí Then
            Me.FormDetalle.SourceObject = "FormRiesgosGestionPlanAcciones"
        End If
        If m_RiesgoUltimoSeleccionado Is Nothing Then
            If m_ObjEdicionActiva.EsActivo Then
                RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_ObjRiesgoActivo.EstadoEnum, m_Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                Me.lblEstadoRiesgo.Visible = True
            Else
                Me.lblEstadoRiesgo.Visible = False
            End If
           
            Set m_RiesgoUltimoSeleccionado = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo)
        Else
            If m_ObjRiesgoActivo.IDRiesgo <> m_RiesgoUltimoSeleccionado.IDRiesgo Then
                If m_ObjEdicionActiva.EsActivo Then
                    RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_ObjRiesgoActivo.EstadoEnum, m_Error
                    If m_Error <> "" Then
                        Err.Raise 1000
                    End If
                    Me.lblEstadoRiesgo.Visible = True
                Else
                    Me.lblEstadoRiesgo.Visible = False
                End If
                Set m_RiesgoUltimoSeleccionado = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo)
            End If
        End If
        AvanceCerrar
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    If TypeOf m_ObjSeleccionado Is PCAccion Then
        Set m_PCAccionSeleccionado = m_ObjSeleccionado
        Set m_ObjPCAccionActiva = m_PCAccionSeleccionado
        Set m_ObjRiesgoActivo = m_PCAccionSeleccionado.Contingencia.riesgo
        Me.m_EsMitigacion = EnumSiNo.No
        If m_SinCargarFormulario <> EnumSiNo.Sí Then
            Me.FormDetalle.SourceObject = "FormRiesgosGestionPlanAcciones"
        End If
        If m_RiesgoUltimoSeleccionado Is Nothing Then
            If m_ObjEdicionActiva.EsActivo Then
                RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_ObjRiesgoActivo.EstadoEnum, m_Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                Me.lblEstadoRiesgo.Visible = True
            Else
                Me.lblEstadoRiesgo.Visible = False
            End If
           
            Set m_RiesgoUltimoSeleccionado = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo)
        Else
            If m_ObjRiesgoActivo.IDRiesgo <> m_RiesgoUltimoSeleccionado.IDRiesgo Then
                If m_ObjEdicionActiva.EsActivo Then
                    RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_ObjRiesgoActivo.EstadoEnum, m_Error
                    If m_Error <> "" Then
                        Err.Raise 1000
                    End If
                    Me.lblEstadoRiesgo.Visible = True
                Else
                    Me.lblEstadoRiesgo.Visible = False
                End If
                Set m_RiesgoUltimoSeleccionado = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo)
            End If
        End If
        AvanceCerrar
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Arbol_NodeClick se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub cmdSalir_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Public Sub ComandoActualizarContador_Click()
    
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
    End If
    Set m_ObjEdicionActiva = Constructor.getEdicion(m_ObjEdicionActiva.IDEdicion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CacheArbolRiesgosTx_RebuildEdicion m_ObjEdicionActiva, , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CargarArbol p_Refrescando:=EnumSiNo.Sí, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoActualizarContador_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
Public Function getNodoSeleccionado(Optional ByRef p_Error As String) As MSComctlLib.Node
    
    Dim m_Nodo As MSComctlLib.Node
    On Error GoTo errores
    
    For Each m_Nodo In m_Arbol.Nodes
        If m_Nodo.Selected = True Then
            Set getNodoSeleccionado = m_Nodo
            Exit Function
        End If
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getNodoSeleccionado ha producido el error num: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        
    End If
End Function
Public Function getNodoRiesgo(Optional p_Riesgo As riesgo, Optional ByRef p_Error As String) As MSComctlLib.Node
    
   
    On Error GoTo errores
    If p_Riesgo Is Nothing Then
        If m_ObjSeleccionado Is Nothing Then
            p_Error = "No hay nodo seleccionado"
            Err.Raise 1000
        End If
        If TypeOf m_ObjSeleccionado Is PM Then
            Set p_Riesgo = m_ObjSeleccionado.riesgo
        ElseIf TypeOf m_ObjSeleccionado Is PMAccion Then
            Set p_Riesgo = m_ObjSeleccionado.Mitigacion.riesgo
        ElseIf TypeOf m_ObjSeleccionado Is PC Then
            Set p_Riesgo = m_ObjSeleccionado.riesgo
        ElseIf TypeOf m_ObjSeleccionado Is PCAccion Then
            Set p_Riesgo = m_ObjSeleccionado.Contingencia.riesgo
        Else
            p_Error = "No hay nodo seleccionado"
            Err.Raise 1000
        End If
        
        If p_Riesgo Is Nothing Then
            p_Error = "No hay nodo seleccionado"
            Err.Raise 1000
        End If
    End If
    Set getNodoRiesgo = m_Arbol.Nodes("RIESGO|" & p_Riesgo.IDRiesgo)
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getNodoRiesgo ha producido el error num: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        
    End If
End Function
Public Function getNodoPlan(Optional p_Plan As Object, Optional ByRef p_Error As String) As MSComctlLib.Node
    
   
    On Error GoTo errores
    If p_Plan Is Nothing Then
        If m_ObjSeleccionado Is Nothing Then
            p_Error = "No hay nodo seleccionado"
            Err.Raise 1000
        End If
        If TypeOf m_ObjSeleccionado Is PM Then
            Set p_Plan = m_ObjSeleccionado
        ElseIf TypeOf m_ObjSeleccionado Is PMAccion Then
            Set p_Plan = m_ObjSeleccionado.Mitigacion
        ElseIf TypeOf m_ObjSeleccionado Is PC Then
            Set p_Plan = m_ObjSeleccionado
        ElseIf TypeOf m_ObjSeleccionado Is PCAccion Then
            Set p_Plan = m_ObjSeleccionado.Contingencia
        Else
            p_Error = "No hay nodo seleccionado"
            Err.Raise 1000
        End If
        
        If p_Plan Is Nothing Then
            p_Error = "No hay nodo seleccionado"
            Err.Raise 1000
        End If
    End If
    If TypeOf p_Plan Is PM Then
        Set getNodoPlan = m_Arbol.Nodes("PM|" & p_Plan.riesgo.IDRiesgo & "|" & p_Plan.IDMitigacion)
    Else
        Set getNodoPlan = m_Arbol.Nodes("PC|" & p_Plan.riesgo.IDRiesgo & "|" & p_Plan.IDContingencia)
    End If
    
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getNodoPlan ha producido el error num: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        
    End If
End Function


Private Sub ComandoAnexos_Click()
   
    Dim m_Argumento As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If TypeOf m_ObjSeleccionado Is Edicion Then
        m_Argumento = "Edición"
        
    ElseIf TypeOf m_ObjSeleccionado Is riesgo Then
        m_Argumento = "Riesgo"
        Set m_ObjRiesgoActivo = m_RiesgoSeleccionado
    Else
        m_Error = "El nodo seleccionado no es ni del tipo riesgo ni del tipo edición"
        Err.Raise 1000
    End If
    If FormularioAbierto("FormAnexos") Then
        DoCmd.Close acForm, "FormAnexos", acSaveNo
    End If
   
    
    DoCmd.OpenForm "FormAnexos", openArgs:=m_Argumento
    
    
    

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAnexos_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub


Private Sub ComandoAyuda_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    AbrirAyuda m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAyuda_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
  
Private Function getRiesgoSeleccioanado(Optional ByRef p_Error As String) As riesgo
    
    On Error GoTo errores
    If TypeOf m_ObjSeleccionado Is riesgo Then
        Set getRiesgoSeleccioanado = m_ObjSeleccionado
        Exit Function
    End If
    If TypeOf m_ObjSeleccionado Is Edicion Then
        Exit Function
    End If
    If TypeOf m_ObjSeleccionado Is PM Then
        Set getRiesgoSeleccioanado = m_ObjSeleccionado.riesgo
        Exit Function
    End If
    If TypeOf m_ObjSeleccionado Is PC Then
        Set getRiesgoSeleccioanado = m_ObjSeleccionado.riesgo
        Exit Function
    End If
    If TypeOf m_ObjSeleccionado Is PMAccion Then
        Set getRiesgoSeleccioanado = m_ObjSeleccionado.Mitigacion.riesgo
        Exit Function
    End If
    If TypeOf m_ObjSeleccionado Is PCAccion Then
        Set getRiesgoSeleccioanado = m_ObjSeleccionado.Contingencia.riesgo
        Exit Function
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getRiesgoSeleccioanado ha devuelto el error: " & _
                Err.Number & vbNewLine & "Detalle: " & Err.Description
        
    End If
End Function

Private Function getColRiesgosAAPlicar( _
                                        Optional p_VerRetirados As String = "No", _
                                        Optional ByRef p_Error As String) As Scripting.Dictionary
    
    On Error GoTo errores
    
    If Left$(UCase$(Trim$(Nz(p_VerRetirados, ""))), 1) = "S" Then
        Set getColRiesgosAAPlicar = m_ObjEdicionActiva.colRiesgos
        
    Else
        Set getColRiesgosAAPlicar = m_ObjEdicionActiva.colRiesgosActivos
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getColRiesgosAAPlicar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        
    End If
End Function

Private Sub ComandoEditarProyecto_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set m_ObjProyectoActivo = m_ObjEdicionActiva.Proyecto
    m_EsAlta = EnumSiNo.No
    If FormularioAbierto("FormGestionRiesgos") Then
        DoCmd.Close acForm, "FormGestionRiesgos", acSaveNo
    End If
    DoCmd.OpenForm "FormGestionRiesgos"
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEditarProyecto_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Public Sub ComboEdicion_BeforeUpdate(Cancel As Integer)
    Dim blnCamEdicionCambiada As Boolean
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    Me.ComandoRiesgoPriorizacion.Enabled = False
    Me.ComandoAnexos.Enabled = False
    Me.ComandoPublicacion.Enabled = False
    Me.ComandoEditarProyecto.Enabled = False
    If Nz(Me.ComboEdicion, "") = "" Then
        Set m_ObjEdicionActiva = Nothing
        blnCamEdicionCambiada = True
        
        CargarArbol p_Refrescando:=EnumSiNo.No, p_Error:=m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        ActualizarEtiquetaUltimoCambio
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    If m_ObjEdicionActiva Is Nothing Then
        If m_ObjProyectoActivo Is Nothing Then
            m_Error = "No se puede determinar ni la edición activa ni la gestión de riesgos activa"
            Err.Raise 1000
        End If
        Set m_ObjEdicionActiva = Constructor.getEdicionV2(m_ObjProyectoActivo.IDProyecto, Me.ComboEdicion, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        blnCamEdicionCambiada = True
    Else
        If Me.ComboEdicion <> m_ObjEdicionActiva.Edicion Then
            Set m_ObjEdicionActiva = Constructor.getEdicionV2(m_ObjEdicionActiva.IDProyecto, Me.ComboEdicion, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            blnCamEdicionCambiada = True
        End If
    End If
    
    Me.ComandoAnexos.Enabled = True
    Me.ComandoPublicacion.Enabled = True
    Me.ComandoEditarProyecto.Enabled = True
    If blnCamEdicionCambiada = True Then
        m_EsEdicionActiva = m_ObjEdicionActiva.EsActivo
        Me.Caption = m_ObjEntorno.TituloUsuarioConectado & " " & "riesgos de la edición " & m_ObjEdicionActiva.Edicion & " de " & m_ObjEdicionActiva.Proyecto.Proyecto
        CargarArbol p_Refrescando:=EnumSiNo.No, p_Error:=m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        ActualizarEtiquetaUltimoCambio
    End If
    If m_EsEdicionActiva = EnumSiNo.Sí Then
        Me.ComandoRiesgoPriorizacion.Enabled = True
    Else
        Me.ComandoRiesgoPriorizacion.Enabled = False
    End If
    Me.AllowEdits = True
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComboEdicion_BeforeUpdate se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub




Private Sub comboVerDescripcion_BeforeUpdate(Cancel As Integer)
    Dim m_VerDescripcion As String
    Dim m_Actualizado As EnumSiNo
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    

    m_Error = ""
    If CacheArbolRiesgos_UsarCache(m_Error) <> EnumSiNo.No Then
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Actualizado = CacheArbolRiesgos_ActualizarTitulosRiesgosEnTreeView(m_Arbol, m_ObjEdicionActiva, Nz(Me.ComboVerDescripcion, "Sí"), , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_Actualizado <> EnumSiNo.No Then
            GoTo salir_ok
        End If
    End If

    CargarArbol p_Refrescando:=EnumSiNo.No, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
salir_ok:
    Me.AllowEdits = True
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al comboVerDescripcion_BeforeUpdate se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComboVerRetirados_Change()
    On Error GoTo errores

    If Me.ComboVerRetirados.ListIndex < 0 Then
        Exit Sub
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False

    If CacheArbolRiesgos_UsarCache(m_Error) <> EnumSiNo.No Then
        If m_Error <> "" Then
            Err.Raise 1000
        End If

        Dim m_Actualizado As EnumSiNo
        m_Actualizado = CacheArbolRiesgos_AplicarFiltroRetiradosEnTreeView( _
            m_Arbol, m_ObjEdicionActiva, Nz(Me.ComboVerRetirados, "No"), Nz(Me.ComboVerDescripcion, "Sí"), m_ListImages, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_Actualizado <> EnumSiNo.No Then
            GoTo salir_ok
        End If
    End If

    CargarArbol p_Refrescando:=EnumSiNo.No, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
salir_ok:
    Me.AllowEdits = True
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComboVerRetirados_Change se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    Me.AllowEdits = True
End Sub

Private Sub Form_Load()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If IsNumeric(Nz(Me.openArgs(), "")) Then
        Set m_EdicionAlAbrir = Constructor.getEdicion(p_IDEdicion:=Me.openArgs, p_Error:=m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    ' ===============================
    '  Restaurar posición del separador
    ' ===============================
    
    MeWidth = Me.InsideWidth - 100
    MeWidthHalf = MeWidth \ 2
    ArbolLeft = Me.Arbol.Left
    SubFormLeft = Me.FormDetalle.Left
    
    Dim w As Long
    w = CLng(Val(GetSetting(C_APP, C_SECTION, C_KEY_ARbolW, "0")))
    
    ' Valor por defecto
    If w <= 0 Then
        w = SepW
    End If
    
    ' Clamp de seguridad
    If w < SepWMin Then w = SepWMin
    If w > MeWidthHalf Then w = MeWidthHalf
    
    ' Aplicar layout
    Me.Arbol.Width = w
    Me.Sep.Left = ArbolLeft + w
    Me.FormDetalle.Left = Me.Sep.Left + Me.Sep.Width
    Me.FormDetalle.Width = MeWidth - w

    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
    End If
    m_NodoEdicion.Selected = True
    Arbol_NodeClick m_NodoEdicion
    
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub Sep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button <> 1 Then Exit Sub

    MouseCursor IDC_SIZEWE
    m_sngLeftResizePos = X

    MeWidth = Me.InsideWidth - 100
    MeWidthHalf = MeWidth \ 2   ' división entera consistente

    ArbolLeft = Me.Arbol.Left
    SubFormLeft = Me.FormDetalle.Left
End Sub


Private Sub Sep_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    MouseCursor IDC_SIZEWE
    If Button <> 1 Then Exit Sub

    With Me.Sep
        Dim newLeft As Single
        Dim minLeft As Long
        Dim maxLeft As Long

        newLeft = .Left + X - m_sngLeftResizePos

        ' límites reales del separador
        minLeft = ArbolLeft + SepWMin
        maxLeft = ArbolLeft + MeWidthHalf

        If newLeft < minLeft Then newLeft = ArbolLeft + SepW
        If newLeft > maxLeft Then newLeft = maxLeft

        .Left = newLeft

        ' layout
        Me.Arbol.Width = .Left - ArbolLeft
        Me.FormDetalle.Left = .Left + .Width
        Me.FormDetalle.Width = (ArbolLeft + MeWidth) - Me.FormDetalle.Left
    End With
End Sub


Private Sub Sep_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button <> 1 Then Exit Sub

    ' Recalcular límites por si el usuario ha redimensionado el formulario
    MeWidth = Me.InsideWidth - 100
    MeWidthHalf = MeWidth \ 2

    With Me.Sep
        ' Clamp derecha
        If .Left > MeWidthHalf Then .Left = MeWidthHalf

        ' Clamp izquierda (mínimo)
        If .Left < SepWMin Then .Left = SepW

        ' Recalcular layout final
        Me.Arbol.Width = IIf((.Left - ArbolLeft) < SepWMin, SepWMin, (.Left - ArbolLeft))
        .Left = Me.Arbol.Left + Me.Arbol.Width

        Me.FormDetalle.Left = .Left + .Width
        Me.FormDetalle.Width = MeWidth - Me.Arbol.Width
    End With

    ' Persistir ancho del árbol (posición del separador)
    SaveSetting C_APP, C_SECTION, C_KEY_ARbolW, CStr(Me.Arbol.Width)
End Sub


Public Function SeleccionarNodo( _
                                p_Objeto As Object, _
                                Optional p_Error As String) As String
    
    Dim m_Nodo As MSComctlLib.Node
    Dim m_Key As String
    Dim m_IdRiesgo As String
    Dim m_IDPM As String
    Dim m_IdPMA As String
    Dim m_IDPC As String
    Dim m_IdPCA As String
    
    Dim m_Riesgo As riesgo
    Dim m_PM As PM
    Dim m_PC As PC
    Dim m_PMA As PMAccion
    Dim m_PCA As PCAccion
    
    On Error GoTo errores
    
    '   RIESGO1  RIESGO|IDRIESGO1
    '       PM1     PM|IDRIESGO1|IDPM1
    '           ACCIÓN1 PMACCION|IDPM1|IDACCION1
    '           ACCIÓN2 PMACCION|IDPM1|IDACCION2
    '       PM2     PM|IDRIESGO1|IDPM2
    '           ACCIÓN1 PMACCION|IDPM2|IDACCION1
    '   RIESGO2  RIESGO|IDRIESGO2
    
    If p_Objeto Is Nothing Then
        p_Error = "Se ha de indicar el Riesgo"
        Err.Raise 1000
    End If
    If TypeOf p_Objeto Is Edicion Then
        m_Key = "EDICION|" & m_ObjEdicionActiva.IDEdicion
        
    ElseIf TypeOf p_Objeto Is riesgo Then
        m_Key = "RIESGO|" & p_Objeto.IDRiesgo
    ElseIf TypeOf p_Objeto Is PM Then
        m_Key = "PM|" & p_Objeto.riesgo.IDRiesgo & "|" & p_Objeto.IDMitigacion
    ElseIf TypeOf p_Objeto Is PC Then
        m_Key = "PC|" & p_Objeto.riesgo.IDRiesgo & "|" & p_Objeto.IDContingencia
    ElseIf TypeOf p_Objeto Is PMAccion Then
        m_Key = "PMACCION|" & p_Objeto.PM.IDMitigacion & "|" & p_Objeto.IDAccionMitigacion
    ElseIf TypeOf p_Objeto Is PCAccion Then
        m_Key = "PCACCION|" & p_Objeto.PC.IDContingencia & "|" & p_Objeto.IDAccionContingencia
    Else
        p_Error = "No se conoce el nodo seleccionado"
        Err.Raise 1000
    End If
    On Error Resume Next
    Set m_Nodo = m_Arbol.Nodes(m_Key)
    If Err.Number <> 0 Then
        On Error GoTo errores
        p_Error = "No se encuentra el riesgo"
        Err.Raise 1000
    End If
    On Error GoTo errores
    m_Arbol.SelectedItem = m_Nodo
    Arbol_NodeClick m_Nodo
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarNodo ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function



Public Function CargarArbol( _
                                Optional p_Refrescando As EnumSiNo = EnumSiNo.No, _
                                Optional ByRef p_Error As String _
                                ) As String

    
    
    
'   EDICION1 EDICION|IDEDICION
'       RIESGO1  RIESGO|IDRIESGO1
'           PM1     PM|IDRIESGO1|IDPM1
'               ACCIÓN1 PMACCION|IDPM1|IDACCION1
'               ACCIÓN2 PMACCION|IDPM1|IDACCION2
'           PM2     PM|IDRIESGO1|IDPM2
'               ACCIÓN1 PMACCION|IDPM2|IDACCION1
'       RIESGO2  RIESGO|IDRIESGO2

    
    
    Dim m_Riesgo As riesgo
    Dim m_Key As String
    Dim m_NombreIcono As String
    Dim m_TituloNodo As String
    Dim m_IdRiesgo As Variant
    Dim m_VerRetirados As String
    Dim m_VerDescripcion As String
    Dim m_VerRetiradosCache As String
    Dim m_VerDescripcionCache As String
    Dim m_CargadoDesdeCache As EnumSiNo
    
    On Error GoTo errores
    
    With m_Arbol
        .Style = 7 'tvwTreelinesPlusMinusText
        '.Style = 6 'tvwTreelinesPlusMinusText
        .MousePointer = ccArrow
        .LineStyle = 1 'tvwRootLines
        .LabelEdit = tvwManual
        .ImageList = m_ListImages
        '.BorderStyle = ccFixedSingle
        .BorderStyle = 0
        .Appearance = ccFlat
        .OLEDragMode = ccOLEDragManual
        .OLEDropMode = ccOLEDropNone
        .Indentation = 300
        .PathSeparator = "\"
        .HideSelection = False
        .sorted = False
        .FullRowSelect = True
        .Enabled = True
        .CheckBoxes = False
        .SingleSel = True
        .Scroll = True
        .HotTracking = False
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Refresh
        .Nodes.Clear
        
    End With
    If m_ObjEdicionActiva Is Nothing Then
        Me.FormDetalle.SourceObject = ""
        Exit Function
    End If
    
    
    
    m_VerRetirados = Nz(Me.ComboVerRetirados, "No")
    If Left$(UCase$(Trim$(m_VerRetirados)), 1) = "S" Then
        m_VerRetirados = "Sí"
    Else
        m_VerRetirados = "No"
    End If
    m_VerDescripcion = Nz(Me.ComboVerDescripcion, "Sí")
    If Left$(UCase$(Trim$(m_VerDescripcion)), 1) = "S" Then
        m_VerDescripcion = "Sí"
    Else
        m_VerDescripcion = "No"
    End If
    m_VerRetiradosCache = m_VerRetirados
    m_VerDescripcionCache = m_VerDescripcion
    
    If CacheArbolRiesgos_UsarCache(p_Error) <> EnumSiNo.No Then
        If p_Error <> "" Then
            Err.Raise 1000
        End If

        m_CargadoDesdeCache = CacheArbolRiesgos_CargarEnTreeView(m_Arbol, m_ListImages, m_ObjEdicionActiva, m_VerRetiradosCache, m_VerDescripcionCache, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If

        If m_CargadoDesdeCache = EnumSiNo.No Or p_Refrescando = EnumSiNo.Sí Then
            CacheArbolRiesgosTx_RebuildEdicion m_ObjEdicionActiva, , p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            m_CargadoDesdeCache = CacheArbolRiesgos_CargarEnTreeView(m_Arbol, m_ListImages, m_ObjEdicionActiva, m_VerRetiradosCache, m_VerDescripcionCache, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If

        If m_CargadoDesdeCache <> EnumSiNo.No Then
            On Error Resume Next
            Set m_NodoEdicion = m_Arbol.Nodes("EDICION|" & m_ObjEdicionActiva.IDEdicion)
            On Error GoTo errores
            If Not m_NodoEdicion Is Nothing Then
                m_NodoEdicion.Expanded = True
            End If
            SeleccionarNodo m_ObjEdicionActiva
            m_Arbol.Refresh
            CargarArbol = "OK"
            Exit Function
        End If
    End If
    m_Key = "EDICION|" & m_ObjEdicionActiva.IDEdicion
    m_NombreIcono = FSO.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaCerrada32)
    
    m_TituloNodo = m_ObjEdicionActiva.NombreParaNodo
    Set m_NodoEdicion = m_Arbol.Nodes.Add(, , m_Key, m_TituloNodo, m_NombreIcono)
    m_NodoEdicion.ExpandedImage = FSO.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaAbierta32)
    
    Set m_ColRiesgosAplicados = getColRiesgosAAPlicar(m_VerRetirados, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColRiesgosAplicados Is Nothing Then
        Exit Function
    End If
    For Each m_IdRiesgo In m_ColRiesgosAplicados
        Set m_Riesgo = m_ColRiesgosAplicados(m_IdRiesgo)
       ' If m_Riesgo.CodigoRiesgo = "R003" Then Stop
        CargarArbolRiesgo p_NodoEdicion:=m_NodoEdicion, _
                                            p_Riesgo:=m_Riesgo, _
                                            p_borrarNodoSeleccionado:=EnumSiNo.Sí, _
                                            p_VerDescripcion:=m_VerDescripcion, _
                                            p_Refrescando:=p_Refrescando, p_Error:=p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_Riesgo = Nothing
    Next
    m_NodoEdicion.Expanded = True
    SeleccionarNodo m_ObjEdicionActiva
    m_Arbol.Refresh
    CargarArbol = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarArbol ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function EliminarArbolNodos( _
                                    p_NodoSuperior As MSComctlLib.Node, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim m_Nodo As MSComctlLib.Node
    
    
    On Error GoTo errores
    If p_NodoSuperior Is Nothing Then
        Exit Function
    End If
    Me.FormDetalle.SourceObject = ""
    m_Arbol.Nodes.Remove p_NodoSuperior.Key
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EliminarArbolNodos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function BorrarNodosHijos( _
                                    p_NodoPadre As MSComctlLib.Node, _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                    
    Dim i As Long
    Dim NItems As Long
    Dim m_Nodo As MSComctlLib.Node
    
    Dim m_ColNodosABorrar As Scripting.Dictionary
    Dim varItem As Variant
    
    On Error GoTo errores
    
    
    
    If p_NodoPadre Is Nothing Then
        Exit Function
    End If
    NItems = p_NodoPadre.Children
    If NItems = 0 Then
        Exit Function
    End If
    Set m_Nodo = p_NodoPadre.Child
    For i = 1 To NItems
        If m_ColNodosABorrar Is Nothing Then
            Set m_ColNodosABorrar = New Scripting.Dictionary
            m_ColNodosABorrar.CompareMode = TextCompare
        End If
        If Not m_ColNodosABorrar.Exists(m_Nodo.Key) Then
            m_ColNodosABorrar.Add m_Nodo.Key, m_Nodo.Key
        End If
        Set m_Nodo = m_Nodo.Next
       
    Next
    Set m_Nodo = Nothing
    If m_ColNodosABorrar Is Nothing Then
        Exit Function
    End If
    For Each varItem In m_ColNodosABorrar
        m_Arbol.Nodes.Remove varItem
    Next
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método BorrarNodosHijos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function CargarArbolRiesgo( _
                                    p_NodoEdicion As MSComctlLib.Node, _
                                    p_Riesgo As riesgo, _
                                    Optional ByRef p_borrarNodoSeleccionado As EnumSiNo = EnumSiNo.Sí, _
                                    Optional ByRef p_VerDescripcion As String = "Sí", _
                                    Optional p_Refrescando As EnumSiNo = EnumSiNo.No, _
                                    Optional ByRef p_Error As String _
                                    ) As String

    
    
    
    Dim m_Key As String
    Dim m_NodoInterno As Object
    
    Dim m_NombreIcono As String
    Dim m_TituloNodo As String
    Dim m_NodoRiesgo As MSComctlLib.Node
    Dim m_ColPMs As Scripting.Dictionary
    Dim m_PM As PM
    Dim m_IDPM As Variant
    Dim m_ColPCs As Scripting.Dictionary
    Dim m_PC As PC
    Dim m_IDPC As Variant
    
    
    
    On Error GoTo errores
    
    If p_NodoEdicion Is Nothing Then
        p_Error = "No se conoce el Nodo edición"
        Err.Raise 1000
    End If
    If p_borrarNodoSeleccionado = Empty Then
        p_borrarNodoSeleccionado = EnumSiNo.Sí
    End If
    'If p_Riesgo.CodigoRiesgo = "R004" Then Stop
'   EDICION1 EDICION|IDEDICION
'       RIESGO1  RIESGO|IDRIESGO1
'           PM1     PM|IDRIESGO1|IDPM1
'               ACCIÓN1 PMACCION|IDPM1|IDACCION1
'               ACCIÓN2 PMACCION|IDPM1|IDACCION2
'           PM2     PM|IDRIESGO1|IDPM2
'               ACCIÓN1 PMACCION|IDPM2|IDACCION1
'       RIESGO2  RIESGO|IDRIESGO2
    m_Key = "RIESGO|" & p_Riesgo.IDRiesgo
    If p_borrarNodoSeleccionado = EnumSiNo.Sí Then
        On Error Resume Next
        m_Arbol.Nodes.Remove m_Key
        On Error GoTo errores
        Err.Clear
    Else
        'hemos de colocarnos en el nodo del riesgo y borrar los hijos
        On Error Resume Next
        Set m_NodoRiesgo = m_Arbol.Nodes(m_Key)
        m_NodoRiesgo.Selected = True
        BorrarNodosHijos m_NodoRiesgo, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If p_Refrescando = EnumSiNo.Sí Then
        ' Ya no reseteamos propiedades dinámicas
    End If
    m_NombreIcono = getNombreIconoRiesgoArbolDeRiesgos(p_Riesgo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_VerDescripcion = "Sí" Then
        m_TituloNodo = p_Riesgo.NombreNodoDesc
    Else
        m_TituloNodo = p_Riesgo.NombreNodoEstado
    End If
   
    If p_borrarNodoSeleccionado = EnumSiNo.Sí Then
        Set m_NodoRiesgo = m_Arbol.Nodes.Add(p_NodoEdicion.Key, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
    Else
        m_NodoRiesgo.Text = m_TituloNodo
        
        m_NodoRiesgo.Image = m_NombreIcono
    End If
    Dim m_Color As String
    
    m_Color = p_Riesgo.ColorIconoCalculado
    If m_Color = "Negro" Then
        m_NodoRiesgo.ForeColor = vbBlack
    ElseIf m_Color = "Rojo" Then
        m_NodoRiesgo.ForeColor = vbRed
    End If
    
    '-------------
    ' PMS
    '-----------------
    Set m_ColPMs = p_Riesgo.ColPMs
    If Not m_ColPMs Is Nothing Then
        For Each m_IDPM In m_ColPMs
            Set m_PM = m_ColPMs(m_IDPM)
            CargarArbolPM P_NodoRiesgo:=m_NodoRiesgo, p_Riesgo:=p_Riesgo, _
                        p_PM:=m_PM, p_borrarNodoSeleccionado:=EnumSiNo.Sí, _
                        p_Refrescando:=p_Refrescando, p_Error:=p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
    End If
    '-------------
    ' PCS
    '-----------------
    Set m_ColPCs = p_Riesgo.ColPCs
    If Not m_ColPCs Is Nothing Then
        For Each m_IDPC In m_ColPCs
            Set m_PC = m_ColPCs(m_IDPC)
            'If p_Riesgo.CodigoRiesgo = "R016" Then Stop
            CargarArbolPC P_NodoRiesgo:=m_NodoRiesgo, p_Riesgo:=p_Riesgo, _
                        p_PC:=m_PC, p_borrarNodoSeleccionado:=EnumSiNo.Sí, _
                        p_Refrescando:=p_Refrescando, p_Error:=p_Error
        Next
    End If
        
       
    p_NodoEdicion.Expanded = True
    m_Arbol.Refresh
    CargarArbolRiesgo = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarArbolRiesgo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function CargarArbolPM( _
                                P_NodoRiesgo As MSComctlLib.Node, _
                                p_Riesgo As riesgo, _
                                p_PM As PM, _
                                Optional ByRef p_borrarNodoSeleccionado As EnumSiNo = EnumSiNo.Sí, _
                                Optional p_Refrescando As EnumSiNo = EnumSiNo.No, _
                                Optional ByRef p_Error As String _
                                ) As String

    
    
    
        
    Dim m_NombreIcono As String
    Dim m_TituloNodo As String
    Dim m_Nodo As MSComctlLib.Node
    Dim m_Key As String
    Dim m_NodoPM As MSComctlLib.Node
   
    
        
    Dim m_ColPMAS As Scripting.Dictionary
    Dim m_PMA As PMAccion
    Dim m_IdPMA As Variant
    
    
    '   EDICION1 EDICION|IDEDICION
    '       RIESGO1  RIESGO|IDRIESGO1
    '           PM1     PM|IDRIESGO1|IDPM1
    '               ACCIÓN1 PMACCION|IDPM1|IDACCION1
    '               ACCIÓN2 PMACCION|IDPM1|IDACCION2
    '           PM2     PM|IDRIESGO1|IDPM2
    '               ACCIÓN1 PMACCION|IDPM2|IDACCION1
    '       RIESGO2  RIESGO|IDRIESGO2
     
    
    
    On Error GoTo errores
    If P_NodoRiesgo Is Nothing Then
        p_Error = "El nodo está vacío"
        Err.Raise 1000
    End If
    If p_Riesgo Is Nothing Then
        p_Error = "El riesgo está vacío"
        Err.Raise 1000
    End If
    If p_borrarNodoSeleccionado = Empty Then
        p_borrarNodoSeleccionado = EnumSiNo.Sí
    End If
    
   
   
    On Error Resume Next
    If Err.Number <> 0 Then
        On Error GoTo errores
        m_Error = "No se ha podido obtener el nodo del riesgo"
        Err.Raise 1000
    End If
    On Error GoTo errores
    m_TituloNodo = p_PM.CodMitigacion
    
    If p_Refrescando = EnumSiNo.Sí Then
        ' Ya no reseteamos propiedades dinámicas
    End If
    
    m_NombreIcono = p_PM.NombreIcono

    
    m_Key = "PM|" & p_Riesgo.IDRiesgo & "|" & p_PM.IDMitigacion
    If p_borrarNodoSeleccionado = EnumSiNo.Sí Then
        Set m_NodoPM = m_Arbol.Nodes.Add(P_NodoRiesgo.Key, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
        Set p_PM.colAcciones = Nothing
    Else
        m_NodoPM.Text = m_TituloNodo
        m_NodoPM.Image = m_NombreIcono
    End If

    Set m_ColPMAS = p_PM.colAcciones
    If Not m_ColPMAS Is Nothing Then
        For Each m_IdPMA In m_ColPMAS
            Set m_PMA = m_ColPMAS(m_IdPMA)
            m_TituloNodo = m_PMA.CodAccion
            
            m_NombreIcono = m_PMA.NombreIcono
           
            m_Key = "PMACCION|" & p_PM.IDMitigacion & "|" & m_PMA.IDAccionMitigacion
            Set m_Nodo = m_Arbol.Nodes.Add(m_NodoPM.Key, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
            If m_NombreIcono = "Item_Completo_32x32.ico" Then
                m_Nodo.ForeColor = vbBlack
            Else
                m_Nodo.ForeColor = vbRed
            End If
            'm_Nodo.ForeColor = vbBlack
            Set m_PMA = Nothing
        Next
    End If
    m_NodoPM.Expanded = True
    m_Arbol.Refresh
    CargarArbolPM = "ok"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarArbolPM ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function CargarArbolAccion( _
                                P_NodoPlan As MSComctlLib.Node, _
                                p_Accion As Object, _
                                Optional ByRef p_Error As String _
                                ) As String

    
    
    
        
    Dim m_NombreIcono As String
    Dim m_TituloNodo As String
    Dim m_Nodo As MSComctlLib.Node
    Dim m_Key As String
    
    
    '               ACCIÓN1 PMACCION|IDPM1|IDACCION1
     
    On Error GoTo errores
    If P_NodoPlan Is Nothing Then
        p_Error = "EL nodo está vacío"
        Err.Raise 1000
    End If
    If p_Accion Is Nothing Then
        p_Error = "La acción está vacía"
        Err.Raise 1000
    End If
    m_TituloNodo = p_Accion.CodAccion
    m_NombreIcono = getNombreIconoAccionArbolDeRiesgos(p_Accion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    If TypeOf p_Accion Is PMAccion Then
        m_Key = "PMACCION|" & p_Accion.Mitigacion.IDMitigacion & "|" & p_Accion.IDAccionMitigacion
    ElseIf TypeOf p_Accion Is PCAccion Then
        m_Key = "PCACCION|" & p_Accion.Contingencia.IDContingencia & "|" & p_Accion.IDAccionContingencia
    Else
        p_Error = "No es acción de contingencia ni de mitigación"
        Err.Raise 1000
    End If
    Set m_Nodo = m_Arbol.Nodes.Add(P_NodoPlan.Key, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
           
        
    
    
    m_Arbol.Refresh
    CargarArbolAccion = "ok"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarArbolAccion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function CargarArbolPC( _
                                P_NodoRiesgo As MSComctlLib.Node, _
                                p_Riesgo As riesgo, _
                                p_PC As PC, _
                                Optional ByRef p_borrarNodoSeleccionado As EnumSiNo = EnumSiNo.Sí, _
                                Optional p_Refrescando As EnumSiNo = EnumSiNo.No, _
                                Optional ByRef p_Error As String _
                                ) As String

    
    
    
        
    Dim m_NombreIcono As String
    Dim m_TituloNodo As String
    Dim m_Nodo As MSComctlLib.Node
    Dim m_Key As String
    Dim m_NodoPC As MSComctlLib.Node
   
    
        
    Dim m_ColPCAS As Scripting.Dictionary
    Dim m_PCA As PCAccion
    Dim m_IdPCA As Variant
    
    
    '   EDICION1 EDICION|IDEDICION
    '       RIESGO1  RIESGO|IDRIESGO1
    '           PC1     PC|IDRIESGO1|IDPC1
    '               ACCIÓN1 PCACCION|IDPC1|IDACCION1
    '               ACCIÓN2 PCACCION|IDPC1|IDACCION2
    '           PC2     PC|IDRIESGO1|IDPC2
    '               ACCIÓN1 PCACCION|IDPC2|IDACCION1
    '       RIESGO2  RIESGO|IDRIESGO2
     
    
    
    On Error GoTo errores
    If P_NodoRiesgo Is Nothing Then
        p_Error = "El nodo está vacío"
        Err.Raise 1000
    End If
    If p_Riesgo Is Nothing Then
        p_Error = "El riesgo está vacío"
        Err.Raise 1000
    End If
    If p_borrarNodoSeleccionado = Empty Then
        p_borrarNodoSeleccionado = EnumSiNo.Sí
    End If
    'If p_Riesgo.CodigoRiesgo = "R009" Then Stop
   
    
    On Error Resume Next
    If Err.Number <> 0 Then
        On Error GoTo errores
        m_Error = "No se ha podido obtener el nodo del riesgo"
        Err.Raise 1000
    End If
    On Error GoTo errores
    m_Key = "PC|" & p_PC.riesgo.IDRiesgo & "|" & p_PC.IDContingencia
    
    If p_borrarNodoSeleccionado = EnumSiNo.Sí Then
        On Error Resume Next
        m_Arbol.Nodes.Remove m_Key
        On Error GoTo errores
        Err.Clear
    Else
        'hemos de colocarnos en el nodo del riesgo y borrar los hijos
        On Error Resume Next
        Set m_NodoPC = m_Arbol.Nodes(m_Key)
        m_NodoPC.Selected = True
        BorrarNodosHijos m_NodoPC, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_TituloNodo = p_PC.CodContingencia
    If p_Refrescando = EnumSiNo.Sí Then
        ' Ya no reseteamos propiedades dinámicas
    End If
   
    m_NombreIcono = p_PC.NombreIcono
    
    m_Key = "PC|" & p_Riesgo.IDRiesgo & "|" & p_PC.IDContingencia
    If p_borrarNodoSeleccionado = EnumSiNo.Sí Then
        Set m_NodoPC = m_Arbol.Nodes.Add(P_NodoRiesgo.Key, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
    Else
        m_NodoPC.Text = m_TituloNodo
        m_NodoPC.Image = m_NombreIcono
    End If

    Set m_ColPCAS = p_PC.colAcciones
    If Not m_ColPCAS Is Nothing Then
        For Each m_IdPCA In m_ColPCAS
            Set m_PCA = m_ColPCAS(m_IdPCA)
            m_TituloNodo = m_PCA.CodAccion
            
            m_NombreIcono = m_PCA.NombreIcono
            
            m_Key = "PCACCION|" & p_PC.IDContingencia & "|" & m_PCA.IDAccionContingencia
            Set m_Nodo = m_Arbol.Nodes.Add(m_NodoPC.Key, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
            
            Set m_PCA = Nothing
        Next
    End If
    m_NodoPC.Expanded = True
    m_Arbol.Refresh
    CargarArbolPC = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarArbolPC ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


