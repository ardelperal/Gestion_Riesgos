VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProyectoSuministrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public ID As String
Public IDProyecto As String
Public IDSuministrador As String
Public GestionCalidad As String

Private m_ObjSuministrador As Suministrador
Private m_ObjProyecto As Proyecto
Private m_sIDCalculado As String
Private m_eEsCalidadCalculado As EnumSiNo
Private m_objEdicionActivaSuministrador As EdicionSuministrador


Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get IDCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDCalculado = DameID("TbProyectosSuministradores", "ID", getdb(), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDCalculado = m_sIDCalculado
    
    Exit Property
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método ProyectoSuministrador.IDCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Suministrador() As Suministrador
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjSuministrador Is Nothing Then
        Set Suministrador = m_ObjSuministrador
        Exit Property
    End If
    
    Set m_ObjSuministrador = Constructor.getSuministrador(p_Id:=Me.IDSuministrador, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Suministrador = m_ObjSuministrador
    Exit Property
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método ProyectoSuministrador.Suministrador ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Proyecto() As Proyecto
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjProyecto Is Nothing Then
        Set Proyecto = m_ObjProyecto
        Exit Property
    End If
    
    Set m_ObjProyecto = Constructor.getProyecto(Me.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set Proyecto = m_ObjProyecto
    
    Exit Property
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método ProyectoSuministrador.Proyecto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EsCalidadCalculado() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.GestionCalidad <> "Sí" And Me.GestionCalidad <> "No" Then
        Exit Property
    End If
    If Me.GestionCalidad = "Sí" Then
        m_eEsCalidadCalculado = EnumSiNo.Sí
    Else
        m_eEsCalidadCalculado = EnumSiNo.No
    End If
    EsCalidadCalculado = m_eEsCalidadCalculado
    
    Exit Property
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método ProyuectoSuministrador.EsCalidadCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EdicionActivaSuministrador() As EdicionSuministrador
    
    Dim m_Proyecto As Proyecto
    Dim m_EdicionActiva As Edicion
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objEdicionActivaSuministrador Is Nothing Then
        Set EdicionActivaSuministrador = m_objEdicionActivaSuministrador
        Exit Property
    End If
    If Me.IDSuministrador = "" Then
        Exit Property
    End If
    Set m_Proyecto = Me.Proyecto
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Proyecto Is Nothing Then
        m_Error = "No se ha podido encontrar la gestión de riesgos"
        Err.Raise 1000
    End If
    Set m_EdicionActiva = m_Proyecto.EdicionActiva
    m_Error = m_Proyecto.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_EdicionActiva Is Nothing Then
        Set m_objEdicionActivaSuministrador = Constructor.getSuministradorEnEdicion(, m_EdicionActiva.IDEdicion, Me.IDSuministrador, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Set EdicionActivaSuministrador = m_objEdicionActivaSuministrador
     Exit Property
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método ProyuectoSuministrador.EdicionActivaSuministrador ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set EdicionActivaSuministrador(p_Valor As EdicionSuministrador)
    Set m_objEdicionActivaSuministrador = p_Valor
End Property





Public Function MotivoNoOK( _
                                Optional ByRef p_Error As String _
                            ) As String
    
    
    Dim m_Suministrador As Suministrador
    On Error GoTo errores
    
    If EsTecnico = EnumSiNo.Sí Then
        MotivoNoOK = "No está autorizado a realizar esta operación"
        Exit Function
    End If
    If Me.IDProyecto = "" Then
        MotivoNoOK = "No se conoce el IDProyecto"
        Exit Function
    End If
    If Me.IDSuministrador = "" Then
        MotivoNoOK = "El Suminstrador es obligatorio"
        Exit Function
    End If
    If Me.GestionCalidad <> "Sí" And Me.GestionCalidad <> "No" Then
        MotivoNoOK = "GestionCalidad ha de ser Sí o No"
        Exit Function
    End If
    Set m_Suministrador = Me.Suministrador
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Suministrador Is Nothing Then
        MotivoNoOK = "El Suminstrador no parece estar registrado"
        Exit Function
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MotivoNoOK ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Registrar( _
                            p_ProyectoSuministradorAlInicio As ProyectoSuministrador, _
                            Optional ByRef p_Error As String _
                            ) As String
                            
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    Dim blnIntegrarElSuministradorEnEdicionActiva As Boolean
    Dim blnQuitarElSuministradorEnEdicionActiva As Boolean
    Dim m_EdicionSuministrador As EdicionSuministrador
    Dim m_EdicionActiva As Edicion
    On Error GoTo errores
    
    Me.Error = ""
    
    m_MotivoNoOK = MotivoNoOK(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If p_ProyectoSuministradorAlInicio Is Nothing Then
        Me.ID = Me.IDCalculado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_SQL = "TbProyectosSuministradores"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbProyectosSuministradores " & _
                "WHERE ID=" & p_ProyectoSuministradorAlInicio.ID & ";"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If p_ProyectoSuministradorAlInicio Is Nothing Then
            .AddNew
                .Fields("ID") = Me.ID
                .Fields("IDProyecto") = Me.IDProyecto
                .Fields("IDSuministrador") = Me.IDSuministrador
                .Fields("GestionCalidad") = Me.GestionCalidad
            .Update
            If Me.GestionCalidad = "Sí" Then
                blnIntegrarElSuministradorEnEdicionActiva = True
            End If
        Else
            If .EOF Then
                p_Error = "NO existe ningún Proyecto suministrador con ese ID"
                Err.Raise 1000
            End If
            .Edit
                
                .Fields("IDSuministrador") = Me.IDSuministrador
                .Fields("GestionCalidad") = Me.GestionCalidad
            .Update
            If Me.GestionCalidad <> p_ProyectoSuministradorAlInicio.GestionCalidad Then
                If Me.GestionCalidad = "Sí" And p_ProyectoSuministradorAlInicio.GestionCalidad = "No" Then
                    blnIntegrarElSuministradorEnEdicionActiva = True
                End If
                If Me.GestionCalidad = "No" And p_ProyectoSuministradorAlInicio.GestionCalidad = "Sí" Then
                    blnQuitarElSuministradorEnEdicionActiva = True
                End If
            End If
        End If
        
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
   
    If blnIntegrarElSuministradorEnEdicionActiva = False And blnQuitarElSuministradorEnEdicionActiva = False Then
        Exit Function
    End If
    Set m_EdicionActiva = Me.Proyecto.EdicionActiva
    If m_EdicionActiva Is Nothing Then
        Exit Function
    End If
    If blnIntegrarElSuministradorEnEdicionActiva = True Then
        Set m_EdicionSuministrador = New EdicionSuministrador
        With m_EdicionSuministrador
            .IDEdicion = m_EdicionActiva.IDEdicion
            .IDSuministrador = Me.IDSuministrador
            .AltaPorCalidad p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
        End With
    
    End If
    If blnQuitarElSuministradorEnEdicionActiva = True Then
        Set m_EdicionSuministrador = Constructor.getSuministradoresEnEdicion(m_EdicionActiva.IDEdicion, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_EdicionSuministrador.EliminarPorCalidad p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Registrar ha devuelto el error: " & Err.Description
    End If
End Function



Public Function CambiarCalidad( _
                                p_Calidad As String, _
                                Optional ByRef p_Error As String _
                                ) As String
                            
    
    Dim m_SQL As String
    Dim m_EdicionActivaSuministrador As EdicionSuministrador
    Dim blnIntegrarElSuministradorEnEdicionActiva As Boolean
    Dim blnQuitarElSuministradorEnEdicionActiva As Boolean
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If p_Calidad <> "Sí" And p_Calidad <> "No" Then
        p_Error = "La gestión de la calidad ha de ser sí o no"
        Err.Raise 1000
    End If
    m_SQL = "UPDATE TbProyectosSuministradores " & _
            "SET GestionCalidad = '" & p_Calidad & "' " & _
            "WHERE ID=" & Me.ID & ";"
    getdb().Execute m_SQL
    Me.GestionCalidad = p_Calidad
    If p_Calidad = "Sí" Then
        Set m_EdicionActivaSuministrador = Me.EdicionActivaSuministrador
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_EdicionActivaSuministrador Is Nothing Then
            Set m_EdicionActivaSuministrador = New EdicionSuministrador
            With m_EdicionActivaSuministrador
                .IDEdicion = Me.Proyecto.EdicionActiva.IDEdicion
                .IDSuministrador = Me.IDSuministrador
                .AltaPorCalidad p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                Set Me.EdicionActivaSuministrador = m_EdicionActivaSuministrador
            End With
        End If
        Exit Function
    End If
    Set m_EdicionActivaSuministrador = Me.EdicionActivaSuministrador
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_EdicionActivaSuministrador Is Nothing Then
        m_EdicionActivaSuministrador.EliminarPorCalidad p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set Me.EdicionActivaSuministrador = Nothing
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CambiarCalidad ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Eliminar( _
                            Optional ByRef p_Error As String _
                            ) As String
                            
    Dim m_SQL As String
    Dim m_Col As Scripting.Dictionary
    Dim m_Id As Variant
    Dim m_EdicionSuministrador As EdicionSuministrador
    
    On Error GoTo errores
    
    Me.Error = ""

    Set m_EdicionSuministrador = Me.EdicionActivaSuministrador
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    m_SQL = "DELETE * " & _
        "FROM TbProyectosSuministradores " & _
        "WHERE ID=" & Me.ID & ";"
    getdb().Execute m_SQL
    
    If Not m_EdicionSuministrador Is Nothing Then
        m_EdicionSuministrador.EliminarPorCalidad p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Eliminar ha devuelto el error: " & Err.Description
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
        .Add "ID"
        .Add "IDProyecto"
        .Add "IDSuministrador"
        .Add "GestionCalidad"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método ProyectoSuministrador.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "ID" Then
        getPropiedad = Me.ID
        Exit Function
    End If
    If m_NombrePropiedad = "IDProyecto" Then
        getPropiedad = Me.IDProyecto
        Exit Function
    End If
    If m_NombrePropiedad = "IDSuministrador" Then
        getPropiedad = Me.IDSuministrador
        Exit Function
    End If
    If m_NombrePropiedad = "GestionCalidad" Then
        getPropiedad = Me.GestionCalidad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en ProyectoSuministrador.getPropiedad"
     Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ProyectoSuministrador.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "ID" Then
        Me.ID = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDProyecto" Then
        Me.IDProyecto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDSuministrador" Then
        Me.IDSuministrador = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "GestionCalidad" Then
        Me.GestionCalidad = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en ProyectoSuministrador.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ProyectoSuministrador.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

