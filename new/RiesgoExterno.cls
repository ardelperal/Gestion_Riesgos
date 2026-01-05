VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RiesgoExterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' Identificadores (Long en ERD)
Public IDRiesgoExt As Long
Public CodRiesgo As String
Public IDRiesgo As Long ' Vínculo a TbRiesgos
Public IDEdicion As Long

' Metadatos
Public Origen As String
Public FechaDetectado As Variant ' Variant para Fechas
Public FechaAltaRegistro As Variant
Public UsuarioRegistra As String

' Definición
Public Descripcion As String
Public CausaRaiz As String

' Lógica de Integración
Public Trasladar As String ' "Sí" o "No"
Public MotivoNoIntegrado As String
Public FechaMotivo As Variant

' Datos Específicos por Origen
Public Suministrador As String
Public Pedido As String
Public ProveedorPedido As String
Public CausaRiesgoPedido As String

' Tipificación
Public RequiereRiesgoDeBiblioteca As String
Public CodRiesgoBiblioteca As String
Public RiesgoPendienteRetipificacion As String

' Errores
Public Error As String

Public Function MotivoNoOK(Optional ByVal p_ObjAlInicio As RiesgoExterno, Optional ByRef p_Error As String) As String
    ' Validación
End Function