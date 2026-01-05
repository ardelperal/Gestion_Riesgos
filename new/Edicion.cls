VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Edicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' Identificadores
Public IDEdicion As Long
Public IDProyecto As Long

' Datos de la Edición
Public Edicion As Integer
Public FechaEdicion As Variant
Public FechaPublicacion As Variant
Public Elaborado As String
Public Revisado As String
Public Aprobado As String

' Publicación e Informes
Public EntregadoAClienteORAC As String
Public Comentarios As String
Public PermitidoImprimirExcel As String
Public IDDocumentoAGEDO As Long
Public NombreArchivoInforme As String
Public FechaMaxProximaPublicacion As Variant
Public FechaPreparadaParaPublicar As Variant
Public UsuarioProponePublicar As String

' Calidad
Public PropuestaRechazadaPorCalidadFecha As Variant
Public PropuestaRechazadaPorCalidadMotivo As String
Public UsuarioCalidadRechazaPropuesta As String
Public NotasCalidadParaPublicar As String

' Metadatos
Public FechaUltimoCambio As Variant
Public UsuarioUltimoCambio As String

' Errores
Public Error As String
