VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Proyecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' Identificadores
Public IDProyecto As Long
Public IDExpediente As Long

' Datos Principales
Public Proyecto As String
Public Juridica As String
Public NombreProyecto As String
Public Cliente As String
Public CodigoDocumento As String

' Fechas
Public FechaPrevistaCierre As Variant
Public FechaCierre As Variant
Public FechaRegistroInicial As Variant
Public FechaFirmaContrato As Variant
Public FechaMaxProximaPublicacion As Variant

' Responsables
Public Elaborado As String
Public Revisado As String
Public Aprobado As String
Public NombreUsuarioCalidad As String
Public CorreoRAC As String

' Flags y Configuración
Public ParaInformeAvisos As String ' "Sí"/"No"
Public EnUTE As String
Public RequiereRiesgoDeBiblioteca As String
Public RiesgosDeLaOferta As String
Public RiesgosDelSubContratista As String
Public Ordinal As Long
Public CadenaNombreAutorizados As String
Public NombreParaNodo As String

' Errores
Public Error As String

Public Function MotivoNoOK(Optional ByRef p_Error As String) As String
    ' Validación de Reglas de Negocio (Dominio Puro)
    If Me.IDExpediente = 0 Then
        MotivoNoOK = "Falta el ID del Expediente asociado."
        Exit Function
    End If
    
    If Trim(Me.NombreProyecto) = "" Then
        MotivoNoOK = "El nombre del proyecto no puede estar vacío."
        Exit Function
    End If
    
    ' Nota: Aquí se deben portar el resto de validaciones del legacy
    ' que no dependan de consultas a la base de datos.
End Function
