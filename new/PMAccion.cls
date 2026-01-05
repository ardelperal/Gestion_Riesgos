VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PMAccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' Identificadores
Public IDAccionMitigacion As Long
Public IDMitigacion As Long

' Datos de la Acción
Public CodAccion As String
Public Accion As String
Public ResponsableAccion As String
Public Estado As String ' "Planificada", "Finalizada", etc.
Public EsUltimaAccion As String ' "Sí"/"No"

' Fechas
Public FechaInicio As Variant
Public FechaFinPrevista As Variant
Public FechaFinReal As Variant

' Errores
Public Error As String
