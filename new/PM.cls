VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' Identificadores (Long)
Public IDMitigacion As Long
Public IDRiesgo As Long

' Datos del Plan
Public CodMitigacion As String
Public DisparadorDelPlan As String
Public Estado As String ' "Detectado", "Activo", etc.

' Fechas (Variant)
Public FechaDeActivacion As Variant
Public FechaDesactivacion As Variant

' Errores
Public Error As String
