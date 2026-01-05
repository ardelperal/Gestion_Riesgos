VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' Identificadores (Long)
Public IDContingencia As Long
Public IDRiesgo As Long

' Datos del Plan
Public CodContingencia As String
Public DisparadorDelPlan As String
Public Estado As String

' Fechas (Variant)
Public FechaDeActivacion As Variant
Public FechaDesactivacion As Variant

' Errores
Public Error As String
