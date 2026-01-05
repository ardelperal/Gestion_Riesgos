VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Expediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' Identificadores
Public IDExpediente As Long
Public IDExpedientePadre As Long

' Datos Principales
Public Nemotecnico As String
Public Titulo As String
Public CodProyecto As String
Public CodExp As String
Public CodExpLargo As String
Public CodigoActividad As String

' Fechas
Public FechaInicioContrato As Variant
Public FechaFinContrato As Variant
Public FechaFinGarantia As Variant
Public FechaFirmaContrato As Variant
Public FechaCreacion As Variant
Public FechaUltimoCambio As Variant

' Otros datos
Public Ordinal As Long
Public IDOrganoContratacion As Long
Public IDResponsableCalidad As String
Public Tipo As String
Public Estado As String
Public Ambito As String

' Errores
Public Error As String
