VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Suministrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' Identificadores
Public IDSuministrador As Long

' Datos
Public Nemotecnico As String
Public NombreSuministrador As String
Public CIF As String
Public Activo As String ' "S"/"N"

' Errores
Public Error As String
