VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Entorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' Variables de Entorno (POCO)
Public DatosEnLocal As String ' "Sí"/"No"
Public EnDesarrollo As String
Public EnPruebas As String
Public IDAplicacion As String

' Rutas
Public RutaAplicacionesRemotas As String
Public RutaAplicacionRemota As String
Public RutaAplicacionesLocal As String
Public RutaAplicacionLocal As String

' Configuración de Negocio
Public JPMesesAvisoEntreEdiciones As Integer
Public JPDiasPreviosParaElAviso As Integer
Public CalDiaInicialMesAviso As Integer
Public GenerarInformeTipo As String ' "WORD", "EXCEL", etc.

' Estado de Usuario
Public UsuarioConectadoEsAdministrador As String
Public UsuarioConectadoEsDeCalidad As String

' Errores
Public Error As String
