VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' Datos Identificativos
Public ID As Long
Public CorreoUsuario As String
Public UsuarioRed As String
Public Nombre As String
Public Matricula As String

' Estado y Fechas
Public Activado As Boolean
Public FechaAlta As Variant
Public FechaBaja As Variant
Public FechaUltimaConexion As Variant

' Perfil Calculado (Service/Domain logic)
Public EsAdministrador As Boolean
Public EsCalidad As Boolean
Public EsTecnico As Boolean

' Permisos
Public Permisos As UsuarioPermisos

' Errores
Public Error As String
