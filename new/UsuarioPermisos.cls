VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsuarioPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Public CorreoUsuario As String
Public IDAplicacion As Integer


Option Compare Database
Option Explicit

Public IDUsuario As Long
Public IDAplicacion As Integer

' Permisos Específicos (Ajustar según campos reales de la tabla de permisos)
Public EsUsuarioAdministrador As String
Public EsUsuarioCalidad As String
Public EsUsuarioEconomia As String
Public EsUsuarioSecretaria As String
Public EsUsuarioTecnico As String
Public EsUsuarioCalidadAvisos As String
' Errores
Public Error As String



' Errores
Public Error As String
