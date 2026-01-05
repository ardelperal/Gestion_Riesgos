Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : AppConstants
' Purpose   : Centraliza las constantes de sistema para evitar problemas de codificación.
'---------------------------------------------------------------------------------------

Public Const DB_YES As String = "S"  ' Usamos "S" en lugar de "Sí" para evitar la tilde
Public Const DB_NO As String = "N"

Public Enum eTriState
    tsNo = 0
    tsYes = 1
    tsUndefined = 2
End Enum
