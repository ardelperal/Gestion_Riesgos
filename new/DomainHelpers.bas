Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : DomainHelpers
' Purpose   : Lógica de negocio pura compartida por múltiples entidades de dominio.
'             Sin dependencias de base de datos.
'---------------------------------------------------------------------------------------

Public Function CalcularValoracion(ByVal p_Impacto As String, ByVal p_Vulnerabilidad As String) As String
    ' Matriz de Riesgo Estándar (Impacto x Vulnerabilidad)
    
    If p_Impacto = "" Or p_Vulnerabilidad = "" Then Exit Function
    
    Select Case p_Impacto
        Case "Muy Alto"
            Select Case p_Vulnerabilidad
                Case "Muy Alta": CalcularValoracion = "Muy Alto"
                Case "Alta": CalcularValoracion = "Muy Alto"
                Case "Media": CalcularValoracion = "Alto"
                Case "Baja": CalcularValoracion = "Medio"
                Case "Muy Baja": CalcularValoracion = "Bajo"
            End Select
            
        Case "Alto"
            Select Case p_Vulnerabilidad
                Case "Muy Alta": CalcularValoracion = "Muy Alto"
                Case "Alta": CalcularValoracion = "Alto"
                Case "Media": CalcularValoracion = "Alto"
                Case "Baja": CalcularValoracion = "Medio"
                Case "Muy Baja": CalcularValoracion = "Bajo"
            End Select
            
        Case "Medio"
            Select Case p_Vulnerabilidad
                Case "Muy Alta": CalcularValoracion = "Alto"
                Case "Alta": CalcularValoracion = "Alto"
                Case "Media": CalcularValoracion = "Medio"
                Case "Baja": CalcularValoracion = "Bajo"
                Case "Muy Baja": CalcularValoracion = "Bajo"
            End Select
            
        Case "Bajo"
            Select Case p_Vulnerabilidad
                Case "Muy Alta": CalcularValoracion = "Medio"
                Case "Alta": CalcularValoracion = "Medio"
                Case "Media": CalcularValoracion = "Bajo"
                Case "Baja": CalcularValoracion = "Bajo"
                Case "Muy Baja": CalcularValoracion = "Muy Bajo"
            End Select
            
        Case "Muy Bajo"
            Select Case p_Vulnerabilidad
                Case "Muy Alta": CalcularValoracion = "Bajo"
                Case "Alta": CalcularValoracion = "Bajo"
                Case "Media": CalcularValoracion = "Bajo"
                Case "Baja": CalcularValoracion = "Muy Bajo"
                Case "Muy Baja": CalcularValoracion = "Muy Bajo"
            End Select
    End Select
End Function
