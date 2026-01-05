VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RiesgoServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function RegistrarParaNoRetipificacion( _
                                                p_Id As String, _
                                                p_Descripcion As String, _
                                                p_CausaRaiz As String, _
                                                p_CodigoNoRetipificacion As String, _
                                                p_Fecha As String, _
                                                Optional ByRef p_Error As String _
                                                ) As String
                        
    
    Dim m_SQL As String
    
    On Error GoTo errores
    If p_Id = "" Then
        p_Error = "Falta el IDRiesgo"
        Err.Raise 1000
    End If
    If p_Descripcion = "" Then
        p_Error = "Falta la Descripción"
        Err.Raise 1000
    End If
    If p_CausaRaiz = "" Then
        p_Error = "Falta la causa raíz"
        Err.Raise 1000
    End If
    m_SQL = "UPDATE TbRiesgos " & _
            "SET TbRiesgos.RiesgoPendienteRetipificacion = 'No', " & _
            "TbRiesgos.Descripcion = '" & p_Descripcion & "', " & _
            "TbRiesgos.CausaRaiz = '" & p_CausaRaiz & "', " & _
            "TbRiesgos.FechaRiesgoRetipificado =#" & Format(p_Fecha, "mm/dd/yyyy") & "#, " & _
            "TbRiesgos.CodRiesgoBiblioteca = '" & p_CodigoNoRetipificacion & "' " & _
            "WHERE TbRiesgos.IDRiesgo=" & p_Id & ";"
    CacheArbolRiesgosTx_EjecutarSqlYActualizarRiesgoId CLng(p_Id), m_SQL, , , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo RegistrarParaNoRetipificacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
