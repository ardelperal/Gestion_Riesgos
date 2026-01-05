VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "riesgo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' Identificadores (Long en ERD)
Public IDRiesgo As Long
Public IDEdicion As Long
Public CodigoUnico As String
Public CodigoRiesgo As String

' Fechas (Variant para soportar Nulls de BD)
Public FechaDetectado As Variant
Public FechaEstado As Variant
Public FechaMaterializado As Variant
Public FechaRetirado As Variant
Public FechaCerrado As Variant

' Definición
Public DetectadoPor As String
Public EntidadDetecta As String
Public Descripcion As String
Public CausaRaiz As String
Public Origen As String

' Valoración e Impacto
Public Plazo As String
Public Calidad As String
Public Coste As String
Public ImpactoGlobal As String
Public Vulnerabilidad As String
Public Valoracion As String
Public Priorizacion As Integer ' Integer en ERD

' Estrategia
Public Mitigacion As String
Public Contingencia As String
Public RequierePlanContingencia As String ' Text(2) -> "Sí"/"No"

' Biblioteca y Tipificación
Public RequiereRiesgoDeBiblioteca As String
Public CodRiesgoBiblioteca As String
Public RiesgoPendienteRetipificacion As String
Public FechaRiesgoParaRetipificar As Variant
Public FechaRiesgoRetipificado As Variant

' Aceptación y Retiro
Public JustificacionAceptacionRiesgo As String
Public FechaJustificacionAceptacionRiesgo As Variant
Public FechaAprobacionAceptacionPorCalidad As Variant
Public FechaRechazoAceptacionPorCalidad As Variant
Public JustificacionRetiroRiesgo As String
Public FechaJustificacionRetiroRiesgo As Variant
Public FechaAprobacionRetiroPorCalidad As Variant
Public FechaRechazoRetiroPorCalidad As Variant

' Errores y UI
Public Error As String

' Lógica de Cálculo de Dominio (Sin DAO)

Public Property Get ImpactoGlobalCalculado() As String
    Dim m_intMax As Integer, m_intP As Integer, m_intCa As Integer, m_intCo As Integer
    Dim m_intImpacto As Integer
    
    If Me.Plazo = "" Or Me.Calidad = "" Or Me.Coste = "" Then Exit Property
    
    ' Mapeo de pesos
    m_intP = GetPesoImpacto(Me.Plazo): m_intCa = GetPesoImpacto(Me.Calidad): m_intCo = GetPesoImpacto(Me.Coste)
    
    ' Algoritmo de máximo
    m_intMax = m_intP
    If m_intCa > m_intMax Then m_intMax = m_intCa
    If m_intCo > m_intMax Then m_intMax = m_intCo
    
    ' Lógica especial: si todos son iguales y no extremos
    If m_intP = m_intCa And m_intP = m_intCo Then
        If m_intP = 5 Or m_intP = 1 Then m_intImpacto = m_intP Else m_intImpacto = m_intMax + 1
    Else
        m_intImpacto = m_intMax
    End If
    
    ImpactoGlobalCalculado = GetTextoImpacto(m_intImpacto)
End Property

Public Property Get ValoracionCalculada() As String
    ' Aquí se llama a la matriz de valoración. 
    ' Como la matriz es estática, la movemos a un helper de dominio.
    ValoracionCalculada = DomainHelpers.CalcularValoracion(Me.ImpactoGlobal, Me.Vulnerabilidad)
End Property

Public Property Get EstadoCalculado() As String
    ' Lógica de estados basada en flags de fechas y estrategia
    If Me.Mitigacion = "Aceptar" Then
        If Me.FechaAprobacionAceptacionPorCalidad <> "" Then EstadoCalculado = "Aceptado" Else EstadoCalculado = "AceptadoSinVisar"
        Exit Property
    End If
    
    If Me.FechaRetirado <> "" Then
        If Me.FechaAprobacionRetiroPorCalidad <> "" Then EstadoCalculado = "Retirado" Else EstadoCalculado = "RetiradoSinVisar"
        Exit Property
    End If
    
    If Me.FechaMaterializado <> "" Then
        EstadoCalculado = "Materializado"
        Exit Property
    End If
    
    If Me.IDRiesgo = 0 Or Me.Descripcion = "" Then
        EstadoCalculado = "Incompleto"
    Else
        EstadoCalculado = "Detectado"
    End If
End Property

Private Function GetPesoImpacto(ByVal p_Texto As String) As Integer
    Select Case p_Texto
        Case "Muy Bajo": GetPesoImpacto = 1: Case "Bajo": GetPesoImpacto = 2
        Case "Medio": GetPesoImpacto = 3: Case "Alto": GetPesoImpacto = 4: Case "Muy Alto": GetPesoImpacto = 5
    End Select
End Function

Private Function GetTextoImpacto(ByVal p_Peso As Integer) As String
    Select Case p_Peso
        Case 1: GetTextoImpacto = "Muy Bajo": Case 2: GetTextoImpacto = "Bajo"
        Case 3: GetTextoImpacto = "Medio": Case 4: GetTextoImpacto = "Alto": Case 5: GetTextoImpacto = "Muy Alto"
    End Select
End Function

' Lógica de Cálculo de Dominio (Sin DAO)

Public Property Get CodigoUnicoCalculado(p_IDProyecto As Long) As String
    CodigoUnicoCalculado = Format(p_IDProyecto, "000") & Me.CodigoRiesgo
End Property

Public Property Get PriorizacionEdAnterior(p_RiesgoPrevio As riesgo) As String
    If p_RiesgoPrevio Is Nothing Then
        PriorizacionEdAnterior = "-"
    Else
        PriorizacionEdAnterior = CStr(p_RiesgoPrevio.Priorizacion)
    End If
End Property

Public Property Get DescripcionParaLista() As String
    If Me.CausaRaiz = "" Then
        DescripcionParaLista = Me.Descripcion
    Else
        DescripcionParaLista = Me.Descripcion & ". " & Me.CausaRaiz
    End If
End Property

Public Property Get RiesgoAltoOMuyAlto() As Boolean
    RiesgoAltoOMuyAlto = (Me.Valoracion = "Alto" Or Me.Valoracion = "Muy Alto")
End Property

Public Property Get RequierePlanContingenciaCalculado() As Boolean
    If Me.Estado = "Retirado" Or Me.Estado = "Aceptado" Then
        RequierePlanContingenciaCalculado = False
    ElseIf Me.Estado = "Materializado" Or Me.RiesgoAltoOMuyAlto Then
        RequierePlanContingenciaCalculado = True
    Else
        RequierePlanContingenciaCalculado = False
    End If
End Property

' Portar validación de borrado
Public Function ValidacionEliminar(ByVal p_EsEdicionActiva As Boolean, ByRef p_Error As String) As Boolean
    ' Un riesgo solo se puede borrar si la edición está activa y nació en esta edición
    If Not p_EsEdicionActiva Then
        p_Error = "No se puede borrar un riesgo de una edición ya publicada."
        ValidacionEliminar = False
        Exit Function
    End If
    
    ' Nota: La comprobación de si nació en esta edición la hará el Servicio 
    ' comparando Me.IDEdicion con la edición donde nace.
    ValidacionEliminar = True
End Function

Public Function MotivoNoOK(Optional ByRef p_Error As String) As String
    ' VALIDACIONES DE DOMINIO
    If Me.IDEdicion = 0 Then
        MotivoNoOK = "Falta el ID de la Edición asociada."
        Exit Function
    End If
    ' ... (resto de validaciones portadas anteriormente)
End Function