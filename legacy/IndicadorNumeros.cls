VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IndicadorNumeros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_sIDentificados As String
Private m_sIDentificadosSemestre1 As String
Private m_sIDentificadosSemestre2 As String
Private m_sIDentificadosResumenAnual As String
Private m_sMaterializados As String
Private m_sMaterializadosSemestre1 As String
Private m_sMaterializadosSemestre2 As String
Private m_sMaterializadosResumenAnual As String

Public Property Get identificados() As String

    identificados = m_sIDentificados

End Property

Public Property Let identificados(ByVal sNewValue As String)

    m_sIDentificados = sNewValue

End Property

Public Property Get IDentificadosSemestre1() As String

    IDentificadosSemestre1 = m_sIDentificadosSemestre1

End Property

Public Property Let IDentificadosSemestre1(ByVal sNewValue As String)

    m_sIDentificadosSemestre1 = sNewValue

End Property

Public Property Get IDentificadosSemestre2() As String

    IDentificadosSemestre2 = m_sIDentificadosSemestre2

End Property

Public Property Let IDentificadosSemestre2(ByVal sNewValue As String)

    m_sIDentificadosSemestre2 = sNewValue

End Property

Public Property Get IDentificadosResumenAnual() As String

    IDentificadosResumenAnual = m_sIDentificadosResumenAnual

End Property

Public Property Let IDentificadosResumenAnual(ByVal sNewValue As String)

    m_sIDentificadosResumenAnual = sNewValue

End Property

Public Property Get materializados() As String

    materializados = m_sMaterializados

End Property

Public Property Let materializados(ByVal sNewValue As String)

    m_sMaterializados = sNewValue

End Property

Public Property Get MaterializadosSemestre1() As String

    MaterializadosSemestre1 = m_sMaterializadosSemestre1

End Property

Public Property Let MaterializadosSemestre1(ByVal sNewValue As String)

    m_sMaterializadosSemestre1 = sNewValue

End Property

Public Property Get MaterializadosSemestre2() As String

    MaterializadosSemestre2 = m_sMaterializadosSemestre2

End Property

Public Property Let MaterializadosSemestre2(ByVal sNewValue As String)

    m_sMaterializadosSemestre2 = sNewValue

End Property

Public Property Get MaterializadosResumenAnual() As String

    MaterializadosResumenAnual = m_sMaterializadosResumenAnual

End Property

Public Property Let MaterializadosResumenAnual(ByVal sNewValue As String)

    m_sMaterializadosResumenAnual = sNewValue

End Property

