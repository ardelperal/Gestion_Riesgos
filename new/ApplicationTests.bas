Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : ApplicationTests
' Purpose   : Suite de pruebas unitarias e integración para validar las nuevas capas.
'---------------------------------------------------------------------------------------

Public Sub RunAllTests()
    Debug.Print "--- INICIANDO SUITE DE TESTS ---"
    
    Test_Bootstrapper_EVE
    Test_UsuarioService_CargaPerfil
    Test_Riesgo_ValidacionDominio
    Test_RiesgoRepository_CRUD
    Test_Planes_Integracion
    
    Debug.Print "--- SUITE DE TESTS FINALIZADA ---"
End Sub

' 2. Test de Validación de Negocio (Dominio Puro)
Private Sub Test_Riesgo_ValidacionDominio()
    Dim m_Riesgo As riesgo
    Dim m_Motivo As String
    
    Set m_Riesgo = New riesgo
    
    ' Escenario 1: Falta descripción
    m_Riesgo.IDEdicion = 1
    m_Motivo = m_Riesgo.MotivoNoOK()
    If m_Motivo <> "" Then
        Debug.Print "[OK] Test_Riesgo_ValidacionDominio: Detectada falta de descripción."
    Else
        Debug.Print "[FAIL] Test_Riesgo_ValidacionDominio: No detectó falta de descripción."
    End If
    
    ' Escenario 2: Todo OK
    With m_Riesgo
        .Descripcion = "DESCRIPCION VALIDA"
        .Plazo = "Medio": .Calidad = "Bajo": .Coste = "Alto"
        .Vulnerabilidad = "Media"
        .Mitigacion = "Mitigar"
    End With
    m_Motivo = m_Riesgo.MotivoNoOK()
    If m_Motivo = "" Then
        Debug.Print "[OK] Test_Riesgo_ValidacionDominio: Validación exitosa con datos correctos."
    Else
        Debug.Print "[FAIL] Test_Riesgo_ValidacionDominio: Rechazó datos válidos por: " & m_Motivo
    End If
End Sub

' ... (procedimientos anteriores)

' 4. Test de Integración de Planes y Acciones
Private Sub Test_Planes_Integracion()
    Dim m_Riesgo As riesgo
    Dim m_PM As PM
    Dim m_Accion As PMAccion
    Dim m_Error As String
    
    On Error GoTo errores
    
    ' A. Crear un Riesgo base para el test
    Set m_Riesgo = New riesgo
    m_Riesgo.IDEdicion = 1
    m_Riesgo.Descripcion = "RIESGO PARA TEST DE PLANES"
    RiesgoRepository.Save m_Riesgo, , m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    Debug.Print "[INFO] Test_Planes_Integracion: Riesgo base creado (ID:" & m_Riesgo.IDRiesgo & ")"
    
    ' B. Crear un Plan de Mitigación (PM) vinculado al riesgo
    Set m_PM = New PM
    m_PM.IDRiesgo = m_Riesgo.IDRiesgo
    m_PM.CodMitigacion = "PM-TEST-01"
    m_PM.Estado = "Activo"
    
    PlanRepository.SavePM m_PM, , m_Error
    If m_Error <> "" Then
        Debug.Print "[FAIL] Test_Planes_Integracion (SavePM): " & m_Error
        Exit Sub
    End If
    Debug.Print "[OK] Test_Planes_Integracion: Plan de Mitigación persistido (ID:" & m_PM.IDMitigacion & ")"
    
    ' C. Crear una Acción dentro del Plan
    Set m_Accion = New PMAccion
    m_Accion.IDMitigacion = m_PM.IDMitigacion
    m_Accion.CodAccion = "A01"
    m_Accion.Accion = "ACCIÓN DE PRUEBA"
    m_Accion.Estado = "Planificada"
    
    PlanRepository.SavePMAccion m_Accion, , m_Error
    If m_Error <> "" Then
        Debug.Print "[FAIL] Test_Planes_Integracion (SavePMAccion): " & m_Error
        Exit Sub
    End If
    Debug.Print "[OK] Test_Planes_Integracion: Acción de Plan persistida correctamente."
    
    ' D. Limpieza (Cascade Delete en BD debería borrar PM y Accion al borrar Riesgo)
    RiesgoRepository.Delete m_Riesgo.IDRiesgo, , m_Error
    Debug.Print "[OK] Test_Planes_Integracion: Limpieza realizada."
    
    Exit Sub
errores:
    Debug.Print "[ERROR] Test_Planes_Integracion: " & Err.Description
End Sub
    Dim m_Error As String
    Dim m_Resultado As String
    
    On Error GoTo errores
    m_Resultado = ApplicationBootstrapper.EVE("", m_Error)
    
    If m_Resultado = "OK" And Not m_ObjEntorno Is Nothing Then
        Debug.Print "[OK] Test_Bootstrapper_EVE: Entorno inicializado."
        Debug.Print "     Ruta Local detectada: " & m_ObjEntorno.RutaAplicacionesLocal
    Else
        Debug.Print "[FAIL] Test_Bootstrapper_EVE: " & m_Error
    End If
    Exit Sub
errores:
    Debug.Print "[ERROR] Test_Bootstrapper_EVE: " & Err.Description
End Sub

' 2. Test de Persistencia de Riesgo (CRUD completo)
Private Sub Test_RiesgoRepository_CRUD()
    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_IDTest As Long
    
    On Error GoTo errores
    
    ' A. Crear objeto de prueba
    Set m_Riesgo = New riesgo
    With m_Riesgo
        .IDEdicion = 1 ' Usamos una edición que sepamos que existe en entorno de test
        .Descripcion = "TEST RIESGO REFACTOR"
        .Origen = "TEST"
        .Estado = "Incompleto"
    End With
    
    ' B. Test INSERT
    RiesgoRepository.Save m_Riesgo, , m_Error
    If m_Error <> "" Then
        Debug.Print "[FAIL] Test_RiesgoRepository_CRUD (Insert): " & m_Error
        Exit Sub
    End If
    m_IDTest = m_Riesgo.IDRiesgo
    Debug.Print "[OK] Test_RiesgoRepository_CRUD: Riesgo insertado con ID " & m_IDTest
    
    ' C. Test GET y UPDATE
    Set m_Riesgo = RiesgoRepository.GetById(m_IDTest, , m_Error)
    m_Riesgo.Descripcion = "TEST RIESGO ACTUALIZADO"
    RiesgoRepository.Save m_Riesgo, , m_Error
    
    Set m_Riesgo = RiesgoRepository.GetById(m_IDTest, , m_Error)
    If m_Riesgo.Descripcion = "TEST RIESGO ACTUALIZADO" Then
        Debug.Print "[OK] Test_RiesgoRepository_CRUD: Riesgo actualizado correctamente."
    Else
        Debug.Print "[FAIL] Test_RiesgoRepository_CRUD (Update): No se persistió el cambio."
    End If
    
    ' D. Test DELETE (Limpieza)
    RiesgoRepository.Delete m_IDTest, , m_Error
    Debug.Print "[OK] Test_RiesgoRepository_CRUD: Riesgo de prueba eliminado."
    
    Exit Sub
errores:
    Debug.Print "[ERROR] Test_RiesgoRepository_CRUD: " & Err.Description
End Sub
