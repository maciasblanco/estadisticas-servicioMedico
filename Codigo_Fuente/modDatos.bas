Attribute VB_Name = "modDatos"
'Attribute VB_Name = "modDatos"
Option Explicit

' -------------------------------------------------------------------
'  MÓDULO: modDatos
'  Gestión de la hoja de trabajo (copia CRUD de DatosOriginales)
' -------------------------------------------------------------------

Public Function ObtenerHojaTrabajo() As String
    Dim nombre As String
    nombre = NombreHojaTrabajo()
    
    If Not ExisteHoja(nombre) Then
        CrearHojaTrabajo nombre
    End If
    
    ObtenerHojaTrabajo = nombre
End Function

Public Sub CrearHojaTrabajo(ByVal nombreDestino As String)
    On Error GoTo Errores
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If ExisteHoja(nombreDestino) Then GoTo Salir
    
    If Not ExisteHoja(HOJA_DATOS_ORIGINAL) Then
        MsgBox "No existe la hoja '" & HOJA_DATOS_ORIGINAL & "'.", vbCritical
        GoTo Salir
    End If
    
    ' Copiar la hoja original completa (conserva formatos y datos)
    Dim wsOrig As Worksheet
    Set wsOrig = ThisWorkbook.Sheets(HOJA_DATOS_ORIGINAL)
    wsOrig.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    
    Dim wsNueva As Worksheet
    Set wsNueva = ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    wsNueva.Name = nombreDestino
    
    RegistrarHistorico ACC_CONFIG, nombreDestino, "", "CREAR_COPIA", "", "", _
                       "Hoja de trabajo creada desde " & HOJA_DATOS_ORIGINAL
                       
Salir:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
Errores:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Error al crear la hoja de trabajo:" & vbCrLf & Err.Description, vbCritical
End Sub

Public Function ExisteHojaTrabajo() As Boolean
    ExisteHojaTrabajo = ExisteHoja(NombreHojaTrabajo())
End Function

