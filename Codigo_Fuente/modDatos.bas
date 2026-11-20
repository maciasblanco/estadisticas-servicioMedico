Attribute VB_Name = "modDatos"
'Attribute VB_Name = "modDatos"
Option Explicit

' -------------------------------------------------------------------
'  MÓDULO: modDatos
'  Gestión de la hoja de trabajo (copia CRUD de DatosOriginales)
'
'  Reglas:
'   1) DatosOriginales SIEMPRE debe existir.
'   2) Si no existe, se ofrece crearla vacía (solo encabezados)
'      únicamente si estamos en un nuevo año (no se pierden históricos).
'   3) Las operaciones del CRUD trabajan sobre la copia:
'      "CRUD Datos desde dd/mm/yyyy hasta dd/mm/yyyy"
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' Devuelve el nombre de la hoja de trabajo activa.
' Si no existe la original, lanza error (o pregunta si crear una nueva).
' Si la copia no existe, la crea a partir de la original.
' -------------------------------------------------------------------
Public Function ObtenerHojaTrabajo() As String
    ' 1) La hoja original es obligatoria
    If Not ExisteHoja(HOJA_DATOS_ORIGINAL) Then
        If Not OfrecerCrearHojaOriginal() Then
            ObtenerHojaTrabajo = ""
            Exit Function
        End If
    End If
    
    ' 2) Nombre de la copia
    Dim nombre As String
    nombre = NombreHojaTrabajo()
    
    ' 3) Crear copia si no existe
    If Not ExisteHoja(nombre) Then
        CrearHojaTrabajo nombre
    End If
    
    ObtenerHojaTrabajo = nombre
End Function

' -------------------------------------------------------------------
' Pregunta al usuario si desea crear una nueva hoja original vacía.
' Solo se permite si estamos en un nuevo año (para no perder históricos).
' Devuelve TRUE si se creó, FALSE si se canceló o no procede.
' -------------------------------------------------------------------
Private Function OfrecerCrearHojaOriginal() As Boolean
    OfrecerCrearHojaOriginal = False
    
    ' Verificar que sea inicio de nuevo año
    Dim anioSistema As Integer
    anioSistema = CInt(SISTEMA_ANIO)
    
    If Year(Date) <> anioSistema Then
        MsgBox "No existe la hoja original '" & HOJA_DATOS_ORIGINAL & "'." & vbCrLf & vbCrLf & _
               "Solo se permite crear una nueva hoja de datos originales " & vbCrLf & _
               "al inicio de un nuevo año (año del sistema: " & anioSistema & ")." & vbCrLf & vbCrLf & _
               "Año actual: " & Year(Date) & ".", _
               vbCritical, "No se puede crear la hoja"
        Exit Function
    End If
    
    ' Preguntar al usuario
    Dim resp As VbMsgBoxResult
    resp = MsgBox("No existe la hoja original '" & HOJA_DATOS_ORIGINAL & "'." & vbCrLf & vbCrLf & _
                  "¿Desea crear una nueva hoja de datos originales vacía?" & vbCrLf & vbCrLf & _
                  "Se creará con las mismas columnas y título," & vbCrLf & _
                  "pero SIN datos (solo encabezados)." & vbCrLf & vbCrLf & _
                  "Esta acción solo debería hacerse al inicio de un nuevo año.", _
                  vbYesNo + vbQuestion, "Crear hoja original")
    
    If resp = vbNo Then Exit Function
    
    ' Crear hoja original vacía
    CrearHojaOriginalVacia
    OfrecerCrearHojaOriginal = True
End Function

' -------------------------------------------------------------------
' Crea la hoja 'DatosOriginales' vacía (solo encabezados).
' -------------------------------------------------------------------
Private Sub CrearHojaOriginalVacia()
    On Error GoTo Errores
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.Name = HOJA_DATOS_ORIGINAL
    
    Dim encabezados As Variant
    encabezados = Array("FECHA", "NOMBRE", "CEDULA", "SEXO", "EDAD", _
                        "ATENCION", "MEDICO", "CANT", "PROCEDENCIA", "DIAGNOSTICO")
    
    Dim i As Long
    For i = LBound(encabezados) To UBound(encabezados)
        ws.Cells(1, i - LBound(encabezados) + 1).Value = encabezados(i)
    Next i
    
    With ws.Range("A1:J1")
        .Font.Bold = True
        .Interior.Color = COLOR_PRIMARIO
        .Font.Color = COLOR_TEXTO_CLARO
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Columns("A:J").AutoFit
    ws.Columns("B").ColumnWidth = 30
    
    RegistrarHistorico ACC_CONFIG, HOJA_DATOS_ORIGINAL, "", "CREAR_ORIGINAL", "", "", _
                       "Hoja original creada vacía para nuevo año " & Year(Date)
    
    MsgBox "Hoja '" & HOJA_DATOS_ORIGINAL & "' creada vacía." & vbCrLf & vbCrLf & _
           "Ya puede pegar o ingresar los datos del nuevo año.", _
           vbInformation, "Hoja creada"
    
Salir:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
Errores:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Error al crear la hoja original:" & vbCrLf & Err.Description, vbCritical
End Sub

' -------------------------------------------------------------------
' Crea la hoja de trabajo copiando los datos de DatosOriginales.
' -------------------------------------------------------------------
Public Sub CrearHojaTrabajo(ByVal nombreDestino As String)
    On Error GoTo Errores
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Si ya existe, no hacer nada
    If ExisteHoja(nombreDestino) Then GoTo Salir
    
    ' La original debe existir
    If Not ExisteHoja(HOJA_DATOS_ORIGINAL) Then
        MsgBox "No existe la hoja original '" & HOJA_DATOS_ORIGINAL & "'.", vbCritical
        GoTo Salir
    End If
    
    Dim encabezados As Variant
    encabezados = Array("FECHA", "NOMBRE", "CEDULA", "SEXO", "EDAD", _
                        "ATENCION", "MEDICO", "CANT", "PROCEDENCIA", "DIAGNOSTICO")
    
    Dim wsNueva As Worksheet
    Set wsNueva = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsNueva.Name = nombreDestino
    
    ' Encabezados
    Dim i As Long
    For i = LBound(encabezados) To UBound(encabezados)
        wsNueva.Cells(1, i - LBound(encabezados) + 1).Value = encabezados(i)
    Next i
    With wsNueva.Range("A1:J1")
        .Font.Bold = True
        .Interior.Color = COLOR_PRIMARIO
        .Font.Color = COLOR_TEXTO_CLARO
        .HorizontalAlignment = xlCenter
    End With
    
    ' Copiar datos de la original
    Dim wsOrig As Worksheet
    Set wsOrig = ThisWorkbook.Sheets(HOJA_DATOS_ORIGINAL)
    Dim lastRow As Long
    lastRow = wsOrig.Cells(wsOrig.Rows.count, COL_CEDULA).End(xlUp).Row
    
    If lastRow >= 2 Then
        wsOrig.Range(wsOrig.Cells(2, 1), _
                     wsOrig.Cells(lastRow, 10)).Copy _
                     Destination:=wsNueva.Cells(2, 1)
    End If
    
    wsNueva.Columns("A:J").AutoFit
    wsNueva.Columns("B").ColumnWidth = 30
    
    ' Registrar en auditoría
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

' -------------------------------------------------------------------
' Devuelve TRUE si la hoja de trabajo ya existe.
' -------------------------------------------------------------------
Public Function ExisteHojaTrabajo() As Boolean
    ExisteHojaTrabajo = ExisteHoja(NombreHojaTrabajo())
End Function

' -------------------------------------------------------------------
' Elimina la hoja de trabajo y la recrea desde cero (copia limpia).
' -------------------------------------------------------------------
Public Sub RecrearHojaTrabajo()
    If Not ExisteHoja(HOJA_DATOS_ORIGINAL) Then
        MsgBox "No existe la hoja original '" & HOJA_DATOS_ORIGINAL & "'.", vbCritical
        Exit Sub
    End If
    
    Dim nombre As String
    nombre = NombreHojaTrabajo()
    
    If MsgBox("¿Recrear la hoja de trabajo '" & nombre & "'?" & vbCrLf & vbCrLf & _
              "Se perderán TODOS los cambios hechos en ella." & vbCrLf & _
              "La hoja original '" & HOJA_DATOS_ORIGINAL & "' NO se verá afectada.", _
              vbYesNo + vbQuestion, "Recrear hoja de trabajo") = vbNo Then Exit Sub
    
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(nombre).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    CrearHojaTrabajo nombre
    MsgBox "Hoja de trabajo recreada correctamente:" & vbCrLf & nombre, _
           vbInformation, "Listo"
End Sub

' -------------------------------------------------------------------
' Verifica que la hoja original exista y tenga la estructura esperada.
' -------------------------------------------------------------------
Public Function ValidarHojaOriginal() As Boolean
    ValidarHojaOriginal = False
    
    If Not ExisteHoja(HOJA_DATOS_ORIGINAL) Then
        MsgBox "No existe la hoja original '" & HOJA_DATOS_ORIGINAL & "'.", _
               vbCritical, "Hoja original no encontrada"
        Exit Function
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_DATOS_ORIGINAL)
    
    If CStr(ws.Cells(FILA_ENCABEZADO, COL_FECHA).Value) = "" Then
        MsgBox "La hoja '" & HOJA_DATOS_ORIGINAL & "' no tiene el encabezado FECHA en la columna A.", _
               vbCritical, "Estructura incorrecta"
        Exit Function
    End If
    
    ValidarHojaOriginal = True
End Function

