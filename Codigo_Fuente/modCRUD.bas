Attribute VB_Name = "modCRUD"
'Attribute VB_Name = "modCRUD"
Option Explicit

' -------------------------------------------------------------------
'  MÓDULO: modCRUD
'  Operaciones de Creación, Lectura, Actualización y Borrado
'  Trabaja SIEMPRE sobre la copia "CRUD Datos desde ... hasta ..."
'  La hoja DatosOriginales NUNCA se modifica.
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' BUSCAR: devuelve el número de coincidencias y llena la matriz.
' -------------------------------------------------------------------
Public Function BuscarRegistros(ByVal criterio As String, _
                                ByRef resultados() As Variant) As Long
    On Error GoTo Errores
    
    ' 1) Obtener la hoja de trabajo (crea la copia si no existe)
    Dim nombreHoja As String
    nombreHoja = ObtenerHojaTrabajo()
    If nombreHoja = "" Then
        BuscarRegistros = 0
        Exit Function
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, COL_CEDULA).End(xlUp).Row
    If lastRow < 2 Then
        BuscarRegistros = 0
        Exit Function
    End If
    
    Dim count As Long: count = 0
    criterio = UCase(Trim(criterio))
    Dim i As Long
    
    ' Primera pasada: contar coincidencias
    For i = 2 To lastRow
        Dim textoFila As String
        textoFila = UCase(CStr(ws.Cells(i, COL_NOMBRE).Value) & "|" & _
                          CStr(ws.Cells(i, COL_CEDULA).Value))
        If Len(criterio) = 0 Or InStr(1, textoFila, criterio, vbTextCompare) > 0 Then
            count = count + 1
        End If
    Next i
    
    If count = 0 Then
        BuscarRegistros = 0
        Exit Function
    End If
    
    ' Redimensionar matriz
    ReDim resultados(1 To count, 1 To 10)
    Dim k As Long: k = 0
    
    ' Segunda pasada: llenar matriz
    For i = 2 To lastRow
        Dim textoFila2 As String
        textoFila2 = UCase(CStr(ws.Cells(i, COL_NOMBRE).Value) & "|" & _
                           CStr(ws.Cells(i, COL_CEDULA).Value))
        If Len(criterio) = 0 Or InStr(1, textoFila2, criterio, vbTextCompare) > 0 Then
            k = k + 1
            resultados(k, 1) = ws.Cells(i, COL_FECHA).Value
            resultados(k, 2) = ws.Cells(i, COL_NOMBRE).Value
            resultados(k, 3) = ws.Cells(i, COL_CEDULA).Value
            resultados(k, 4) = ws.Cells(i, COL_SEXO).Value
            resultados(k, 5) = ws.Cells(i, COL_EDAD).Value
            resultados(k, 6) = ws.Cells(i, COL_ATENCION).Value
            resultados(k, 7) = ws.Cells(i, COL_MEDICO).Value
            resultados(k, 8) = ws.Cells(i, COL_CANT).Value
            resultados(k, 9) = ws.Cells(i, COL_PROCEDENCIA).Value
            resultados(k, 10) = ws.Cells(i, COL_DIAGNOSTICO).Value
        End If
    Next i
    
    BuscarRegistros = count
    Exit Function
Errores:
    MsgBox "Error al buscar registros:" & vbCrLf & Err.Description, vbCritical
    BuscarRegistros = 0
End Function

' -------------------------------------------------------------------
' INSERTAR: agrega un nuevo registro al final de la hoja de trabajo.
' -------------------------------------------------------------------
Public Function InsertarRegistro(ByRef valores() As Variant) As Boolean
    On Error GoTo Errores
    InsertarRegistro = False
    
    If Not TienePermiso("INSERT") Then
        MsgBox MSG_SIN_PERMISO, vbExclamation, "Permiso Denegado"
        Exit Function
    End If
    
    Dim nombreHoja As String
    nombreHoja = ObtenerHojaTrabajo()
    If nombreHoja = "" Then Exit Function
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    
    ' Validar cédula duplicada
    Dim cedulaNueva As String: cedulaNueva = Trim(CStr(valores(3)))
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, COL_CEDULA).End(xlUp).Row
    Dim j As Long
    For j = 2 To lastRow
        If Trim(CStr(ws.Cells(j, COL_CEDULA).Value)) = cedulaNueva Then
            MsgBox "Ya existe un registro con la cédula " & cedulaNueva & ".", _
                   vbExclamation, "Cédula duplicada"
            Exit Function
        End If
    Next j
    
    ' Escribir en la primera fila libre
    Dim fila As Long: fila = lastRow + 1
    Dim i As Long, detalle As String
    For i = LBound(valores) To UBound(valores)
        ws.Cells(fila, i - LBound(valores) + 1).Value = valores(i)
        detalle = detalle & CStr(valores(i)) & " | "
    Next i
    
    RegistrarHistorico ACC_INSERT, nombreHoja, fila, "REGISTRO COMPLETO", "", detalle, "Registro creado"
    InsertarRegistro = True
    Exit Function
Errores:
    MsgBox "Error al insertar:" & vbCrLf & Err.Description, vbCritical
    InsertarRegistro = False
End Function

' -------------------------------------------------------------------
' ACTUALIZAR: modifica una fila existente en la hoja de trabajo.
' -------------------------------------------------------------------
Public Function ActualizarRegistro(ByVal fila As Long, _
                                   ByRef valores() As Variant) As Boolean
    On Error GoTo Errores
    ActualizarRegistro = False
    
    If Not TienePermiso("UPDATE") Then
        MsgBox MSG_SIN_PERMISO, vbExclamation, "Permiso Denegado"
        Exit Function
    End If
    
    If fila < 2 Then
        MsgBox "Fila inválida para actualizar.", vbExclamation
        Exit Function
    End If
    
    Dim nombreHoja As String
    nombreHoja = ObtenerHojaTrabajo()
    If nombreHoja = "" Then Exit Function
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    
    Dim i As Long, col As Long
    For i = LBound(valores) To UBound(valores)
        col = i - LBound(valores) + 1
        Dim valAnt As Variant: valAnt = ws.Cells(fila, col).Value
        If CStr(valAnt) <> CStr(valores(i)) Then
            ws.Cells(fila, col).Value = valores(i)
            RegistrarHistorico ACC_UPDATE, nombreHoja, fila, _
                CStr(ws.Cells(FILA_ENCABEZADO, col).Value), _
                valAnt, valores(i), "Modificación"
        End If
    Next i
    
    ActualizarRegistro = True
    Exit Function
Errores:
    MsgBox "Error al actualizar:" & vbCrLf & Err.Description, vbCritical
    ActualizarRegistro = False
End Function

' -------------------------------------------------------------------
' ELIMINAR: borra una fila de la hoja de trabajo.
' -------------------------------------------------------------------
Public Function EliminarRegistro(ByVal fila As Long) As Boolean
    On Error GoTo Errores
    EliminarRegistro = False
    
    If Not TienePermiso("DELETE") Then
        MsgBox MSG_SIN_PERMISO, vbExclamation, "Permiso Denegado"
        Exit Function
    End If
    
    If fila < 2 Then
        MsgBox "Fila inválida para eliminar.", vbExclamation
        Exit Function
    End If
    
    Dim nombreHoja As String
    nombreHoja = ObtenerHojaTrabajo()
    If nombreHoja = "" Then Exit Function
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    
    ' Respaldar la fila antes de borrar
    Dim backup As String, i As Long
    For i = 1 To 10
        backup = backup & CStr(ws.Cells(fila, i).Value) & " | "
    Next i
    
    RegistrarHistorico ACC_DELETE, nombreHoja, fila, "REGISTRO COMPLETO", backup, "", "Registro eliminado"
    ws.Rows(fila).Delete Shift:=xlUp
    EliminarRegistro = True
    Exit Function
Errores:
    MsgBox "Error al eliminar:" & vbCrLf & Err.Description, vbCritical
    EliminarRegistro = False
End Function

