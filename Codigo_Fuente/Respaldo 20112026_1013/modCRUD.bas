Attribute VB_Name = "modCRUD"
'Option Explicit

' -------------------------------------------------------------------
'  MÓDULO: modCRUD (Fase 4)
'  Operaciones de Creación, Lectura, Actualización y Borrado
' -------------------------------------------------------------------

Public Function BuscarRegistros(ByVal criterio As String, ByRef resultados() As Variant) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("DatosOriginales")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, COL_CEDULA).End(xlUp).Row
    Dim count As Long: count = 0
    criterio = UCase(Trim(criterio))
    Dim i As Long
    
    ' Primera pasada: contar cuántos registros coinciden
    For i = 2 To lastRow
        Dim textoFila As String
        textoFila = UCase(CStr(ws.Cells(i, COL_NOMBRE).Value) & "|" & CStr(ws.Cells(i, COL_CEDULA).Value))
        If Len(criterio) = 0 Or InStr(1, textoFila, criterio, vbTextCompare) > 0 Then
            count = count + 1
        End If
    Next i
    
    If count = 0 Then
        BuscarRegistros = 0
        Exit Function
    End If
    
    ' Redimensionar la matriz de resultados
    ReDim resultados(1 To count, 1 To 10)
    Dim k As Long: k = 0
    
    ' Segunda pasada: llenar la matriz
    For i = 2 To lastRow
        Dim textoFila2 As String
        textoFila2 = UCase(CStr(ws.Cells(i, COL_NOMBRE).Value) & "|" & CStr(ws.Cells(i, COL_CEDULA).Value))
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
End Function

Public Function InsertarRegistro(ByRef valores() As Variant) As Boolean
    On Error GoTo Errores
    
    If Not TienePermiso("INSERT") Then
        MsgBox MSG_SIN_PERMISO, vbExclamation, "Permiso Denegado"
        InsertarRegistro = False: Exit Function
    End If
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(HOJA_DATOS)
    Dim fila As Long: fila = ws.Cells(ws.Rows.count, COL_CEDULA).End(xlUp).Row + 1
    Dim i As Long
    Dim detalle As String
    
    ' Escribir en Excel y armar el texto de auditoría
    For i = LBound(valores) To UBound(valores)
        ws.Cells(fila, i - LBound(valores) + 1).Value = valores(i)
        detalle = detalle & CStr(valores(i)) & " | "
    Next i
    
    ' Guardar en historial usando el string armado manualmente (evita el error de Join)
    RegistrarHistorico ACC_INSERT, HOJA_DATOS, fila, "REGISTRO COMPLETO", "", detalle, "Registro creado"
    InsertarRegistro = True
    Exit Function
Errores:
    MsgBox "Error al insertar: " & Err.Description, vbCritical: InsertarRegistro = False
End Function

Public Function ActualizarRegistro(ByVal fila As Long, ByRef valores() As Variant) As Boolean
    On Error GoTo Errores
    
    If Not TienePermiso("UPDATE") Then
        MsgBox MSG_SIN_PERMISO, vbExclamation, "Permiso Denegado"
        ActualizarRegistro = False: Exit Function
    End If
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(HOJA_DATOS)
    Dim i As Long, col As Long
    
    For i = LBound(valores) To UBound(valores)
        col = i - LBound(valores) + 1
        Dim valAnt As Variant: valAnt = ws.Cells(fila, col).Value
        If CStr(valAnt) <> CStr(valores(i)) Then
            ws.Cells(fila, col).Value = valores(i)
            RegistrarHistorico ACC_UPDATE, HOJA_DATOS, fila, CStr(ws.Cells(FILA_ENCABEZADO, col).Value), valAnt, valores(i), "Modificación"
        End If
    Next i
    
    ActualizarRegistro = True
    Exit Function
Errores:
    MsgBox "Error al actualizar: " & Err.Description, vbCritical: ActualizarRegistro = False
End Function

Public Function EliminarRegistro(ByVal fila As Long) As Boolean
    On Error GoTo Errores
    
    If Not TienePermiso("DELETE") Then
        MsgBox MSG_SIN_PERMISO, vbExclamation, "Permiso Denegado"
        EliminarRegistro = False: Exit Function
    End If
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(HOJA_DATOS)
    Dim backup As String
    Dim i As Long
    
    For i = 1 To 10
        backup = backup & CStr(ws.Cells(fila, i).Value) & " | "
    Next i
    
    RegistrarHistorico ACC_DELETE, HOJA_DATOS, fila, "REGISTRO COMPLETO", backup, "", "Registro eliminado"
    ws.Rows(fila).Delete Shift:=xlUp
    EliminarRegistro = True
    Exit Function
Errores:
    MsgBox "Error al eliminar: " & Err.Description, vbCritical: EliminarRegistro = False
End Function
' -------------------------------------------------------------------
' FASE 7: Exportación desde frmCRUD
' -------------------------------------------------------------------
Private Sub cmdExportar_Click()
    If Not VerificarPermisoExport() Then Exit Sub
    
    ' Preguntar formato
    Dim resp As VbMsgBoxResult
    resp = MsgBox("¿En qué formato desea exportar los resultados?" & vbCrLf & vbCrLf & _
                  "SÍ = Excel (.xlsx)" & vbCrLf & _
                  "NO = PDF" & vbCrLf & _
                  "CANCELAR = Abortar", _
                  vbYesNoCancel + vbQuestion, "Exportar resultados")
    
    If resp = vbCancel Then Exit Sub
    
    Dim extension As String
    Dim filtro As String
    If resp = vbYes Then
        extension = "xlsx"
        filtro = "Libro de Excel (*.xlsx), *.xlsx"
    Else
        extension = "pdf"
        filtro = "Documento PDF (*.pdf), *.pdf"
    End If
    
    ' Asegurar carpeta de exportaciones
    If Not AsegurarCarpeta(CARPETA_EXPORT) Then Exit Sub
    
    ' Nombre por defecto
    Dim nombreDef As String
    nombreDef = CARPETA_EXPORT & GenerarNombreArchivo("BusquedaPacientes", extension)
    
    Dim ruta As String
    ruta = SeleccionarRutaGuardar(filtro, nombreDef)
    If ruta = "" Then Exit Sub
    
    Dim ok As Boolean
    If resp = vbYes Then
        ok = ExportarResultadosBusqueda(ruta, Me.txtBuscar.Value)
    Else
        ' Para PDF: generar Excel temporal, luego PDF
        Dim rutaTmp As String
        rutaTmp = Environ$("TEMP") & "\tmp_busqueda_" & Format(Now, "yyyymmddhhmmss") & ".xlsx"
        If ExportarResultadosBusqueda(rutaTmp, Me.txtBuscar.Value) Then
            ok = ExportarHojaAExcelTemporalAPDF(rutaTmp, ruta, "Resultados de Búsqueda")
            On Error Resume Next
            Kill rutaTmp
            On Error GoTo 0
        End If
    End If
    
    If ok Then
        MsgBox "Exportación exitosa:" & vbCrLf & ruta, vbInformation, "Listo"
        If MsgBox("¿Desea abrir la carpeta?", vbYesNo + vbQuestion, "Abrir") = vbYes Then
            AbrirCarpeta CARPETA_EXPORT
        End If
    Else
        MsgBox "No se pudo completar la exportación.", vbExclamation, "Error"
    End If
End Sub

' Helper para exportar un .xlsx ya generado a PDF
Private Function ExportarHojaAExcelTemporalAPDF(ByVal rutaXlsx As String, _
                                                ByVal rutaPDF As String, _
                                                ByVal titulo As String) As Boolean
    ExportarHojaAExcelTemporalAPDF = False
    On Error GoTo Errores
    
    Dim wbTmp As Workbook
    Set wbTmp = Workbooks.Open(rutaXlsx, ReadOnly:=True)
    
    wbTmp.Sheets(1).ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=rutaPDF, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    
    wbTmp.Close SaveChanges:=False
    ExportarHojaAExcelTemporalAPDF = True
    Exit Function
Errores:
    If Not wbTmp Is Nothing Then wbTmp.Close SaveChanges:=False
    MsgBox "Error al convertir a PDF: " & Err.Description, vbCritical
End Function
