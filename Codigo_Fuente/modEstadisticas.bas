Attribute VB_Name = "modEstadisticas"
Attribute VB_Name = "modEstadisticas"
Option Explicit

' -------------------------------------------------------------------
'  MÓDULO: modEstadisticas (FASE 8)
'  Genera el reporte matricial basado en la imagen del usuario
' -------------------------------------------------------------------

Public Sub GenerarReporteEstadistico(ByVal fIni As Date, ByVal fFin As Date, ByVal tituloPeriodo As String)
    ' 1. Validar permisos
    If Not TienePermiso("ESTADISTICA") Then
        MsgBox "No tiene permisos para generar estadísticas.", vbExclamation, "Acceso Denegado"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 2. Preparar la hoja del reporte
    Dim wsRep As Worksheet
    Dim nombreHoja As String: nombreHoja = "Reporte_" & Format(fIni, "yyyymmdd")
    
    ' Si la hoja ya existe, la borramos para crearla de nuevo limpia
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(nombreHoja).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Crear nueva hoja al final
    Set wsRep = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsRep.Name = nombreHoja

    ' 3. Leer datos de la hoja original
    Dim wsDatos As Worksheet: Set wsDatos = ThisWorkbook.Sheets(HOJA_DATOS)
    Dim lastRow As Long: lastRow = wsDatos.Cells(wsDatos.Rows.count, COL_CEDULA).End(xlUp).Row
    
    ' Diccionarios para contar de forma rápida
    Dim dictCat As Object, dictSist As Object, dictEcos As Object, dictMed As Object
    Set dictCat = CreateObject("Scripting.Dictionary")
    Set dictSist = CreateObject("Scripting.Dictionary")
    Set dictEcos = CreateObject("Scripting.Dictionary")
    Set dictMed = CreateObject("Scripting.Dictionary")

    Dim i As Long, fechaFila As Date, cat As String, sist As String, med As String, cant As Double
    
    ' 4. Recorrer y agrupar datos
    For i = 2 To lastRow
        If IsDate(wsDatos.Cells(i, COL_FECHA).Value) Then
            fechaFila = CDate(wsDatos.Cells(i, COL_FECHA).Value)
            
            ' Filtrar por rango de fechas
            If fechaFila >= fIni And fechaFila <= fFin Then
                cat = UCase(Trim(CStr(wsDatos.Cells(i, COL_PROCEDENCIA).Value))) ' Asumiendo que Categoría es Procedencia
                sist = UCase(Trim(CStr(wsDatos.Cells(i, COL_DIAGNOSTICO).Value)))
                med = UCase(Trim(CStr(wsDatos.Cells(i, COL_MEDICO).Value)))
                cant = Val(wsDatos.Cells(i, COL_CANT).Value)
                If cant = 0 Then cant = 1 ' Si está vacío, cuenta como 1
                
                ' Sumar a Categoría
                If cat <> "" Then dictCat(cat) = dictCat(cat) + cant
                
                ' Sumar a Sistema/Diagnóstico
                If sist <> "" Then dictSist(sist) = dictSist(sist) + cant
                
                ' Sumar a Médicos
                If med <> "" Then dictMed(med) = dictMed(med) + cant
            End If
        End If
    Next i

    ' 5. Dibujar el reporte en la hoja
    DibujarReporte wsRep, dictCat, dictSist, dictMed, tituloPeriodo

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Reporte generado exitosamente en la hoja: " & nombreHoja, vbInformation, "Éxito"
End Sub

' -------------------------------------------------------------------
' Función que dibuja la tabla con el formato de la imagen
' -------------------------------------------------------------------
Private Sub DibujarReporte(ws As Worksheet, dictCat As Object, dictSist As Object, dictMed As Object, titulo As String)
    Dim fila As Long: fila = 1
    Dim col As Long
    
    ' --- TÍTULO PRINCIPAL ---
    ws.Cells(fila, 1).Value = "ESTADÍSTICAS DEL SERVICIO MÉDICO " & UCase(titulo)
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 3)).Merge
    ws.Cells(fila, 1).Font.Size = 14
    ws.Cells(fila, 1).Font.Bold = True
    ws.Cells(fila, 1).HorizontalAlignment = xlCenter
    fila = fila + 1
    
    ' --- ENCABEZADOS ---
    ws.Cells(fila, 1).Value = "CATEGORÍA"
    ws.Cells(fila, 2).Value = "TIPO"
    ws.Cells(fila, 3).Value = "TOTAL"
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 3)).Interior.Color = RGB(200, 220, 240)
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 3)).Font.Bold = True
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 3)).Borders.LineStyle = xlContinuous
    fila = fila + 1

    ' --- BLOQUE 1: CATEGORIA GENERAL ---
    Dim filaIniCat As Long: filaIniCat = fila
    ' Aquí pones las categorías que quieras fijas. Si quieres que sean dinámicas, se puede cambiar.
    Dim catsFijas As Variant
    catsFijas = Array("TRABAJADORES", "FAMILIARES", "PACIENTES", "JUBILADOS")
    
    Dim c As Variant, totalCat As Double
    totalCat = 0
    For Each c In catsFijas
        ws.Cells(fila, 2).Value = c
        If dictCat.Exists(c) Then
            ws.Cells(fila, 3).Value = dictCat(c)
            totalCat = totalCat + dictCat(c)
        Else
            ws.Cells(fila, 3).Value = 0
        End If
        ws.Cells(fila, 2).Borders.LineStyle = xlContinuous
        ws.Cells(fila, 3).Borders.LineStyle = xlContinuous
        fila = fila + 1
    Next c
    
    ' Total del bloque
    ws.Cells(fila, 2).Value = "TOTAL MES"
    ws.Cells(fila, 2).Font.Bold = True
    ws.Cells(fila, 3).Value = totalCat
    ws.Cells(fila, 3).Font.Bold = True
    ws.Cells(fila, 2).Borders.LineStyle = xlContinuous
    ws.Cells(fila, 3).Borders.LineStyle = xlContinuous
    
    ' Combinar celdas de la primera columna para este bloque
    ws.Range(ws.Cells(filaIniCat, 1), ws.Cells(fila, 1)).Merge
    ws.Cells(filaIniCat, 1).Value = "CATEGORÍA GENERAL"
    ws.Cells(filaIniCat, 1).HorizontalAlignment = xlCenter
    ws.Cells(filaIniCat, 1).VerticalAlignment = xlCenter
    ws.Cells(filaIniCat, 1).Font.Bold = True
    
    fila = fila + 2 ' Espacio entre bloques

    ' --- BLOQUE 2: SISTEMAS AFECTADOS ---
    Dim filaIniSist As Long: filaIniSist = fila
    ' Extraemos todas las claves (diagnósticos) del diccionario
    Dim keysSist As Variant: keysSist = dictSist.keys
    Dim totalSist As Double: totalSist = 0
    Dim k As Variant
    
    ' Si no hay datos, ponemos un texto por defecto
    If dictSist.count = 0 Then
        ws.Cells(fila, 2).Value = "SIN DATOS EN EL PERIODO"
    Else
        ' Ordenar alfabéticamente (opcional, se puede omitir)
        For Each k In keysSist
            ws.Cells(fila, 2).Value = k
            ws.Cells(fila, 3).Value = dictSist(k)
            ws.Cells(fila, 2).Borders.LineStyle = xlContinuous
            ws.Cells(fila, 3).Borders.LineStyle = xlContinuous
            totalSist = totalSist + dictSist(k)
            fila = fila + 1
        Next k
    End If
    
    ws.Cells(fila, 2).Value = "TOTAL"
    ws.Cells(fila, 2).Font.Bold = True
    ws.Cells(fila, 3).Value = totalSist
    ws.Cells(fila, 3).Font.Bold = True
    ws.Cells(fila, 2).Borders.LineStyle = xlContinuous
    ws.Cells(fila, 3).Borders.LineStyle = xlContinuous
    
    ws.Range(ws.Cells(filaIniSist, 1), ws.Cells(fila, 1)).Merge
    ws.Cells(filaIniSist, 1).Value = "SISTEMAS AFECTADOS"
    ws.Cells(filaIniSist, 1).HorizontalAlignment = xlCenter
    ws.Cells(filaIniSist, 1).VerticalAlignment = xlCenter
    ws.Cells(filaIniSist, 1).Font.Bold = True
    
    fila = fila + 2

    ' --- BLOQUE 3: CONSULTA POR MÉDICOS ---
    Dim filaIniMed As Long: filaIniMed = fila
    Dim keysMed As Variant: keysMed = dictMed.keys
    Dim totalMed As Double: totalMed = 0
    
    ' Agregamos los médicos de tu imagen (puedes agregar más)
    Dim medsFijos As Variant
    medsFijos = Array("DRA. ARBELI HERNANDEZ", "DRA. RAQUEL RODRIGUEZ", "DRA. ROSANA VARGAS", _
                      "DRA. DALITZA FERMIN", "DRA. MARKIS GARCIA", "DR. EUSEBEL CASTILLO", _
                      "DR. ISAAC MAR PEDIATRA", "DR. HUGO PADILLA CONSULTA")
                      
    For Each k In medsFijos
        ws.Cells(fila, 2).Value = k
        ' Si el médico existe en los datos, ponemos su total
        If dictMed.Exists(UCase(k)) Then
            ws.Cells(fila, 3).Value = dictMed(UCase(k))
            totalMed = totalMed + dictMed(UCase(k))
        Else
            ws.Cells(fila, 3).Value = 0
        End If
        ws.Cells(fila, 2).Borders.LineStyle = xlContinuous
        ws.Cells(fila, 3).Borders.LineStyle = xlContinuous
        fila = fila + 1
    Next k
    
    ws.Cells(fila, 2).Value = "TOTAL"
    ws.Cells(fila, 2).Font.Bold = True
    ws.Cells(fila, 3).Value = totalMed
    ws.Cells(fila, 3).Font.Bold = True
    ws.Cells(fila, 2).Borders.LineStyle = xlContinuous
    ws.Cells(fila, 3).Borders.LineStyle = xlContinuous
    
    ws.Range(ws.Cells(filaIniMed, 1), ws.Cells(fila, 1)).Merge
    ws.Cells(filaIniMed, 1).Value = "CONSULTA POR MÉDICOS"
    ws.Cells(filaIniMed, 1).HorizontalAlignment = xlCenter
    ws.Cells(filaIniMed, 1).VerticalAlignment = xlCenter
    ws.Cells(filaIniMed, 1).Font.Bold = True

    ' Ajustar ancho de columnas
    ws.Columns("A:C").AutoFit
    ws.Columns("B").ColumnWidth = 35
    ws.Columns("C").ColumnWidth = 15
End Sub

