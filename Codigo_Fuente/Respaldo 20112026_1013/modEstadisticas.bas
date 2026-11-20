Attribute VB_Name = "modEstadisticas"
'Attribute VB_Name = "modEstadisticas"
Option Explicit

' Variable global para guardar la última matriz generada y usarla al exportar
Public gMatrizReporte As Variant

' -------------------------------------------------------------------
' 1. FUNCIÓN QUE CALCULA LOS DATOS Y DEVUELVE UNA MATRIZ (Para el ListBox)
' -------------------------------------------------------------------
Public Function GenerarMatrizPrevia(ByVal fIni As Date, ByVal fFin As Date) As Variant
    Dim wsDatos As Worksheet: Set wsDatos = ThisWorkbook.Sheets("DatosOriginales")
    Dim lastRow As Long: lastRow = wsDatos.Cells(wsDatos.Rows.count, COL_CEDULA).End(xlUp).Row
    
    Dim dictCat As Object, dictSist As Object, dictMed As Object
    Set dictCat = CreateObject("Scripting.Dictionary")
    Set dictSist = CreateObject("Scripting.Dictionary")
    Set dictMed = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, fechaFila As Date, cat As String, sist As String, med As String
    
    ' Recorrer y agrupar datos
    For i = 2 To lastRow
        If IsDate(wsDatos.Cells(i, COL_FECHA).Value) Then
            fechaFila = CDate(wsDatos.Cells(i, COL_FECHA).Value)
            
            ' Filtrar por rango de fechas
            If fechaFila >= fIni And fechaFila <= fFin Then
                ' OJO: Usamos COL_PROCEDENCIA para Categoría (según tu reporte final)
                ' y COL_ATENCION también es válido según tu captura, pero usemos Procedencia para mantener lógica
                cat = UCase(Trim(CStr(wsDatos.Cells(i, COL_PROCEDENCIA).Value)))
                sist = UCase(Trim(CStr(wsDatos.Cells(i, COL_DIAGNOSTICO).Value)))
                med = UCase(Trim(CStr(wsDatos.Cells(i, COL_MEDICO).Value)))
                
                ' Si cat está vacío, le asignamos un nombre para no perder el dato
                If cat = "" Then cat = "SIN CATEGORÍA"
                
                If cat <> "" Then dictCat(cat) = dictCat(cat) + 1
                If sist <> "" Then dictSist(sist) = dictSist(sist) + 1
                If med <> "" Then dictMed(med) = dictMed(med) + 1
            End If
        End If
    Next i
    
    ' --- VALIDACIÓN PARA EVITAR ERROR 9 ---
    ' Si todos los diccionarios están vacíos, no hay datos en el período
    If dictCat.count = 0 And dictSist.count = 0 And dictMed.count = 0 Then
        GenerarMatrizPrevia = Empty
        Exit Function
    End If
    
    ' --- CALCULAR TAMAÑO DE LA MATRIZ ---
    Dim totalFilas As Long
    totalFilas = dictCat.count + dictSist.count + dictMed.count
    
    Dim tempArr() As Variant
    ReDim tempArr(1 To totalFilas, 1 To 3) ' 3 Columnas: Tipo, Nombre, Total
    
    Dim k As Variant, filaActual As Long
    filaActual = 1
    
    ' Llenar Categorías (Procedencia)
    For Each k In dictCat.keys
        tempArr(filaActual, 1) = "CATEGORÍA"
        tempArr(filaActual, 2) = k
        tempArr(filaActual, 3) = dictCat(k)
        filaActual = filaActual + 1
    Next k
    
    ' Llenar Sistemas (Diagnóstico)
    For Each k In dictSist.keys
        tempArr(filaActual, 1) = "SISTEMA AFECTADO"
        tempArr(filaActual, 2) = k
        tempArr(filaActual, 3) = dictSist(k)
        filaActual = filaActual + 1
    Next k
    
    ' Llenar Médicos
    For Each k In dictMed.keys
        tempArr(filaActual, 1) = "MÉDICO"
        tempArr(filaActual, 2) = k
        tempArr(filaActual, 3) = dictMed(k)
        filaActual = filaActual + 1
    Next k
    
    GenerarMatrizPrevia = tempArr
    gMatrizReporte = tempArr ' Guardamos globalmente para exportar luego
End Function

' -------------------------------------------------------------------
' 2. SUBRUTINA PARA GENERAR EL REPORTE EN UNA HOJA NUEVA
' -------------------------------------------------------------------
Public Sub GenerarReporteEstadistico(ByVal fIni As Date, ByVal fFin As Date, ByVal tituloPeriodo As String)
    If Not TienePermiso("ESTADISTICA") Then
        MsgBox "No tiene permisos para generar estadísticas.", vbExclamation, "Acceso Denegado"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsRep As Worksheet
    Dim nombreHoja As String: nombreHoja = "Reporte_" & Format(fIni, "yyyymmdd")
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(nombreHoja).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsRep = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsRep.Name = nombreHoja

    Dim wsDatos As Worksheet: Set wsDatos = ThisWorkbook.Sheets(HOJA_DATOS)
    Dim lastRow As Long: lastRow = wsDatos.Cells(wsDatos.Rows.count, COL_CEDULA).End(xlUp).Row
    
    Dim dictCat As Object, dictSist As Object, dictMed As Object
    Set dictCat = CreateObject("Scripting.Dictionary")
    Set dictSist = CreateObject("Scripting.Dictionary")
    Set dictMed = CreateObject("Scripting.Dictionary")

    Dim i As Long, fechaFila As Date, cat As String, sist As String, med As String, cant As Double
    
    For i = 2 To lastRow
        If IsDate(wsDatos.Cells(i, COL_FECHA).Value) Then
            fechaFila = CDate(wsDatos.Cells(i, COL_FECHA).Value)
            
            If fechaFila >= fIni And fechaFila <= fFin Then
                cat = UCase(Trim(CStr(wsDatos.Cells(i, COL_PROCEDENCIA).Value)))
                sist = UCase(Trim(CStr(wsDatos.Cells(i, COL_DIAGNOSTICO).Value)))
                med = UCase(Trim(CStr(wsDatos.Cells(i, COL_MEDICO).Value)))
                cant = Val(wsDatos.Cells(i, COL_CANT).Value)
                If cant = 0 Then cant = 1
                
                If cat <> "" Then dictCat(cat) = dictCat(cat) + cant
                If sist <> "" Then dictSist(sist) = dictSist(sist) + cant
                If med <> "" Then dictMed(med) = dictMed(med) + cant
            End If
        End If
    Next i

    DibujarReporte wsRep, dictCat, dictSist, dictMed, tituloPeriodo

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Reporte generado exitosamente en la hoja: " & nombreHoja, vbInformation, "Éxito"
End Sub

' -------------------------------------------------------------------
' 3. DIBUJAR EL REPORTE EN LA HOJA (Ya lo tenías, se mantiene igual)
' -------------------------------------------------------------------
Private Sub DibujarReporte(ws As Worksheet, dictCat As Object, dictSist As Object, dictMed As Object, titulo As String)
    Dim fila As Long: fila = 1
    
    ws.Cells(fila, 1).Value = "ESTADÍSTICAS DEL SERVICIO MÉDICO " & UCase(titulo)
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 3)).Merge
    ws.Cells(fila, 1).Font.Size = 14
    ws.Cells(fila, 1).Font.Bold = True
    ws.Cells(fila, 1).HorizontalAlignment = xlCenter
    fila = fila + 1
    
    ws.Cells(fila, 1).Value = "CATEGORÍA"
    ws.Cells(fila, 2).Value = "TIPO"
    ws.Cells(fila, 3).Value = "TOTAL"
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 3)).Interior.Color = RGB(200, 220, 240)
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 3)).Font.Bold = True
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 3)).Borders.LineStyle = xlContinuous
    fila = fila + 1

    ' --- BLOQUE 1: CATEGORIA GENERAL ---
    Dim filaIniCat As Long: filaIniCat = fila
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
    
    ws.Cells(fila, 2).Value = "TOTAL MES"
    ws.Cells(fila, 2).Font.Bold = True
    ws.Cells(fila, 3).Value = totalCat
    ws.Cells(fila, 3).Font.Bold = True
    ws.Cells(fila, 2).Borders.LineStyle = xlContinuous
    ws.Cells(fila, 3).Borders.LineStyle = xlContinuous
    
    ws.Range(ws.Cells(filaIniCat, 1), ws.Cells(fila, 1)).Merge
    ws.Cells(filaIniCat, 1).Value = "CATEGORÍA GENERAL"
    ws.Cells(filaIniCat, 1).HorizontalAlignment = xlCenter
    ws.Cells(filaIniCat, 1).VerticalAlignment = xlCenter
    ws.Cells(filaIniCat, 1).Font.Bold = True
    
    fila = fila + 2

    ' --- BLOQUE 2: SISTEMAS AFECTADOS ---
    Dim filaIniSist As Long: filaIniSist = fila
    Dim keysSist As Variant: keysSist = dictSist.keys
    Dim totalSist As Double: totalSist = 0
    Dim k As Variant
    
    If dictSist.count = 0 Then
        ws.Cells(fila, 2).Value = "SIN DATOS EN EL PERIODO"
    Else
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
    
    Dim medsFijos As Variant
    medsFijos = Array("DRA. ARBELI HERNANDEZ", "DRA. RAQUEL RODRIGUEZ", "DRA. ROSANA VARGAS", _
                      "DRA. DALITZA FERMIN", "DRA. MARKIS GARCIA", "DR. EUSEBEL CASTILLO", _
                      "DR. ISAAC MAR PEDIATRA", "DR. HUGO PADILLA CONSULTA")
                      
    For Each k In medsFijos
        ws.Cells(fila, 2).Value = k
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

    ws.Columns("A:C").AutoFit
    ws.Columns("B").ColumnWidth = 35
    ws.Columns("C").ColumnWidth = 15
End Sub

