Attribute VB_Name = "modExportacion"
'Attribute VB_Name = "modExportacion"
Option Explicit

' -------------------------------------------------------------------
'  MÓDULO: modExportacion
'  Exportación de datos a Excel y PDF con validación de permisos.
'  Todas las exportaciones trabajan sobre la hoja de trabajo
'  "CRUD Datos desde ... hasta ..." (nunca sobre DatosOriginales).
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' 1. VERIFICACIÓN DE PERMISO
' -------------------------------------------------------------------
Public Function VerificarPermisoExport() As Boolean
    VerificarPermisoExport = False
    If Not TienePermiso("EXPORT") Then
        MsgBox "No tiene permiso para exportar datos." & vbCrLf & vbCrLf & _
               "Solo los roles ADMIN y ESTADISTICA pueden exportar.", _
               vbExclamation, "Permiso denegado"
        RegistrarHistorico ACC_EXPORT, "SISTEMA", "", "", gRolUsuario, "", _
                           "Intento de exportación sin permiso"
        Exit Function
    End If
    VerificarPermisoExport = True
End Function

' -------------------------------------------------------------------
' 2. GESTIÓN DE CARPETAS
' -------------------------------------------------------------------
Public Function AsegurarCarpeta(ByVal ruta As String) As Boolean
    AsegurarCarpeta = False
    On Error GoTo Errores
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(ruta) Then
        fso.CreateFolder ruta
    End If
    
    AsegurarCarpeta = True
    Exit Function
Errores:
    MsgBox "No se pudo crear/verificar la carpeta:" & vbCrLf & ruta & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Error de carpeta"
End Function

Public Sub AbrirCarpeta(ByVal ruta As String)
    On Error Resume Next
    Shell "explorer.exe """ & ruta & """", vbNormalFocus
End Sub

' -------------------------------------------------------------------
' 3. NOMBRE DE ARCHIVO ÚNICO
' -------------------------------------------------------------------
Public Function GenerarNombreArchivo(ByVal prefijo As String, _
                                     ByVal extension As String) As String
    Dim sPref As String
    sPref = SanitizarNombre(prefijo)
    
    Dim sFecha As String
    sFecha = Format(Now, "yyyymmdd_hhmmss")
    
    Dim ext As String
    ext = LCase$(Replace(extension, ".", ""))
    
    GenerarNombreArchivo = sPref & "_" & sFecha & "." & ext
End Function

Private Function SanitizarNombre(ByVal txt As String) As String
    Dim i As Long, c As String, res As String
    res = ""
    For i = 1 To Len(txt)
        c = Mid$(txt, i, 1)
        If (c >= "a" And c <= "z") Or (c >= "A" And c <= "Z") Or _
           (c >= "0" And c <= "9") Or c = "_" Or c = "-" Then
            res = res & c
        ElseIf c = " " Then
            res = res & "_"
        End If
    Next
    If Len(res) = 0 Then res = "Reporte"
    SanitizarNombre = res
End Function

' -------------------------------------------------------------------
' 4. SELECTOR DE RUTA (GetSaveAsFilename)
' -------------------------------------------------------------------
Public Function SeleccionarRutaGuardar(ByVal filtro As String, _
                                       ByVal nombreDefecto As String) As String
    Dim ruta As Variant
    ruta = Application.GetSaveAsFilename( _
        InitialFileName:=nombreDefecto, _
        FileFilter:=filtro, _
        Title:="Guardar reporte como...")
    
    If ruta = False Then
        SeleccionarRutaGuardar = ""
    Else
        SeleccionarRutaGuardar = CStr(ruta)
    End If
End Function

' -------------------------------------------------------------------
' 5. EXPORTAR RANGO A EXCEL (.xlsx)
' -------------------------------------------------------------------
Public Function ExportarRangoAExcel(ByRef rango As Range, _
                                    ByVal rutaDestino As String, _
                                    Optional ByVal tituloHoja As String = "Reporte") As Boolean
    ExportarRangoAExcel = False
    
    If Not VerificarPermisoExport() Then Exit Function
    If rutaDestino = "" Then Exit Function
    If rango Is Nothing Then
        MsgBox "No hay datos para exportar.", vbExclamation, "Sin datos"
        Exit Function
    End If
    
    Dim wbNuevo As Workbook
    Dim wsNueva As Worksheet
    Dim rutaFinal As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error GoTo Errores
    
    rutaFinal = rutaDestino
    If LCase$(Right$(rutaFinal, 5)) <> ".xlsx" Then
        rutaFinal = rutaFinal & ".xlsx"
    End If
    
    Set wbNuevo = Workbooks.Add
    Set wsNueva = wbNuevo.Sheets(1)
    
    Dim nomHoja As String
    nomHoja = Left$(SanitizarNombre(tituloHoja), 31)
    If Len(nomHoja) = 0 Then nomHoja = "Reporte"
    On Error Resume Next
    wsNueva.Name = nomHoja
    Err.Clear
    On Error GoTo Errores
    
    rango.Copy Destination:=wsNueva.Range("A1")
    wsNueva.Columns.AutoFit
    wsNueva.Rows("1:1").Font.Bold = True
    
    wbNuevo.SaveAs Filename:=rutaFinal, FileFormat:=xlOpenXMLWorkbook
    wbNuevo.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    RegistrarHistorico ACC_EXPORT, "EXPORT", "", "EXCEL", rutaFinal, "", _
                       "Exportación a Excel: " & rango.Address
    ExportarRangoAExcel = True
    Exit Function
    
Errores:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Not wbNuevo Is Nothing Then wbNuevo.Close SaveChanges:=False
    MsgBox "Error al exportar a Excel:" & vbCrLf & Err.Description, vbCritical
End Function

' -------------------------------------------------------------------
' 6. EXPORTAR HOJA COMPLETA A EXCEL
' -------------------------------------------------------------------
Public Function ExportarHojaAExcel(ByVal nombreHojaOrigen As String, _
                                   ByVal rutaDestino As String) As Boolean
    ExportarHojaAExcel = False
    
    If Not VerificarPermisoExport() Then Exit Function
    If rutaDestino = "" Then Exit Function
    If Not ExisteHoja(nombreHojaOrigen) Then
        MsgBox "No existe la hoja: " & nombreHojaOrigen, vbExclamation, "Error"
        Exit Function
    End If
    
    Dim wsOrigen As Worksheet
    Dim rangoUsado As Range
    Set wsOrigen = ThisWorkbook.Sheets(nombreHojaOrigen)
    Set rangoUsado = wsOrigen.UsedRange
    
    If rangoUsado Is Nothing Or rangoUsado.Cells.count = 0 Then
        MsgBox "La hoja está vacía.", vbExclamation, "Sin datos"
        Exit Function
    End If
    
    ExportarHojaAExcel = ExportarRangoAExcel(rangoUsado, rutaDestino, _
                                              Left$(nombreHojaOrigen, 30))
End Function

' -------------------------------------------------------------------
' 7. EXPORTAR A PDF
' -------------------------------------------------------------------
Public Function ExportarRangoAPDF(ByRef rango As Range, _
                                  ByVal rutaDestino As String, _
                                  Optional ByVal tituloReporte As String = "") As Boolean
    ExportarRangoAPDF = False
    
    If Not VerificarPermisoExport() Then Exit Function
    If rutaDestino = "" Then Exit Function
    If rango Is Nothing Then Exit Function
    
    Dim rutaFinal As String
    rutaFinal = rutaDestino
    If LCase$(Right$(rutaFinal, 4)) <> ".pdf" Then
        rutaFinal = rutaFinal & ".pdf"
    End If
    
    Dim wbTmp As Workbook
    Dim wsTmp As Worksheet
    Dim titulo As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error GoTo Errores
    
    Set wbTmp = Workbooks.Add
    Set wsTmp = wbTmp.Sheets(1)
    
    titulo = tituloReporte
    If Len(titulo) = 0 Then titulo = "Reporte " & Format(Now, "yyyy-mm-dd hh:mm")
    
    With wsTmp.Range("A1")
        .Value = titulo
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
    wsTmp.Range("A1").EntireRow.RowHeight = 24
    
    wsTmp.Range("A2").Value = "Generado: " & Format(Now, FORMATO_DATETIME) & _
                               "   |   Usuario: " & gUsuarioActual
    wsTmp.Range("A2").Font.Italic = True
    wsTmp.Range("A2").Font.Size = 9
    
    rango.Copy Destination:=wsTmp.Range("A4")
    
    wsTmp.Columns.AutoFit
    wsTmp.PageSetup.Orientation = xlLandscape
    wsTmp.PageSetup.Zoom = False
    wsTmp.PageSetup.FitToPagesWide = 1
    wsTmp.PageSetup.FitToPagesTall = False
    
    wsTmp.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=rutaFinal, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    
    wbTmp.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    RegistrarHistorico ACC_EXPORT, "EXPORT", "", "PDF", rutaFinal, "", _
                       "Exportación a PDF: " & rango.Address
    ExportarRangoAPDF = True
    Exit Function
    
Errores:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Not wbTmp Is Nothing Then wbTmp.Close SaveChanges:=False
    MsgBox "Error al exportar a PDF:" & vbCrLf & Err.Description & vbCrLf & _
           "Número: " & Err.Number, vbCritical
End Function

Public Function ExportarHojaAPDF(ByVal nombreHojaOrigen As String, _
                                 ByVal rutaDestino As String, _
                                 Optional ByVal tituloReporte As String = "") As Boolean
    ExportarHojaAPDF = False
    If Not VerificarPermisoExport() Then Exit Function
    If Not ExisteHoja(nombreHojaOrigen) Then Exit Function
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nombreHojaOrigen)
    Dim rng As Range
    Set rng = ws.UsedRange
    
    If rng Is Nothing Or rng.Cells.count = 0 Then
        MsgBox "La hoja está vacía.", vbExclamation, "Sin datos"
        Exit Function
    End If
    
    ExportarHojaAPDF = ExportarRangoAPDF(rng, rutaDestino, tituloReporte)
End Function

' -------------------------------------------------------------------
' 8. EXPORTAR DATOS DE LA HOJA DE TRABAJO (copia CRUD)
' -------------------------------------------------------------------
Public Function ExportarDatosTrabajo(ByVal rutaDestino As String, _
                                     ByVal formato As String) As Boolean
    ExportarDatosTrabajo = False
    If Not VerificarPermisoExport() Then Exit Function
    
    Dim nombreHoja As String
    nombreHoja = ObtenerHojaTrabajo()
    If nombreHoja = "" Then Exit Function
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(nombreHoja)
    Dim rng As Range: Set rng = ws.UsedRange
    
    If UCase$(formato) = "PDF" Then
        ExportarDatosTrabajo = ExportarRangoAPDF(rng, rutaDestino, _
            "Reporte de Servicios Médicos")
    Else
        ExportarDatosTrabajo = ExportarRangoAExcel(rng, rutaDestino, "Datos")
    End If
End Function

' -------------------------------------------------------------------
' 9. EXPORTAR RESULTADOS DE BÚSQUEDA (para frmCRUD)
' -------------------------------------------------------------------
Public Function ExportarResultadosBusqueda(ByVal rutaDestino As String, _
                                           ByVal textoBusqueda As String) As Boolean
    ExportarResultadosBusqueda = False
    If Not VerificarPermisoExport() Then Exit Function
    
    Dim nombreHoja As String
    nombreHoja = ObtenerHojaTrabajo()
    If nombreHoja = "" Then Exit Function
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(nombreHoja)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, COL_CEDULA).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No hay datos para exportar.", vbExclamation
        Exit Function
    End If
    
    ' Crear hoja temporal
    Dim wsTmp As Worksheet
    Set wsTmp = ThisWorkbook.Sheets.Add
    wsTmp.Name = "TMP_EXPORT"
    
    Dim criterio As String: criterio = UCase$(Trim$(textoBusqueda))
    Dim filaTmp As Long: filaTmp = 1
    Dim i As Long, j As Long
    
    ' Encabezados
    For j = 1 To 10
        wsTmp.Cells(filaTmp, j).Value = ws.Cells(FILA_ENCABEZADO, j).Value
    Next j
    wsTmp.Range(wsTmp.Cells(1, 1), wsTmp.Cells(1, 10)).Font.Bold = True
    wsTmp.Range(wsTmp.Cells(1, 1), wsTmp.Cells(1, 10)).Interior.Color = COLOR_PRIMARIO
    wsTmp.Range(wsTmp.Cells(1, 1), wsTmp.Cells(1, 10)).Font.Color = COLOR_TEXTO_CLARO
    
    ' Filas que coincidan
    For i = 2 To lastRow
        Dim textoFila As String
        textoFila = UCase$(CStr(ws.Cells(i, COL_NOMBRE).Value) & "|" & _
                           CStr(ws.Cells(i, COL_CEDULA).Value))
        If Len(criterio) = 0 Or InStr(1, textoFila, criterio, vbTextCompare) > 0 Then
            filaTmp = filaTmp + 1
            For j = 1 To 10
                wsTmp.Cells(filaTmp, j).Value = ws.Cells(i, j).Value
            Next j
        End If
    Next i
    
    If filaTmp = 1 Then
        Application.DisplayAlerts = False
        wsTmp.Delete
        Application.DisplayAlerts = True
        MsgBox "No hay resultados que exportar.", vbExclamation
        Exit Function
    End If
    
    Dim rng As Range
    Set rng = wsTmp.Range(wsTmp.Cells(1, 1), wsTmp.Cells(filaTmp, 10))
    
    Dim resultado As Boolean
    resultado = ExportarRangoAExcel(rng, rutaDestino, "ResultadosBusqueda")
    
    Application.DisplayAlerts = False
    wsTmp.Delete
    Application.DisplayAlerts = True
    
    ExportarResultadosBusqueda = resultado
End Function

' -------------------------------------------------------------------
' 10. EXPORTAR REPORTE ESTADÍSTICO (para frmEstadisticas)
' -------------------------------------------------------------------
Public Function ExportarReporteAExcel(ByVal tituloPeriodo As String) As Boolean
    ExportarReporteAExcel = False
    If Not VerificarPermisoExport() Then Exit Function
    If IsEmpty(gMatrizReporte) Then
        MsgBox "No hay datos previos para exportar. Genere una previa primero.", _
               vbExclamation, "Sin datos"
        Exit Function
    End If
    
    If Not AsegurarCarpeta(CARPETA_EXPORT) Then Exit Function
    
    Dim ruta As String
    ruta = SeleccionarRutaGuardar("Libro de Excel (*.xlsx), *.xlsx", _
        CARPETA_EXPORT & GenerarNombreArchivo("Reporte_" & tituloPeriodo, "xlsx"))
    If ruta = "" Then Exit Function
    
    ' Crear hoja temporal con la matriz
    Dim wsTmp As Worksheet
    Set wsTmp = ThisWorkbook.Sheets.Add
    wsTmp.Name = "TMP_REPORTE"
    
    Dim i As Long
    wsTmp.Cells(1, 1).Value = "REPORTE: " & tituloPeriodo
    wsTmp.Range("A1:C1").Merge
    wsTmp.Cells(1, 1).Font.Bold = True
    wsTmp.Cells(1, 1).Font.Size = 14
    
    wsTmp.Cells(2, 1).Value = "TIPO"
    wsTmp.Cells(2, 2).Value = "NOMBRE"
    wsTmp.Cells(2, 3).Value = "TOTAL"
    wsTmp.Range("A2:C2").Font.Bold = True
    wsTmp.Range("A2:C2").Interior.Color = RGB(31, 78, 120)
    wsTmp.Range("A2:C2").Font.Color = RGB(255, 255, 255)
    
    For i = 1 To UBound(gMatrizReporte, 1)
        wsTmp.Cells(i + 2, 1).Value = gMatrizReporte(i, 1)
        wsTmp.Cells(i + 2, 2).Value = gMatrizReporte(i, 2)
        wsTmp.Cells(i + 2, 3).Value = gMatrizReporte(i, 3)
    Next i
    
    wsTmp.Columns("A:C").AutoFit
    
    Dim rng As Range
    Set rng = wsTmp.Range(wsTmp.Cells(1, 1), wsTmp.Cells(UBound(gMatrizReporte, 1) + 2, 3))
    
    ExportarReporteAExcel = ExportarRangoAExcel(rng, ruta, "Reporte")
    
    Application.DisplayAlerts = False
    wsTmp.Delete
    Application.DisplayAlerts = True
End Function

