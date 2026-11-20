VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCRUD 
   Caption         =   "UserForm1"
   ClientHeight    =   9300.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14880
   OleObjectBlob   =   "frmCRUD.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmCrud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Version 5#
'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCRUD
'   Caption = "UserForm1"
'   ClientHeight = 9300.001
'   ClientLeft = 120
'   ClientTop = 465
'   ClientWidth = 14880
'   OleObjectBlob   =   "frmCRUD.frx":0000
'   StartUpPosition = 1    'Centrar en propietario
'End
'Attribute VB_Name = "frmCrud"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
Option Explicit

Private mFilaSeleccionada As Long

' -------------------------------------------------------------------
' INICIALIZACIÓN
' -------------------------------------------------------------------
Private Sub UserForm_Initialize()
    ' Centrar formulario
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    ' Configurar ListBox
    Me.lstResultados.ColumnCount = 5
    Me.lstResultados.ColumnWidths = "60;120;80;120;100"
    
    ' Cargar combos y fecha por defecto
    CargarCombos
    Me.txtFecha.Value = Format(Date, FORMATO_FECHA)
End Sub

' -------------------------------------------------------------------
' CARGA DE COMBOS DESDE HOJAS DE CATÁLOGO
' -------------------------------------------------------------------
Private Sub CargarCombos()
    CargarComboDesdeHoja Me.cboSexo, HOJA_SEXO
    CargarComboDesdeHoja Me.cboAtencion, HOJA_ATENCION
    CargarComboDesdeHoja Me.cboMedico, HOJA_MEDICO
    CargarComboDesdeHoja Me.cboDiagnostico, HOJA_DIAGNOSTICO
End Sub

Private Sub CargarComboDesdeHoja(ByRef cbo As MSForms.ComboBox, ByVal nombreHoja As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
    Dim i As Long
    cbo.Clear
    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, 2).Value)) <> "" Then
            cbo.AddItem CStr(ws.Cells(i, 2).Value)
        End If
    Next i
End Sub

' -------------------------------------------------------------------
' BUSCAR
' -------------------------------------------------------------------
Private Sub btnBuscar_Click()
    Dim resultados() As Variant
    Dim total As Long
    
    Me.lstResultados.Clear
    mFilaSeleccionada = 0
    
    total = BuscarRegistros(Me.txtBuscar.Value, resultados)
    
    If total = 0 Then
        MsgBox MSG_SIN_RESULTADOS, vbInformation, "Sin resultados"
        Exit Sub
    End If
    
    Dim i As Long
    For i = 1 To total
        With Me.lstResultados
            .AddItem
            .List(.ListCount - 1, 0) = CStr(resultados(i, 3)) ' Cédula
            .List(.ListCount - 1, 1) = CStr(resultados(i, 2)) ' Nombre
            .List(.ListCount - 1, 2) = CStr(resultados(i, 1)) ' Fecha
            .List(.ListCount - 1, 3) = CStr(resultados(i, 10)) ' Diagnóstico
            .List(.ListCount - 1, 4) = CStr(resultados(i, 7)) ' Médico
        End With
    Next i
End Sub

' -------------------------------------------------------------------
' SELECCIÓN EN EL LISTBOX
' -------------------------------------------------------------------
Private Sub lstResultados_Click()
    If Me.lstResultados.ListIndex = -1 Then Exit Sub
    
    Dim cedula As String
    cedula = Me.lstResultados.List(Me.lstResultados.ListIndex, 0)
    
    Dim nombreHoja As String
    nombreHoja = ObtenerHojaTrabajo()
    If nombreHoja = "" Then Exit Sub
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(nombreHoja)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, COL_CEDULA).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    Dim i As Long
    For i = 2 To lastRow
        If CStr(ws.Cells(i, COL_CEDULA).Value) = cedula Then
            mFilaSeleccionada = i
            
            Me.txtFecha.Value = Format(ws.Cells(i, COL_FECHA).Value, FORMATO_FECHA)
            Me.txtNombre.Value = CStr(ws.Cells(i, COL_NOMBRE).Value)
            Me.txtCedula.Value = CStr(ws.Cells(i, COL_CEDULA).Value)
            Me.cboSexo.Value = CStr(ws.Cells(i, COL_SEXO).Value)
            Me.txtEdad.Value = CStr(ws.Cells(i, COL_EDAD).Value)
            Me.cboAtencion.Value = CStr(ws.Cells(i, COL_ATENCION).Value)
            Me.cboMedico.Value = CStr(ws.Cells(i, COL_MEDICO).Value)
            Me.txtProcedencia.Value = CStr(ws.Cells(i, COL_PROCEDENCIA).Value)
            Me.cboDiagnostico.Value = CStr(ws.Cells(i, COL_DIAGNOSTICO).Value)
            Exit Sub
        End If
    Next i
End Sub

' -------------------------------------------------------------------
' GUARDAR (INSERTAR)
' -------------------------------------------------------------------
Private Sub btnGuardar_Click()
    If Not ValidarCampos Then Exit Sub
    
    Dim valores(1 To 10) As Variant
    valores(1) = CDate(Me.txtFecha.Value)
    valores(2) = UCase(Trim(Me.txtNombre.Value))
    valores(3) = Trim(Me.txtCedula.Value)
    valores(4) = Me.cboSexo.Value
    valores(5) = Val(Me.txtEdad.Value)
    valores(6) = Me.cboAtencion.Value
    valores(7) = Me.cboMedico.Value
    valores(8) = ""
    valores(9) = Me.txtProcedencia.Value
    valores(10) = Me.cboDiagnostico.Value
    
    If InsertarRegistro(valores) Then
        MsgBox MSG_GUARDADO_OK, vbInformation, "Éxito"
        btnLimpiar_Click
        btnBuscar_Click
    End If
End Sub

' -------------------------------------------------------------------
' ACTUALIZAR
' -------------------------------------------------------------------
Private Sub btnActualizar_Click()
    If mFilaSeleccionada = 0 Then
        MsgBox "Seleccione un registro de la lista para actualizar.", vbExclamation, "Atención"
        Exit Sub
    End If
    If Not ValidarCampos Then Exit Sub
    
    Dim valores(1 To 10) As Variant
    valores(1) = CDate(Me.txtFecha.Value)
    valores(2) = UCase(Trim(Me.txtNombre.Value))
    valores(3) = Trim(Me.txtCedula.Value)
    valores(4) = Me.cboSexo.Value
    valores(5) = Val(Me.txtEdad.Value)
    valores(6) = Me.cboAtencion.Value
    valores(7) = Me.cboMedico.Value
    valores(8) = ""
    valores(9) = Me.txtProcedencia.Value
    valores(10) = Me.cboDiagnostico.Value
    
    If ActualizarRegistro(mFilaSeleccionada, valores) Then
        MsgBox MSG_ACTUALIZADO_OK, vbInformation, "Éxito"
        btnLimpiar_Click
        btnBuscar_Click
    End If
End Sub

' -------------------------------------------------------------------
' ELIMINAR
' -------------------------------------------------------------------
Private Sub btnEliminar_Click()
    If mFilaSeleccionada = 0 Then
        MsgBox "Seleccione un registro de la lista para eliminar.", vbExclamation, "Atención"
        Exit Sub
    End If
    
    If MsgBox(MSG_CONFIRMAR_DEL, vbYesNo + vbQuestion, "Confirmar") = vbYes Then
        If EliminarRegistro(mFilaSeleccionada) Then
            MsgBox MSG_ELIMINADO_OK, vbInformation, "Éxito"
            btnLimpiar_Click
            btnBuscar_Click
        End If
    End If
End Sub

' -------------------------------------------------------------------
' LIMPIAR
' -------------------------------------------------------------------
Private Sub btnLimpiar_Click()
    mFilaSeleccionada = 0
    Me.txtFecha.Value = Format(Date, FORMATO_FECHA)
    Me.txtNombre.Value = ""
    Me.txtCedula.Value = ""
    Me.cboSexo.ListIndex = -1
    Me.txtEdad.Value = ""
    Me.cboAtencion.ListIndex = -1
    Me.cboMedico.ListIndex = -1
    Me.txtProcedencia.Value = ""
    Me.cboDiagnostico.ListIndex = -1
    
    Me.txtBuscar.Value = ""
    Me.lstResultados.Clear
    Me.txtNombre.SetFocus
End Sub

' -------------------------------------------------------------------
' VALIDAR CAMPOS OBLIGATORIOS
' -------------------------------------------------------------------
Private Function ValidarCampos() As Boolean
    ValidarCampos = False
    
    If Trim(Me.txtNombre.Value) = "" Then
        MsgBox "El nombre del paciente es obligatorio.", vbExclamation, "Atención"
        Me.txtNombre.SetFocus
        Exit Function
    End If
    
    If Trim(Me.txtCedula.Value) = "" Then
        MsgBox "La cédula del paciente es obligatoria.", vbExclamation, "Atención"
        Me.txtCedula.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(Me.txtEdad.Value) Then
        MsgBox "La edad debe ser un valor numérico.", vbExclamation, "Atención"
        Me.txtEdad.SetFocus
        Exit Function
    End If
    
    ValidarCampos = True
End Function

' -------------------------------------------------------------------
' SALIR
' -------------------------------------------------------------------
Private Sub btnSalir_Click()
    Unload Me
End Sub

' -------------------------------------------------------------------
' EXPORTAR RESULTADOS
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

' -------------------------------------------------------------------
' Helper para exportar un .xlsx ya generado a PDF
' -------------------------------------------------------------------
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

