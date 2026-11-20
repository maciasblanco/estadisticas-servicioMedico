VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstadisticas 
   Caption         =   "UserForm1"
   ClientHeight    =   8865.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14820
   OleObjectBlob   =   "frmEstadisticas.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Variable para guardar el título del periodo actual y usarlo al exportar
Private mTituloPeriodo As String

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    ' Llenar años
    For i = 2020 To Year(Date) + 1
        Me.cboAnio.AddItem i
    Next i
    
    ' Llenar meses
    For i = 1 To 12
        Me.cboMes.AddItem MonthName(i)
    Next i
    
    ' Llenar trimestres
    For i = 1 To 4
        Me.cboTrimestre.AddItem i
    Next i
    
    ' Llenar semestres
    For i = 1 To 2
        Me.cboSemestre.AddItem i
    Next i
    
    ' Centrar formulario
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    Me.btnExportar.Enabled = False ' Deshabilitado hasta que haya previa
End Sub

' -------------------------------------------------------------------
' FUNCIÓN AUXILIAR PARA LLENAR EL LISTBOX
' -------------------------------------------------------------------
Private Sub CargarPrevia(fIni As Date, fFin As Date, titulo As String)
    mTituloPeriodo = titulo
    Dim matriz As Variant
    matriz = GenerarMatrizPrevia(fIni, fFin)
    
    Me.lstPrevia.Clear
    
    ' ? VALIDACIÓN CRÍTICA AÑADIDA
    If Not IsArray(matriz) Then
        MsgBox "No se generaron datos para el período.", vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
    Dim uBoundMatriz As Long
    uBoundMatriz = UBound(matriz, 1)
    If Err.Number <> 0 Then
        MsgBox "La matriz no tiene la estructura esperada.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    Dim i As Long
    For i = 1 To uBoundMatriz
        Me.lstPrevia.AddItem matriz(i, 1) ' Asegúrate de que la columna 1 exista
        ' ... resto de tu código ...
    Next i
End Sub

' -------------------------------------------------------------------
' BOTONES DE GENERACIÓN
' -------------------------------------------------------------------
Private Sub btnMes_Click()
    If Me.cboAnio.Value = "" Or Me.cboMes.Value = "" Then
        MsgBox "Seleccione Año y Mes.", vbExclamation: Exit Sub
    End If
    Dim fIni As Date, fFin As Date
    RangoMes CInt(Me.cboAnio.Value), (Me.cboMes.ListIndex + 1), fIni, fFin
    CargarPrevia fIni, fFin, "AÑO " & Me.cboAnio.Value & " - " & UCase(Me.cboMes.Value)
End Sub

Private Sub btnTrimestre_Click()
    If Me.cboAnio.Value = "" Or Me.cboTrimestre.Value = "" Then
        MsgBox "Seleccione Año y Trimestre.", vbExclamation: Exit Sub
    End If
    Dim fIni As Date, fFin As Date
    RangoTrimestre CInt(Me.cboAnio.Value), CInt(Me.cboTrimestre.Value), fIni, fFin
    CargarPrevia fIni, fFin, "AÑO " & Me.cboAnio.Value & " - TRIMESTRE " & Me.cboTrimestre.Value
End Sub

Private Sub btnSemestre_Click()
    If Me.cboAnio.Value = "" Or Me.cboSemestre.Value = "" Then
        MsgBox "Seleccione Año y Semestre.", vbExclamation: Exit Sub
    End If
    Dim fIni As Date, fFin As Date
    RangoSemestre CInt(Me.cboAnio.Value), CInt(Me.cboSemestre.Value), fIni, fFin
    CargarPrevia fIni, fFin, "AÑO " & Me.cboAnio.Value & " - SEMESTRE " & Me.cboSemestre.Value
End Sub

Private Sub btnRango_Click()
    If Not IsDate(Me.txtDesde.Value) Or Not IsDate(Me.txtHasta.Value) Then
        MsgBox "Ingrese fechas válidas (dd/mm/aaaa).", vbExclamation: Exit Sub
    End If
    Dim fIni As Date, fFin As Date
    fIni = CDate(Me.txtDesde.Value)
    fFin = CDate(Me.txtHasta.Value)
    If fIni > fFin Then
        MsgBox "La fecha 'Desde' no puede ser mayor a 'Hasta'.", vbExclamation: Exit Sub
    End If
    CargarPrevia fIni, fFin, "DEL " & Format(fIni, "dd/mm/yyyy") & " AL " & Format(fFin, "dd/mm/yyyy")
End Sub

' -------------------------------------------------------------------
' BOTÓN PARA EXPORTAR A EXCEL
' -------------------------------------------------------------------
Private Sub btnExportar_Click()
    If mTituloPeriodo = "" Then
        MsgBox "No hay datos para exportar.", vbExclamation: Exit Sub
    End If
    ExportarReporteAExcel mTituloPeriodo ' Llama al módulo modEstadisticas
End Sub

Private Sub btnSalir_Click()
    Unload Me
End Sub

