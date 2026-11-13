Attribute VB_Name = "modSetup"
Option Explicit

' ═══════════════════════════════════════════════════════════════════
'  MÓDULO: modSetup
'  Inicializa hojas de soporte y datos por defecto
' ═══════════════════════════════════════════════════════════════════

Public Sub InicializarSistema()
    On Error GoTo Errores
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    CrearHojaUsuarios
    CrearHojaHistorico
    CrearHojaConfiguracion
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Sistema inicializado correctamente." & vbCrLf & vbCrLf & _
           "Usuario por defecto:" & vbCrLf & _
           "  Usuario: ADMIN" & vbCrLf & _
           "  Clave:   Admin2026!" & vbCrLf & vbCrLf & _
           "Cambie la clave al primer inicio de sesión.", _
           vbInformation, SISTEMA_NOMBRE
    Exit Sub
Errores:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Error al inicializar:" & vbCrLf & Err.Description, vbCritical
End Sub

' ───────────────────────────────────────────────────────────────────
Private Sub CrearHojaUsuarios()
    Dim ws As Worksheet
    If ExisteHoja(HOJA_USUARIOS) Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = HOJA_USUARIOS
    
    With ws
        .Range("A1:H1").Value = Array("ID", "USUARIO", "NOMBRE_COMPLETO", "ROL", _
                                      "HASH_CLAVE", "ACTIVO", "INTENTOS", "ULTIMO_ACCESO")
        .Range("A1:H1").Font.Bold = True
        .Range("A1:H1").Interior.Color = COLOR_PRIMARIO
        .Range("A1:H1").Font.Color = COLOR_TEXTO_CLARO
        .Range("A1:H1").HorizontalAlignment = xlCenter
        
        .Cells(2, UCOL_ID).Value = 1
        .Cells(2, UCOL_USUARIO).Value = "ADMIN"
        .Cells(2, UCOL_NOMBRE).Value = "Administrador del Sistema"
        .Cells(2, UCOL_ROL).Value = ROL_ADMIN
        .Cells(2, UCOL_HASH).Value = "INICIAL"
        .Cells(2, UCOL_ACTIVO).Value = "SI"
        .Cells(2, UCOL_INTENTOS).Value = 0
        
        .Columns("A:H").AutoFit
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 30
        .Columns("E").ColumnWidth = 70
        .Rows("1:1").RowHeight = 22
    End With
End Sub

' ───────────────────────────────────────────────────────────────────
Private Sub CrearHojaHistorico()
    Dim ws As Worksheet
    If ExisteHoja(HOJA_HISTORICO) Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = HOJA_HISTORICO
    
    With ws
        .Range("A1:J1").Value = Array("FECHA_HORA", "USUARIO", "ACCION", "HOJA", _
                                       "ID_REGISTRO", "CAMPO", "VALOR_ANTERIOR", _
                                       "VALOR_NUEVO", "DETALLE", "EQUIPO")
        .Range("A1:J1").Font.Bold = True
        .Range("A1:J1").Interior.Color = COLOR_PRIMARIO
        .Range("A1:J1").Font.Color = COLOR_TEXTO_CLARO
        .Range("A1:J1").HorizontalAlignment = xlCenter
        
        .Columns("A:A").ColumnWidth = 20
        .Columns("B:B").ColumnWidth = 15
        .Columns("C:C").ColumnWidth = 14
        .Columns("D:D").ColumnWidth = 22
        .Columns("E:J").ColumnWidth = 25
        .Rows("1:1").RowHeight = 22
    End With
End Sub

' ───────────────────────────────────────────────────────────────────
Private Sub CrearHojaConfiguracion()
    Dim ws As Worksheet
    If ExisteHoja(HOJA_CONFIG) Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = HOJA_CONFIG
    
    With ws
        .Range("A1:C1").Value = Array("PARAMETRO", "VALOR", "DESCRIPCION")
        .Range("A1:C1").Font.Bold = True
        .Range("A1:C1").Interior.Color = COLOR_PRIMARIO
        .Range("A1:C1").Font.Color = COLOR_TEXTO_CLARO
        .Range("A1:C1").HorizontalAlignment = xlCenter
        
        Dim fila As Long: fila = 2
        .Cells(fila, 1).Value = "VERSION": .Cells(fila, 2).Value = SISTEMA_VERSION: .Cells(fila, 3).Value = "Versión del sistema": fila = fila + 1
        .Cells(fila, 1).Value = "MAX_INTENTOS": .Cells(fila, 2).Value = MAX_INTENTOS: .Cells(fila, 3).Value = "Intentos de login": fila = fila + 1
        .Cells(fila, 1).Value = "MIN_LONG_CLAVE": .Cells(fila, 2).Value = MIN_LONG_CLAVE: .Cells(fila, 3).Value = "Longitud mínima de clave": fila = fila + 1
        .Cells(fila, 1).Value = "RUTA_RESPALDO": .Cells(fila, 2).Value = CARPETA_RESPALDO: .Cells(fila, 3).Value = "Carpeta de respaldos": fila = fila + 1
        .Cells(fila, 1).Value = "RUTA_EXPORT": .Cells(fila, 2).Value = CARPETA_EXPORT: .Cells(fila, 3).Value = "Carpeta de exportaciones"
        
        .Columns("A:C").AutoFit
        .Columns("C").ColumnWidth = 40
    End With
End Sub
