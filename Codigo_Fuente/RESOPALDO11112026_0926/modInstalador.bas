Attribute VB_Name = "modInstalador"
Option Explicit

' -------------------------------------------------------------------
'  MÓDULO: modInstalador
'  Importa automáticamente todos los módulos y formularios
'  desde la carpeta Codigo_Fuente
' -------------------------------------------------------------------

Public Sub ImportarTodoElSistema()
    On Error GoTo Errores
    
    If MsgBox("¿Importar todos los módulos y formularios del sistema?" & vbCrLf & vbCrLf & _
              "Se leerán los archivos de la carpeta 'Codigo_Fuente'", _
              vbYesNo + vbQuestion, "Instalador Automático") = vbNo Then Exit Sub
    
    Dim rutaBase As String
    rutaBase = ThisWorkbook.Path & "\Codigo_Fuente\"
    
    ' Verificar que exista la carpeta
    If Dir(rutaBase, vbDirectory) = "" Then
        MsgBox "No se encontró la carpeta:" & vbCrLf & rutaBase & vbCrLf & vbCrLf & _
               "Créela junto al archivo .xlsm y coloque los archivos .bas y .frm.", _
               vbCritical, "Carpeta no encontrada"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    
    ' -- Lista de archivos a importar (en orden) -----------------
    Dim archivos As Variant
    archivos = Array( _
        "modVariables.bas", _
        "modSetup.bas", _
        "modSeguridad.bas", _
        "modHistorico.bas", _
        "modCRUD.bas", _
        "modValidaciones.bas", _
        "modEstadisticas.bas", _
        "frmLogin.frm", _
        "frmMenuPrincipal.frm", _
        "frmCRUD.frm", _
        "frmEstadisticas.frm", _
        "frmHistorico.frm" _
    )
    
    Dim log As String
    Dim i As Long
    Dim importados As Long: importados = 0
    Dim Errores As Long: Errores = 0
    
    For i = LBound(archivos) To UBound(archivos)
        Dim rutaArchivo As String
        rutaArchivo = rutaBase & archivos(i)
        
        If Dir(rutaArchivo) = "" Then
            log = log & "? No encontrado: " & archivos(i) & vbCrLf
            Errores = Errores + 1
        Else
            On Error Resume Next
            ' Eliminar si ya existe
            Dim nombreSinExt As String
            nombreSinExt = Left(archivos(i), InStrRev(archivos(i), ".") - 1)
            vbProj.VBComponents.Remove vbProj.VBComponents(nombreSinExt)
            Err.Clear
            On Error GoTo Errores
            
            ' Importar
            On Error Resume Next
            vbProj.VBComponents.Import rutaArchivo
            If Err.Number <> 0 Then
                log = log & "? Error: " & archivos(i) & " (" & Err.Description & ")" & vbCrLf
                Errores = Errores + 1
                Err.Clear
            Else
                log = log & "? " & archivos(i) & vbCrLf
                importados = importados + 1
            End If
            On Error GoTo Errores
        End If
    Next i
    
    ' -- Crear hojas de soporte ---------------------------------
    CrearHojaSiNoExiste "Usuarios", _
        Array("ID", "USUARIO", "NOMBRE_COMPLETO", "ROL", "HASH_CLAVE", "ACTIVO", "INTENTOS", "ULTIMO_ACCESO")
    
    CrearHojaSiNoExiste "Historico", _
        Array("FECHA_HORA", "USUARIO", "ACCION", "HOJA", "ID_REGISTRO", "CAMPO", _
              "VALOR_ANTERIOR", "VALOR_NUEVO", "DETALLE", "EQUIPO")
    
    CrearHojaSiNoExiste "Configuracion", _
        Array("PARAMETRO", "VALOR", "DESCRIPCION")
    
    ' -- Crear admin por defecto si no existe ------------------
    CrearAdminPorDefecto
    
    Application.ScreenUpdating = True
    
    ' -- Resultado ----------------------------------------------
    MsgBox "? Instalación completada" & vbCrLf & vbCrLf & _
           "Importados: " & importados & vbCrLf & _
           "Errores: " & Errores & vbCrLf & vbCrLf & _
           log & vbCrLf & _
           "????????????????????????????" & vbCrLf & _
           "Credenciales iniciales:" & vbCrLf & _
           "  Usuario: ADMIN" & vbCrLf & _
           "  Clave:   Admin2026!" & vbCrLf & vbCrLf & _
           "Cierre y vuelva a abrir el archivo.", _
           vbInformation, "Instalador"
    Exit Sub
    
Errores:
    Application.ScreenUpdating = True
    MsgBox "? Error: " & Err.Description & vbCrLf & "Número: " & Err.Number, vbCritical
End Sub

' -------------------------------------------------------------------
Private Sub CrearHojaSiNoExiste(ByVal nombre As String, ByVal encabezados As Variant)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombre)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = nombre
        Dim i As Long
        For i = LBound(encabezados) To UBound(encabezados)
            ws.Cells(1, i - LBound(encabezados) + 1).Value = encabezados(i)
        Next
        ws.Range("A1").Resize(1, UBound(encabezados) - LBound(encabezados) + 1).Font.Bold = True
        ws.Range("A1").Resize(1, UBound(encabezados) - LBound(encabezados) + 1).Interior.Color = RGB(31, 78, 120)
        ws.Range("A1").Resize(1, UBound(encabezados) - LBound(encabezados) + 1).Font.Color = RGB(255, 255, 255)
        ws.Columns.AutoFit
    End If
End Sub

' -------------------------------------------------------------------
Private Sub CrearAdminPorDefecto()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Usuarios")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    ' Verificar si ya hay usuarios
    If ws.Cells(ws.Rows.Count, 2).End(xlUp).Row >= 2 Then Exit Sub
    
    ws.Cells(2, 1).Value = 1
    ws.Cells(2, 2).Value = "ADMIN"
    ws.Cells(2, 3).Value = "Administrador del Sistema"
    ws.Cells(2, 4).Value = "ADMIN"
    ws.Cells(2, 5).Value = "INICIAL" ' Se reemplaza al primer login
    ws.Cells(2, 6).Value = "SI"
    ws.Cells(2, 7).Value = 0
End Sub

