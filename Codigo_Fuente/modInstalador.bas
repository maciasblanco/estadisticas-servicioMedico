Attribute VB_Name = "modInstalador"
'Attribute VB_Name = "modInstalador"
Option Explicit

' -------------------------------------------------------------------
'  MÓDULO: modInstalador
'  Importa automáticamente todos los módulos y formularios
'  desde la carpeta Codigo_Fuente y prepara las hojas de soporte.
' -------------------------------------------------------------------

Public Sub ImportarTodoElSistema()
    On Error GoTo Errores
    
    If MsgBox("¿Importar todos los módulos y formularios del sistema?" & vbCrLf & vbCrLf & _
              "Se leerán los archivos de la carpeta 'Codigo_Fuente'", _
              vbYesNo + vbQuestion, "Instalador Automático") = vbNo Then Exit Sub
    
    Dim rutaBase As String
    rutaBase = ThisWorkbook.Path & "\Codigo_Fuente\"
    
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
        "modDatos.bas", _
        "modCRUD.bas", _
        "modEstadisticas.bas", _
        "modExportacion.bas", _
        "frmLogin.frm", _
        "frmMenuPrincipal.frm", _
        "frmCRUD.frm", _
        "frmEstadisticas.frm" _
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
            Dim nombreSinExt As String
            nombreSinExt = Left(archivos(i), InStrRev(archivos(i), ".") - 1)
            vbProj.VBComponents.Remove vbProj.VBComponents(nombreSinExt)
            Err.Clear
            On Error GoTo Errores
            
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
    CrearHojaSiNoExiste HOJA_USUARIOS, _
        Array("ID", "USUARIO", "NOMBRE_COMPLETO", "ROL", "HASH_CLAVE", _
              "ACTIVO", "INTENTOS", "ULTIMO_ACCESO")
    
    CrearHojaSiNoExiste HOJA_HISTORICO, _
        Array("FECHA_HORA", "USUARIO", "ACCION", "HOJA", "ID_REGISTRO", _
              "CAMPO", "VALOR_ANTERIOR", "VALOR_NUEVO", "DETALLE", "EQUIPO")
    
    CrearHojaSiNoExiste HOJA_CONFIG, _
        Array("PARAMETRO", "VALOR", "DESCRIPCION")
    
    ' -- Catálogos (Sexo, Atención, Médico, Procedencia, Diagnóstico) --
    CrearCatalogosSiNoExiste
    
    ' -- Hoja original: debe existir o preguntarse si se crea vacía --
    If Not ExisteHoja(HOJA_DATOS_ORIGINAL) Then
        MsgBox "ADVERTENCIA:" & vbCrLf & vbCrLf & _
               "No existe la hoja '" & HOJA_DATOS_ORIGINAL & "'." & vbCrLf & vbCrLf & _
               "Debe crearla manualmente con sus datos de origen, o permitir" & vbCrLf & _
               "que el sistema ofrezca crearla vacía (solo si es nuevo año)." & vbCrLf & vbCrLf & _
               "El instalador continuará, pero el CRUD no funcionará" & vbCrLf & _
               "hasta que la hoja original exista.", _
               vbExclamation, "Hoja original faltante"
    Else
        ' La original existe: crear la copia de trabajo
        Dim nombreTrabajo As String
        nombreTrabajo = ObtenerHojaTrabajo()
        If nombreTrabajo <> "" Then
            log = log & "? Hoja de trabajo: " & nombreTrabajo & vbCrLf
        End If
    End If
    
    ' -- Crear admin por defecto si no existe ------------------
    CrearAdminPorDefecto
    
    Application.ScreenUpdating = True
    
    ' -- Resultado ----------------------------------------------
    MsgBox "Instalación completada" & vbCrLf & vbCrLf & _
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
    MsgBox "Error: " & Err.Description & vbCrLf & "Número: " & Err.Number, vbCritical
End Sub

' -------------------------------------------------------------------
' Crea una hoja con encabezados si no existe.
' -------------------------------------------------------------------
Private Sub CrearHojaSiNoExiste(ByVal nombre As String, ByVal encabezados As Variant)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombre)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = nombre
        Dim i As Long
        For i = LBound(encabezados) To UBound(encabezados)
            ws.Cells(1, i - LBound(encabezados) + 1).Value = encabezados(i)
        Next
        With ws.Range("A1").Resize(1, UBound(encabezados) - LBound(encabezados) + 1)
            .Font.Bold = True
            .Interior.Color = RGB(31, 78, 120)
            .Font.Color = RGB(255, 255, 255)
        End With
        ws.Columns.AutoFit
    End If
End Sub

' -------------------------------------------------------------------
' Crea las hojas de catálogo si no existen, con datos por defecto.
' -------------------------------------------------------------------
Private Sub CrearCatalogosSiNoExiste()
    Dim ws As Worksheet
    
    ' Datos-Sexo
    CrearHojaSiNoExiste HOJA_SEXO, Array("ID", "DESCRIPCION")
    Set ws = ThisWorkbook.Sheets(HOJA_SEXO)
    If ws.Cells(ws.Rows.count, 2).End(xlUp).Row = 1 Then
        ws.Cells(2, 1).Value = 1: ws.Cells(2, 2).Value = "MASCULINO"
        ws.Cells(3, 1).Value = 2: ws.Cells(3, 2).Value = "FEMENINO"
    End If
    
    ' Datos-TipoAtencion
    CrearHojaSiNoExiste HOJA_ATENCION, Array("ID", "DESCRIPCION")
    Set ws = ThisWorkbook.Sheets(HOJA_ATENCION)
    If ws.Cells(ws.Rows.count, 2).End(xlUp).Row = 1 Then
        ws.Cells(2, 1).Value = 1: ws.Cells(2, 2).Value = "CONSULTA"
        ws.Cells(3, 1).Value = 2: ws.Cells(3, 2).Value = "EMERGENCIA"
        ws.Cells(4, 1).Value = 3: ws.Cells(4, 2).Value = "CONTROL"
        ws.Cells(5, 1).Value = 4: ws.Cells(5, 2).Value = "REFERIDO"
    End If
    
    ' Datos-MedicoTratante
    CrearHojaSiNoExiste HOJA_MEDICO, Array("ID", "DESCRIPCION")
    Set ws = ThisWorkbook.Sheets(HOJA_MEDICO)
    If ws.Cells(ws.Rows.count, 2).End(xlUp).Row = 1 Then
        ws.Cells(2, 1).Value = 1: ws.Cells(2, 2).Value = "DRA. ARBELI HERNANDEZ"
        ws.Cells(3, 1).Value = 2: ws.Cells(3, 2).Value = "DRA. RAQUEL RODRIGUEZ"
        ws.Cells(4, 1).Value = 3: ws.Cells(4, 2).Value = "DRA. ROSANA VARGAS"
        ws.Cells(5, 1).Value = 4: ws.Cells(5, 2).Value = "DRA. DALITZA FERMIN"
        ws.Cells(6, 1).Value = 5: ws.Cells(6, 2).Value = "DRA. MARKIS GARCIA"
        ws.Cells(7, 1).Value = 6: ws.Cells(7, 2).Value = "DR. EUSEBEL CASTILLO"
        ws.Cells(8, 1).Value = 7: ws.Cells(8, 2).Value = "DR. ISAAC MAR PEDIATRA"
        ws.Cells(9, 1).Value = 8: ws.Cells(9, 2).Value = "DR. HUGO PADILLA CONSULTA"
    End If
    
    ' Datos-Procedencia
    CrearHojaSiNoExiste HOJA_PROCEDENCIA, Array("ID", "DESCRIPCION")
    Set ws = ThisWorkbook.Sheets(HOJA_PROCEDENCIA)
    If ws.Cells(ws.Rows.count, 2).End(xlUp).Row = 1 Then
        ws.Cells(2, 1).Value = 1: ws.Cells(2, 2).Value = "TRABAJADORES"
        ws.Cells(3, 1).Value = 2: ws.Cells(3, 2).Value = "FAMILIARES"
        ws.Cells(4, 1).Value = 3: ws.Cells(3, 2).Value = "PACIENTES"
        ws.Cells(5, 1).Value = 4: ws.Cells(4, 2).Value = "JUBILADOS"
    End If
    
    ' Datos-Diagnostico
    CrearHojaSiNoExiste HOJA_DIAGNOSTICO, Array("ID", "DESCRIPCION")
    Set ws = ThisWorkbook.Sheets(HOJA_DIAGNOSTICO)
    If ws.Cells(ws.Rows.count, 2).End(xlUp).Row = 1 Then
        ws.Cells(2, 1).Value = 1: ws.Cells(2, 2).Value = "RESFRIADO COMUN"
        ws.Cells(3, 1).Value = 2: ws.Cells(3, 2).Value = "GASTRITIS"
        ws.Cells(4, 1).Value = 3: ws.Cells(4, 2).Value = "HIPERTENSION"
        ws.Cells(5, 1).Value = 4: ws.Cells(5, 2).Value = "DIABETES"
    End If
End Sub

' -------------------------------------------------------------------
' Crea el usuario ADMIN si la hoja Usuarios está vacía.
' -------------------------------------------------------------------
Private Sub CrearAdminPorDefecto()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(HOJA_USUARIOS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    If ws.Cells(ws.Rows.count, 2).End(xlUp).Row >= 2 Then Exit Sub
    
    ws.Cells(2, UCOL_ID).Value = 1
    ws.Cells(2, UCOL_USUARIO).Value = "ADMIN"
    ws.Cells(2, UCOL_NOMBRE).Value = "Administrador del Sistema"
    ws.Cells(2, UCOL_ROL).Value = ROL_ADMIN
    ws.Cells(2, UCOL_HASH).Value = "INICIAL"
    ws.Cells(2, UCOL_ACTIVO).Value = "SI"
    ws.Cells(2, UCOL_INTENTOS).Value = 0
End Sub

