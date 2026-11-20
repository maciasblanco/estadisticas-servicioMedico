Attribute VB_Name = "modSetup"
'Attribute VB_Name = "modSetup"
Option Explicit

' -------------------------------------------------------------------
'  MÓDULO: modSetup
'  Inicializa hojas de soporte, datos por defecto y el entorno
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' INICIALIZAR EL ENTORNO DE TRABAJO (Hojas y datos base)
' -------------------------------------------------------------------
Public Sub InicializarEntornoDeTrabajo()
    On Error GoTo Errores
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 1. Crear hojas de soporte si no existen
    CrearHojaSiNoExiste HOJA_USUARIOS, _
        Array("ID", "USUARIO", "NOMBRE_COMPLETO", "ROL", "HASH_CLAVE", "ACTIVO", "INTENTOS", "ULTIMO_ACCESO")
    
    CrearHojaSiNoExiste HOJA_HISTORICO, _
        Array("FECHA_HORA", "USUARIO", "ACCION", "HOJA", "ID_REGISTRO", "CAMPO", "VALOR_ANTERIOR", "VALOR_NUEVO", "DETALLE", "EQUIPO")
    
    CrearHojaSiNoExiste HOJA_CONFIG, _
        Array("PARAMETRO", "VALOR", "DESCRIPCION")
        
    ' 2. Crear catálogos con datos por defecto
    CrearCatalogosBase
    
    ' 3. Crear el usuario ADMIN por defecto
    CrearAdminPorDefecto
    
    ' 4. Crear la hoja de trabajo (CRUD Datos desde... hasta...)
    If ExisteHoja(HOJA_DATOS_ORIGINAL) Then
        Dim nombreCopia As String
        nombreCopia = ObtenerHojaTrabajo() ' Llama a modDatos
        If nombreCopia <> "" Then
            Debug.Print "Hoja de trabajo lista: " & nombreCopia
        End If
    Else
        MsgBox "Aviso: La hoja '" & HOJA_DATOS_ORIGINAL & "' no existe. No se pudo crear la copia de trabajo.", vbExclamation
    End If
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Entorno inicializado correctamente. Hojas base creadas.", vbInformation, "Listo"
    Exit Sub
Errores:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Error al inicializar el entorno: " & Err.Description, vbCritical
End Sub

' -------------------------------------------------------------------
' FUNCIÓN AUXILIAR PARA CREAR HOJAS SI NO EXISTEN
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
' FUNCIÓN PARA LLENAR LOS CATÁLOGOS
' -------------------------------------------------------------------
Private Sub CrearCatalogosBase()
    Dim ws As Worksheet
    
    CrearHojaSiNoExiste HOJA_SEXO, Array("ID", "DESCRIPCION")
    Set ws = ThisWorkbook.Sheets(HOJA_SEXO)
    If ws.Cells(ws.Rows.count, 2).End(xlUp).Row = 1 Then
        ws.Cells(2, 1).Value = 1: ws.Cells(2, 2).Value = "MASCULINO"
        ws.Cells(3, 1).Value = 2: ws.Cells(3, 2).Value = "FEMENINO"
    End If
    
    CrearHojaSiNoExiste HOJA_ATENCION, Array("ID", "DESCRIPCION")
    Set ws = ThisWorkbook.Sheets(HOJA_ATENCION)
    If ws.Cells(ws.Rows.count, 2).End(xlUp).Row = 1 Then
        ws.Cells(2, 1).Value = 1: ws.Cells(2, 2).Value = "CONSULTA"
        ws.Cells(3, 1).Value = 2: ws.Cells(3, 2).Value = "EMERGENCIA"
        ws.Cells(4, 1).Value = 3: ws.Cells(4, 2).Value = "CONTROL"
        ws.Cells(5, 1).Value = 4: ws.Cells(5, 2).Value = "REFERIDO"
    End If
    
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
    
    CrearHojaSiNoExiste HOJA_PROCEDENCIA, Array("ID", "DESCRIPCION")
    Set ws = ThisWorkbook.Sheets(HOJA_PROCEDENCIA)
    If ws.Cells(ws.Rows.count, 2).End(xlUp).Row = 1 Then
        ws.Cells(2, 1).Value = 1: ws.Cells(2, 2).Value = "TRABAJADORES"
        ws.Cells(3, 1).Value = 2: ws.Cells(3, 2).Value = "FAMILIARES"
        ws.Cells(4, 1).Value = 3: ws.Cells(4, 2).Value = "PACIENTES"
        ws.Cells(5, 1).Value = 4: ws.Cells(5, 2).Value = "JUBILADOS"
    End If
    
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
' FUNCIÓN PARA CREAR EL ADMINISTRADOR POR DEFECTO
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

