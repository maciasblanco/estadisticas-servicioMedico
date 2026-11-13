Attribute VB_Name = "modVariables"
Option Explicit

' ═══════════════════════════════════════════════════════════════════
'  MÓDULO: modVariables
'  Constantes, variables globales y utilidades del Sistema Médico
' ═══════════════════════════════════════════════════════════════════

' ───────────────────────────────────────────────────────────────────
' 1. IDENTIDAD DEL SISTEMA
' ───────────────────────────────────────────────────────────────────
Public Const SISTEMA_NOMBRE     As String = "Sistema de Gestión Médica"
Public Const SISTEMA_VERSION    As String = "2.0.0"
Public Const SISTEMA_AUTOR      As String = "Dirección de Salud"
Public Const SISTEMA_ANIO       As String = "2026"

' ───────────────────────────────────────────────────────────────────
' 2. SESIÓN DEL USUARIO
' ───────────────────────────────────────────────────────────────────
Public gUsuarioActual   As String
Public gNombreCompleto  As String
Public gRolUsuario      As String
Public gFechaLogin      As Date
Public gEquipoActual    As String
Public gSesionActiva    As Boolean

' ───────────────────────────────────────────────────────────────────
' 3. NOMBRES DE HOJAS
' ───────────────────────────────────────────────────────────────────
Public Const HOJA_DATOS          As String = "DatosOriginales (2)"
Public Const HOJA_SEXO           As String = "Datos-Sexo"
Public Const HOJA_ATENCION       As String = "Datos-TipoAtencion"
Public Const HOJA_MEDICO         As String = "Datos-MedicoTratante"
Public Const HOJA_PROCEDENCIA    As String = "Datos-Procedencia"
Public Const HOJA_DIAGNOSTICO    As String = "Datos-Diagnostico"
Public Const HOJA_ESTADISTICAS   As String = "Estadisticas agosto 2026"
Public Const HOJA_USUARIOS       As String = "Usuarios"
Public Const HOJA_HISTORICO      As String = "Historico"
Public Const HOJA_CONFIG         As String = "Configuracion"

' ───────────────────────────────────────────────────────────────────
' 4. COLUMNAS HOJA PRINCIPAL
' ───────────────────────────────────────────────────────────────────
Public Const COL_FECHA          As Integer = 1
Public Const COL_NOMBRE         As Integer = 2
Public Const COL_CEDULA         As Integer = 3
Public Const COL_SEXO           As Integer = 4
Public Const COL_EDAD           As Integer = 5
Public Const COL_ATENCION       As Integer = 6
Public Const COL_MEDICO         As Integer = 7
Public Const COL_CANT           As Integer = 8
Public Const COL_PROCEDENCIA    As Integer = 9
Public Const COL_DIAGNOSTICO    As Integer = 10
Public Const FILA_ENCABEZADO    As Long = 1

' ───────────────────────────────────────────────────────────────────
' 5. COLUMNAS HISTÓRICO
' ───────────────────────────────────────────────────────────────────
Public Const HCOL_FECHA_HORA     As Integer = 1
Public Const HCOL_USUARIO        As Integer = 2
Public Const HCOL_ACCION         As Integer = 3
Public Const HCOL_HOJA           As Integer = 4
Public Const HCOL_ID_REGISTRO    As Integer = 5
Public Const HCOL_CAMPO          As Integer = 6
Public Const HCOL_VALOR_ANT      As Integer = 7
Public Const HCOL_VALOR_NUEVO    As Integer = 8
Public Const HCOL_DETALLE        As Integer = 9
Public Const HCOL_EQUIPO         As Integer = 10

' ───────────────────────────────────────────────────────────────────
' 6. COLUMNAS USUARIOS
' ───────────────────────────────────────────────────────────────────
Public Const UCOL_ID             As Integer = 1
Public Const UCOL_USUARIO        As Integer = 2
Public Const UCOL_NOMBRE         As Integer = 3
Public Const UCOL_ROL            As Integer = 4
Public Const UCOL_HASH           As Integer = 5
Public Const UCOL_ACTIVO         As Integer = 6
Public Const UCOL_INTENTOS       As Integer = 7
Public Const UCOL_ULTIMO_ACCESO  As Integer = 8

' ───────────────────────────────────────────────────────────────────
' 7. ACCIONES DE AUDITORÍA
' ───────────────────────────────────────────────────────────────────
Public Const ACC_INSERT         As String = "INSERT"
Public Const ACC_UPDATE         As String = "UPDATE"
Public Const ACC_DELETE         As String = "DELETE"
Public Const ACC_LOGIN          As String = "LOGIN"
Public Const ACC_LOGOUT         As String = "LOGOUT"
Public Const ACC_LOGIN_FALLO    As String = "LOGIN_FALLO"
Public Const ACC_EXPORT         As String = "EXPORT"
Public Const ACC_ESTADISTICA    As String = "ESTADISTICA"
Public Const ACC_RESPALDO       As String = "RESPALDO"
Public Const ACC_CONFIG         As String = "CONFIG"

' ───────────────────────────────────────────────────────────────────
' 8. ROLES
' ───────────────────────────────────────────────────────────────────
Public Const ROL_ADMIN          As String = "ADMIN"
Public Const ROL_MEDICO         As String = "MEDICO"
Public Const ROL_ESTADISTICA    As String = "ESTADISTICA"
Public Const ROL_CONSULTA       As String = "CONSULTA"

' ───────────────────────────────────────────────────────────────────
' 9. PERÍODOS ESTADÍSTICOS
' ───────────────────────────────────────────────────────────────────
Public Const PER_MES            As String = "Por Mes"
Public Const PER_TRIMESTRE      As String = "Por Trimestre"
Public Const PER_SEMESTRE       As String = "Por Semestre"
Public Const PER_RANGO          As String = "Rango Específico"

' ───────────────────────────────────────────────────────────────────
' 10. SEGURIDAD
' ───────────────────────────────────────────────────────────────────
Public Const MAX_INTENTOS       As Integer = 3
Public Const MIN_LONG_CLAVE     As Integer = 8
Public Const SALT_FIJO          As String = "SIS-MED-2026-VZLA"

' ───────────────────────────────────────────────────────────────────
' 11. RUTAS
' ───────────────────────────────────────────────────────────────────
Public Const CARPETA_RESPALDO   As String = "C:\SistemaMedico\Respaldos\"
Public Const CARPETA_EXPORT     As String = "C:\SistemaMedico\Exportaciones\"

' ───────────────────────────────────────────────────────────────────
' 12. FORMATOS
' ───────────────────────────────────────────────────────────────────
Public Const FORMATO_FECHA      As String = "yyyy-mm-dd"
Public Const FORMATO_HORA       As String = "hh:mm:ss"
Public Const FORMATO_DATETIME   As String = "yyyy-mm-dd hh:mm:ss"

' ───────────────────────────────────────────────────────────────────
' 13. LÍMITES
' ───────────────────────────────────────────────────────────────────
Public Const EDAD_MINIMA        As Integer = 0
Public Const EDAD_MAXIMA        As Integer = 120
Public Const ANIO_MINIMO        As Integer = 2020
Public Const ANIO_MAXIMO        As Integer = 2100

' ───────────────────────────────────────────────────────────────────
' 14. COLORES
' ───────────────────────────────────────────────────────────────────
Public Const COLOR_PRIMARIO     As Long = 7882008
Public Const COLOR_SECUNDARIO   As Long = 15194304
Public Const COLOR_EXITO        As Long = 5296274
Public Const COLOR_PELIGRO      As Long = 255
Public Const COLOR_ADVERTENCIA  As Long = 65535
Public Const COLOR_TEXTO_CLARO  As Long = 16777215
Public Const COLOR_FONDO_GRIS   As Long = 15921906

' ───────────────────────────────────────────────────────────────────
' 15. MENSAJES ESTÁNDAR
' ───────────────────────────────────────────────────────────────────
Public Const MSG_SIN_SELECCION  As String = "Debe seleccionar un registro de la lista."
Public Const MSG_CONFIRMAR_DEL  As String = "¿Está seguro de eliminar el registro seleccionado?"
Public Const MSG_GUARDADO_OK    As String = "Registro guardado correctamente."
Public Const MSG_ACTUALIZADO_OK As String = "Registro actualizado correctamente."
Public Const MSG_ELIMINADO_OK   As String = "Registro eliminado correctamente."
Public Const MSG_SIN_RESULTADOS As String = "No se encontraron registros con ese criterio."
Public Const MSG_SIN_PERMISO    As String = "No tiene permisos para realizar esta acción."
Public Const MSG_CAMPOS_VACIOS  As String = "Los campos obligatorios deben estar completos."

' ═══════════════════════════════════════════════════════════════════
' 16. FUNCIONES UTILITARIAS
' ═══════════════════════════════════════════════════════════════════

Public Function ExisteHoja(ByVal nombre As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombre)
    On Error GoTo 0
    ExisteHoja = Not ws Is Nothing
End Function

Public Function TieneRol(ByVal rol As String) As Boolean
    TieneRol = (UCase$(gRolUsuario) = UCase$(rol))
End Function

Public Function EsAdmin() As Boolean
    EsAdmin = TieneRol(ROL_ADMIN)
End Function

Public Function TienePermiso(ByVal accion As String) As Boolean
    Select Case UCase$(accion)
        Case "INSERT", "UPDATE", "DELETE"
            TienePermiso = EsAdmin() Or TieneRol(ROL_MEDICO) Or TieneRol(ROL_ESTADISTICA)
        Case "EXPORT"
            TienePermiso = EsAdmin() Or TieneRol(ROL_ESTADISTICA)
        Case "USUARIOS", "RESPALDO", "CONFIG"
            TienePermiso = EsAdmin()
        Case Else
            TienePermiso = True
    End Select
End Function

Public Sub InicializarSesion(ByVal usuario As String, ByVal nombre As String, ByVal rol As String)
    gUsuarioActual  = UCase$(Trim$(usuario))
    gNombreCompleto = Trim$(nombre)
    gRolUsuario     = UCase$(Trim$(rol))
    gFechaLogin     = Now
    gEquipoActual   = Environ$("COMPUTERNAME")
    gSesionActiva   = True
End Sub

Public Sub CerrarSesion()
    gUsuarioActual  = vbNullString
    gNombreCompleto = vbNullString
    gRolUsuario     = vbNullString
    gFechaLogin     = 0
    gSesionActiva   = False
End Sub

Public Sub RangoMes(ByVal anio As Integer, ByVal mes As Integer, _
                    ByRef fIni As Date, ByRef fFin As Date)
    fIni = DateSerial(anio, mes, 1)
    fFin = DateSerial(anio, mes + 1, 0)
End Sub

Public Sub RangoTrimestre(ByVal anio As Integer, ByVal trimestre As Integer, _
                          ByRef fIni As Date, ByRef fFin As Date)
    Dim mesIni As Integer
    mesIni = ((trimestre - 1) * 3) + 1
    fIni = DateSerial(anio, mesIni, 1)
    fFin = DateSerial(anio, mesIni + 3, 0)
End Sub

Public Sub RangoSemestre(ByVal anio As Integer, ByVal semestre As Integer, _
                         ByRef fIni As Date, ByRef fFin As Date)
    If semestre = 1 Then
        fIni = DateSerial(anio, 1, 1)
        fFin = DateSerial(anio, 6, 30)
    Else
        fIni = DateSerial(anio, 7, 1)
        fFin = DateSerial(anio, 12, 31)
    End If
End Sub