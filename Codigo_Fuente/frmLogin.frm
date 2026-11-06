VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Iniciar Sesión - Sistema Médico"
   ClientHeight    =   2505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8295.001
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Version 5#

   Caption = "Iniciar Sesión - Sistema Médico"
   ClientHeight = 3375
   ClientLeft = 120
   ClientTop = 465
   ClientWidth = 8295
   StartUpPosition = 1    'Centrar en propietario
      Caption = "SISTEMA DE GESTIÓN MÉDICA"
      BeginProperty Font
         Name = "Tahoma"
         Size = 14
         Bold = -1               'True
         Charset = 0
      EndProperty
      Height = 375
      Left = 1500
      Top = 240
      Width = 5300
   End
      Caption = "USUARIO:"
      BeginProperty Font
         Name = "Tahoma"
         Size = 10
         Bold = -1               'True
         Charset = 0
      EndProperty
      Height = 255
      Left = 900
      Top = 960
      Width = 1500
   End
      BeginProperty Font
         Name = "Tahoma"
         Size = 10
         Charset = 0
      EndProperty
      Height = 315
      Left = 2400
      MaxLength = 50
      TabIndex = 0
      Top = 930
      Width = 3900
   End
      Caption = "CLAVE:"
      BeginProperty Font
         Name = "Tahoma"
         Size = 10
         Bold = -1               'True
         Charset = 0
      EndProperty
      Height = 255
      Left = 900
      Top = 1500
      Width = 1500
   End
      BeginProperty Font
         Name = "Tahoma"
         Size = 10
         Charset = 0
      EndProperty
      Height = 315
      Left = 2400
      MaxLength = 100
      PasswordChar = "*"
      TabIndex = 1
      Top = 1470
      Width = 3900
   End
      Caption = "&Ingresar"
      Default = -1            'True
      BeginProperty Font
         Name = "Tahoma"
         Size = 10
         Bold = -1               'True
         Charset = 0
      EndProperty
      Height = 420
      Left = 2400
      TabIndex = 2
      Top = 2100
      Width = 1850
   End
      Cancel = -1             'True
      Caption = "&Salir"
      BeginProperty Font
         Name = "Tahoma"
         Size = 10
         Charset = 0
      EndProperty
      Height = 420
      Left = 4450
      TabIndex = 3
      Top = 2100
      Width = 1850
   End
      Caption = "Versión 2.0.0 · 2026"
      BeginProperty Font
         Name = "Tahoma"
         Size = 8
         Charset = 0
      EndProperty
      ForeColor = &H808080
      Height = 195
      Left = 900
      Top = 2820
      Width = 6400
   End
End
Option Explicit

' -------------------------------------------------------------------
'  FORMULARIO: frmLogin
'  Inicio de sesión del Sistema de Gestión Médica
' -------------------------------------------------------------------

Private mProcesando As Boolean

' -------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Me.Caption = "Iniciar Sesión - " & SISTEMA_NOMBRE
    Me.lblTitulo.Caption = UCase$(SISTEMA_NOMBRE)
    Me.lblVersion.Caption = "Versión " & SISTEMA_VERSION & " · " & SISTEMA_ANIO
    Me.btnIngresar.Default = True
    Me.btnSalir.Cancel = True
    mProcesando = False
End Sub

' -------------------------------------------------------------------
Private Sub UserForm_Activate()
    Me.txtUsuario.SetFocus
End Sub

' -------------------------------------------------------------------
Private Sub btnIngresar_Click()
    If mProcesando Then Exit Sub
    
    Dim usuario As String
    Dim clave As String
    Dim mensaje As String
    
    usuario = UCase$(Trim$(Me.txtUsuario.Value))
    clave = Trim$(Me.txtClave.Value)
    
    If usuario = "" Or clave = "" Then
        MsgBox "Debe ingresar usuario y clave.", vbExclamation, "Atención"
        If usuario = "" Then
            Me.txtUsuario.SetFocus
        Else
            Me.txtClave.SetFocus
        End If
        Exit Sub
    End If
    
    ' -- Bloquear controles mientras se valida --------------------
    mProcesando = True
    Me.btnIngresar.Enabled = False
    Me.btnSalir.Enabled = False
    Me.MousePointer = fmMousePointerHourGlass
    
    On Error GoTo Fallo
    
    If ValidarLogin(usuario, clave, mensaje) Then
        RegistrarHistorico ACC_LOGIN, "SISTEMA", "", "", "", "", _
                            "Ingreso al sistema: " & usuario
        Me.MousePointer = fmMousePointerDefault
        Unload Me
        frmMenuPrincipal.Show
        Exit Sub
    End If
    
    ' -- Login fallido ------------------------------------------
    RegistrarHistorico ACC_LOGIN_FALLO, "SISTEMA", "", "USUARIO", usuario, "", mensaje
    Me.MousePointer = fmMousePointerDefault
    MsgBox mensaje, vbCritical, "Error de acceso"
    Me.txtClave.Value = ""
    Me.txtClave.SetFocus
    mProcesando = False
    Me.btnIngresar.Enabled = True
    Me.btnSalir.Enabled = True
    Exit Sub
    
Fallo:
    Me.MousePointer = fmMousePointerDefault
    mProcesando = False
    Me.btnIngresar.Enabled = True
    Me.btnSalir.Enabled = True
    MsgBox "Error inesperado al validar el acceso:" & vbCrLf & _
           Err.Description & " (" & Err.Number & ")", vbCritical, "Error"
End Sub

' -------------------------------------------------------------------
Private Sub btnSalir_Click()
    If mProcesando Then Exit Sub
    
    Dim resp As VbMsgBoxResult
    resp = MsgBox("¿Desea cerrar el sistema?" & vbCrLf & vbCrLf & _
                  "Sí        ? cerrar el archivo de Excel." & vbCrLf & _
                  "No        ? solo cerrar esta ventana." & vbCrLf & _
                  "Cancelar  ? continuar en el formulario.", _
                  vbYesNoCancel + vbQuestion, "Salir")
    
    Select Case resp
        Case vbYes
            On Error Resume Next
            ThisWorkbook.Close SaveChanges:=False
            If Err.Number <> 0 Then
                MsgBox "No se pudo cerrar el archivo. Ciérrelo manualmente.", _
                       vbExclamation, "Atención"
                Err.Clear
            End If
            On Error GoTo 0
        Case vbNo
            Unload Me
        Case vbCancel
            Me.txtUsuario.SetFocus
    End Select
End Sub

' -------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mProcesando Then
        Cancel = True
        Exit Sub
    End If
    
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        MsgBox "Use el botón Salir para cerrar el sistema." & vbCrLf & vbCrLf & _
               "Si cierra Excel, el sistema se cerrará automáticamente.", _
               vbInformation, "Atención"
    End If
End Sub

' -------------------------------------------------------------------
'  Enter en "Usuario" salta a "Clave" (Enter en "Clave" dispara
'  el botón por defecto "Ingresar").
' -------------------------------------------------------------------
Private Sub txtUsuario_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                               ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        Me.txtClave.SetFocus
    End If
End Sub

