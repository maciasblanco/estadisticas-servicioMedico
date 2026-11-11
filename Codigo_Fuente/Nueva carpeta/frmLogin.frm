VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Iniciar Sesión - Sistema Médico"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   StartUpPosition =   1  'CenterOwner
   Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} lblTitulo 
      Caption         =   "SISTEMA MÉDICO"
      Height          =   300
      Left            =   1320
      Top             =   240
      Width           =   1920
      Font.Size       =   14
      Font.Bold       =   -1  'True
      ForeColor       =   &H00784E1F&
   End
   Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} lblUsuario 
      Caption         =   "Usuario:"
      Height          =   200
      Left            =   600
      Top             =   840
      Width           =   800
      Font.Bold       =   -1  'True
   End
   Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} txtUsuario 
      Height          =   300
      Left            =   600
      Top             =   1080
      Width           =   3360
   End
   Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} lblClave 
      Caption         =   "Clave:"
      Height          =   200
      Left            =   600
      Top             =   1560
      Width           =   800
      Font.Bold       =   -1  'True
   End
   Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} txtClave 
      Height          =   300
      Left            =   600
      Top             =   1800
      Width           =   3360
      PasswordChar    =   "*"
   End
   Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} btnIngresar 
      Caption         =   "Ingresar"
      Height          =   360
      Left            =   2400
      Top             =   2400
      Width           =   1560
      BackColor       =   &H00784E1F&
      ForeColor       =   &H00FFFFFF&
      Font.Bold       =   -1  'True
   End
   Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} btnSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   600
      Top             =   2400
      Width           =   1200
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Option Explicit

Private Sub UserForm_Initialize()
    Me.txtUsuario.SetFocus
End Sub

Private Sub btnIngresar_Click()
    Dim usuario As String, clave As String, mensaje As String
    usuario = UCase(Trim(Me.txtUsuario.Value))
    clave = Trim(Me.txtClave.Value)
    
    If usuario = "" Or clave = "" Then
        MsgBox "Debe ingresar usuario y clave.", vbExclamation, "Atención"
        Exit Sub
    End If
    
    If ValidarLogin(usuario, clave, mensaje) Then
        RegistrarHistorico ACC_LOGIN, "SISTEMA", "", "", "", "", "Ingreso al sistema"
        Unload Me
        frmMenuPrincipal.Show
    Else
        RegistrarHistorico ACC_LOGIN_FALLO, "SISTEMA", "", "USUARIO", usuario, "", mensaje
        MsgBox mensaje, vbCritical, "Error de acceso"
        Me.txtClave.Value = ""
        Me.txtClave.SetFocus
    End If
End Sub

Private Sub btnSalir_Click()
    Unload Me
    ' Dejar Excel abierto sin el sistema
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        MsgBox "Debe ingresar al sistema o cerrar el archivo de Excel.", vbInformation
    End If
End Sub