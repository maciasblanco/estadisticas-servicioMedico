VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "Iniciar Sesión"
   ClientHeight    =   2505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3615
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    Me.txtUsuario.SetFocus
End Sub

Private Sub btnIngresar_Click()
    Dim usuario As String, clave As String, mensaje As String
    Dim usuarioLimpio As String
    
    usuarioLimpio = UCase(Trim(Me.txtUsuario.Value))
    clave = Trim(Me.txtClave.Value)
    
    If usuarioLimpio = "" Or clave = "" Then
        MsgBox "Debe ingresar usuario y clave.", vbExclamation, "Atención"
        Exit Sub
    End If
    
    If ValidarLogin(usuarioLimpio, clave, mensaje) Then
        RegistrarHistorico ACC_LOGIN, "SISTEMA", "", "", "", "", "Ingreso al sistema"
        Unload Me
        ' Línea clave: Abre el menú principal
        frmMenuPrincipal.Show
    Else
        RegistrarHistorico ACC_LOGIN_FALLO, "SISTEMA", "", "USUARIO", usuarioLimpio, "", mensaje
        MsgBox mensaje, vbCritical, "Error de acceso"
        Me.txtClave.Value = ""
        Me.txtClave.SetFocus
    End If
End Sub

Private Sub btnSalir_Click()
    Unload Me
End Sub
