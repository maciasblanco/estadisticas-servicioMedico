VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenuPrincipal 
   Caption         =   "UserForm1"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   OleObjectBlob   =   "frmMenuPrincipal.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmMenuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    ' Muestra el nombre del usuario y su rol
    Me.lblUsuario.Caption = "Usuario: " & gNombreCompleto & " (" & gRolUsuario & ")"
    
    ' Si no es admin, ocultamos el botón de administrar usuarios
    If Not EsAdmin() Then
        Me.btnUsuarios.Visible = False
    End If
    
    ' Centrar el formulario en la pantalla
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

Private Sub btnDatos_Click()
    ' Esta línea abrirá el CRUD que crearemos en la Fase 4
    frmCrud.Show
End Sub

Private Sub btnEstadisticas_Click()
    MsgBox "Aquí abriremos el formulario de Estadísticas (Fase 5).", vbInformation, "En Construcción"
End Sub

Private Sub btnHistorial_Click()
    MsgBox "Aquí abriremos la auditoría de cambios.", vbInformation, "En Construcción"
End Sub

Private Sub btnUsuarios_Click()
    MsgBox "Aquí abriremos la gestión de usuarios.", vbInformation, "En Construcción"
End Sub

Private Sub btnSalir_Click()
    If MsgBox("¿Está seguro que desea cerrar sesión?", vbYesNo + vbQuestion, "Salir") = vbYes Then
        ' Registrar el cierre en el histórico
        RegistrarHistorico ACC_LOGOUT, "SISTEMA", "", "", "", "", "Cierre de sesión"
        
        ' Limpiar variables de sesión
        CerrarSesion
        
        ' Cerrar el menú
        Unload Me
    End If
End Sub
