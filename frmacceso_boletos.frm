VERSION 5.00
Begin VB.Form frmacceso_boletos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clave de acceso"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_contraseña 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1380
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   585
      Width           =   1800
   End
   Begin VB.TextBox txt_usuario 
      Height          =   345
      Left            =   1380
      TabIndex        =   1
      Top             =   165
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña:"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   645
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuario:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   255
      Width           =   585
   End
End
Attribute VB_Name = "frmacceso_boletos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txt_contraseña_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If UCase(Me.txt_usuario) = "JFPEREZ" And UCase(Me.txt_contraseña) = "JFPEREZ" Then
         VAR_GLOBAL_ACCESO_SORTEO = 1
      Else
         VAR_GLOBAL_ACCESO_SORTEO = 0
      End If
      Unload Me
   End If
End Sub

Private Sub txt_usuario_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub
