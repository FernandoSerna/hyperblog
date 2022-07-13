VERSION 5.00
Begin VB.Form frmoracle_autoriza_cerrar_pantalla 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autoriza"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   60
      TabIndex        =   0
      Top             =   -15
      Width           =   3255
      Begin VB.TextBox txt_contraseña 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1365
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   645
         Width           =   1740
      End
      Begin VB.TextBox txt_usuario 
         Height          =   375
         Left            =   1365
         TabIndex        =   3
         Top             =   225
         Width           =   1740
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña:"
         Height          =   195
         Left            =   135
         TabIndex        =   2
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   255
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmoracle_autoriza_cerrar_pantalla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txt_contraseña_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_contraseña_cerrar_pantalla <> "" Then
         If UCase(Me.txt_contraseña) = UCase(var_contraseña_cerrar_pantalla) Then
            Unload Me
         Else
            var_contraseña_cerrar_pantalla = ""
            MsgBox "No esta autorizado", vbOKOnly, "ATENCION"
            Unload Me
         End If
      Else
         var_contraseña_cerrar_pantalla = ""
         MsgBox "No esta autorizado", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_usuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_contraseña.SetFocus
   End If
End Sub

Private Sub txt_usuario_LostFocus()
   If Me.txt_usuario <> "" Then
      rsaux1.Open "select * from tb_usuarios where vcha_usu_usuario = '" + Me.txt_usuario + "' and INTE_USU_CERRAR_PANTALLA_PEDIDOS = 1", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux1.EOF Then
         var_usuario_cerrar_pantalla = rsaux1!vcha_usu_usuario_id
         var_contraseña_cerrar_pantalla = IIf(IsNull(rsaux1!VCHA_USU_PASSWORD), "", rsaux1!VCHA_USU_PASSWORD)
      Else
         MsgBox "Usuario incorrecto", vbOKOnly, "ATENCION"
         var_usuario_cerrar_pantalla = ""
         var_contraseña_cerrar_pantalla = ""
      End If
      rsaux1.Close
   Else
      var_usuario_cerrar_pantalla = ""
      var_contraseña_cerrar_pantalla = ""
      Me.txt_contraseña = ""
   End If
End Sub
