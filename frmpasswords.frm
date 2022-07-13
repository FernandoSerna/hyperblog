VERSION 5.00
Begin VB.Form frmpasswords 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confirmación"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3915
   Icon            =   "frmpasswords.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   3915
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_usuario 
      Height          =   315
      Left            =   1890
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txt_password 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1890
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clave:"
      Height          =   195
      Left            =   915
      TabIndex        =   3
      Top             =   300
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contraseña:"
      Height          =   195
      Left            =   915
      TabIndex        =   2
      Top             =   645
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "frmpasswords.frx":08CA
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmpasswords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_password As String

Private Sub cmd_aceptar_Click()

End Sub


Private Sub Form_Unload(Cancel As Integer)
   If var_activa_menu = True Then
      Frmmenu2.Enabled = True
   End If
End Sub

Private Sub txt_password_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If var_password = txt_password Then
         If var_opcion_seguridad = 1 Then
            If Trim(Me.txt_password) <> "" And Trim(Me.txt_usuario) <> "" Then
               Unload Me
               ejecuta_forma
               rsaux.Open "insert into tb_bitacora_seguridad (VCHA_BOT_FORMA, VCHA_BIT_PROCESO, VCHA_BIT_USUARIO, VCHA_BTI_SUPERVISOR, DTIM_BIT_FECHA) values ('" + var_accion_submenu + "', 'ENTRADA','" + var_clave_usuario_global + "', '', GETDATE())", cnn, adOpenDynamic, adLockOptimistic
            End If
         End If
         If var_opcion_seguridad = 2 Then
            If Trim(Me.txt_password) <> "" And Trim(Me.txt_usuario) <> "" Then
               ejecuta_cambios
               rsaux.Open "insert into tb_bitacora_seguridad (VCHA_BOT_FORMA, VCHA_BIT_PROCESO, VCHA_BIT_USUARIO, VCHA_BTI_SUPERVISOR, DTIM_BIT_FECHA) values ('" + var_accion_submenu + "', '" + var_cadena_seguridad + "','" + var_clave_usuario_global + "', '', GETDATE())", cnn, adOpenDynamic, adLockOptimistic
               Unload Me
             End If
         End If
         If var_opcion_seguridad = 3 Then
            If Trim(Me.txt_password) <> "" And Trim(Me.txt_usuario) <> "" Then
               var_autoriza_mov = True
               Unload Me
            End If
         End If
      Else
         si = MsgBox("Contraseña incorrecta ¿Deseas volverlo a intentar?", vbYesNo, "ATENCION")
         If si = 6 Then
            txt_password = ""
            txt_password.SetFocus
         Else
            Unload Me
            Frmmenu2.Show
         End If
      End If
   End If
End Sub


Private Sub txt_usuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_usuario_LostFocus()
   Dim var_clave_usuario As String
   If Trim(txt_usuario) <> "" Then
      rsaux2.Open "select * from tb_usuarios where vcha_usu_usuario = '" + txt_usuario + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux2.EOF Then
         var_clave_usuario = rsaux2(0).Value
         var_password = rsaux2(4).Value
         If var_clave_usuario = var_clave_usuario_global Then
            txt_password.Enabled = True
            txt_password.SetFocus
         End If
      Else
         MsgBox "Usuario incorrecto", vbOKOnly, "ATENCION"
         txt_usuario.SetFocus
      End If
      rsaux2.Close
   End If
End Sub
