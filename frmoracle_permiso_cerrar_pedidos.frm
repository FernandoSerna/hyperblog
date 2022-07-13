VERSION 5.00
Begin VB.Form frmoracle_permiso_cerrar_pedidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autoriza cerrado de pedidos"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtusuario 
      Height          =   315
      Left            =   1110
      MaxLength       =   13
      TabIndex        =   0
      Top             =   165
      Width           =   1815
   End
   Begin VB.TextBox txtpassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1110
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clave:"
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   225
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contraseña:"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   570
      Width           =   855
   End
End
Attribute VB_Name = "frmoracle_permiso_cerrar_pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   var_si_permiso = 0
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Me.txtpassword <> "" Then
         If Me.txtpassword = var_password_permiso Then
            var_si_permiso = 1
            If Me.txtusuario = "" Or Me.txtpassword = "" Then
               var_si_permiso = 0
            End If
            Unload Me
         Else
            MsgBox "Contraseña incorrecta", vbOKOnly, "ATENCION"
            var_si_permiso = 0
         End If
      Else
         MsgBox "Contraseña incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      var_si_permiso = 0
      Unload Me
   End If
End Sub

Private Sub txtusuario_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Me.txtusuario <> "" Then
         Me.txtpassword.SetFocus
      Else
         MsgBox "Usuario incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      Unload Me
      var_si_permiso = 0
   End If
End Sub

Private Sub txtusuario_LostFocus()
 If Trim(Me.txtusuario) <> "" Then
    rs.Open "select * from tb_usuarios where vcha_usu_usuario = '" + Me.txtusuario + "'", cnn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
       rsaux.Open "select * from TB_ORACLE_USUARIOS_PERMISO_CERRAR_PEDIDOS where vcha_usu_usuario_id = '" + rs!vcha_usu_usuario_id + "'", cnn, adOpenDynamic, adLockOptimistic
       If Not rsaux.EOF Then
          var_usuario_permiso = rsaux!vcha_usu_usuario_id
          var_password_permiso = rs!VCHA_USU_PASSWORD
       Else
          MsgBox "El usuario no tiene permiso para efectuar la acción", vbOKOnly, "ATENCION"
       End If
       rsaux.Close
    Else
       MsgBox "El usuario no existe", vbOKOnly, "ATENCION"
    End If
    rs.Close
 Else
    Me.txtpassword = ""
 End If
End Sub
