VERSION 5.00
Begin VB.Form Frmacceso 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acceso al Sistema"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3915
   Icon            =   "fracceso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   3915
   Begin VB.TextBox txtpassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1890
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtusuario 
      Height          =   315
      Left            =   1890
      MaxLength       =   13
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "fracceso.frx":08CA
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contraseņa:"
      Height          =   195
      Left            =   915
      TabIndex        =   1
      Top             =   645
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clave:"
      Height          =   195
      Left            =   915
      TabIndex        =   0
      Top             =   300
      Width           =   450
   End
End
Attribute VB_Name = "Frmacceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_password As String
Public var_parametros_empresa As String
Dim var_encontro As Integer



Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      End
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   var_top = 3000
   var_left = 3850
   Frmacceso.Top = var_top
   Frmacceso.Left = var_left
   var_encontro = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_encontro = 0 Then
      End
   End If
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
'   On Error GoTo salir:
   Static var_veces As Byte
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(var_password) = txtpassword.Text Then
         If var_clave_usuario_global = "1" Then
            frmmenu1.Show
            var_encontro = 1
            var_usuario_global = txtusuario
            var_passwor_global = txtpassword
            Unload Me
         Else
            Frmempresas.Show
            var_encontro = 1
            var_usuario_global = txtusuario
            var_passwor_global = txtpassword
            Unload Me
         End If
      Else
         MsgBox "No se puede accesar al sistema, asegurese que la contraseņa sea la correcta.", vbOKOnly + vbCritical, "ATENCION"
         txtpassword.Text = ""
         txtpassword.SetFocus
      End If
   End If
   Exit Sub
SALIR:
   MsgBox "Existe un problema con la conexión del sistema", vbOKOnly, "ATENCION"
   End
End Sub


Private Sub txtusuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txtpassword.SetFocus
   End If
End Sub

Private Sub txtusuario_LostFocus()
'On Error GoTo salir:
      KeyAscii = 0
      'MsgBox parametros(0)
      'MsgBox cnn.ConnectionString
      'MsgBox cnn.ConnectionString
      
      rs.Open "SELECT * from TB_USUARIOS WHERE VCHA_USU_USUARIO = '" & txtusuario & "'", cnn, adOpenDynamic, adLockOptimistic
      If rs.RecordCount <> 0 Then
         var_clave_usuario_global = rs!vcha_usu_usuario_id
         var_password = rs!VCHA_USU_PASSWORD
         var_tipo_permiso = rs!INTE_USU_PERMISO
         var_nombre_usuario_global = rs!vcha_usu_nombre
         var_apellidos_usuario_global = rs!vcha_usu_apellidos
         var_nombre_usuario = var_nombre_usuario_global + " " + var_apellidos_usuario_global
         txtpassword.SetFocus
      Else
         MsgBox "La clave de usuario es incorrecta", vbCritical, "ATENCION"
         txtusuario.Text = ""
         txtusuario.SetFocus
         KeyAscii = 0
      End If
      rs.Close
      Exit Sub
SALIR:
   MsgBox "Existe un problema con la conexión del sistema", vbOKOnly, "ATENCION"
   End
End Sub
