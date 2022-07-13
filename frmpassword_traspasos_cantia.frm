VERSION 5.00
Begin VB.Form frmpassword_traspasos_cantia 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Clave de acceso de traspasos"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtusuario 
      Height          =   315
      Left            =   1830
      MaxLength       =   13
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtpassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1830
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   585
      Width           =   1815
   End
   Begin VB.Label lbl_movimiento 
      Height          =   300
      Left            =   1125
      TabIndex        =   4
      Top             =   945
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clave:"
      Height          =   195
      Left            =   855
      TabIndex        =   3
      Top             =   285
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contraseña:"
      Height          =   195
      Left            =   855
      TabIndex        =   2
      Top             =   630
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frmpassword_traspasos_cantia.frx":0000
      Top             =   225
      Width           =   480
   End
End
Attribute VB_Name = "frmpassword_traspasos_cantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_clave_usuario_traspasos As String
Dim var_password_traspasos As String
Public var_parametros_empresa As String
Dim var_encontro As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      var_acepta_traspaso_global = 0
   End If
End Sub

Private Sub Form_Load()
   var_acepta_traspaso_global = 0
   var_cadena_seguridad = ""
   var_top = 3000
   var_left = 3850
   Top = var_top
   Left = var_left
   var_encontro = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_encontro = 0 Then
      'End
   End If
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
   'On Error GoTo SALIR:
   Static var_veces As Byte
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(var_password_traspasos) = txtpassword.Text Then
         var_encontro = 1
         rs.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_traspasos + "' and vcha_mov_movimiento_id = 'T' and vcha_per_almacen_2 = '" + var_almacen_traspaso_cantia + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux.Open "INSERT INTO TB_FIRMAS_TRASPASOS (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_mov_movimiento_id, INTE_TRA_NUMERO, VCHA_USU_USUARIO_ID) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + Me.lbl_movimiento + "'," + CStr(var_numero_traspaso_cantia) + ",'" + var_clave_usuario_traspasos + "')", cnn, adOpenDynamic, adLockOptimistic
            var_acepta_traspaso_global = 1
         Else
            var_acepta_traspaso_global = 0
         End If
         rs.Close
         Unload Me
      Else
         MsgBox "Clave de usuario incorrecta.", vbOKOnly + vbCritical, "ATENCION"
         txtpassword.Text = ""
         txtpassword.SetFocus
      End If
   End If
   Exit Sub
salir:
   MsgBox "Existe un problema con la conexión del sistema", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
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
On Error GoTo salir:
      KeyAscii = 0
      'MsgBox parametros(0)
      rs.Open "SELECT * from TB_USUARIOS WHERE VCHA_USU_USUARIO = '" & txtusuario & "'", cnn, adOpenDynamic, adLockOptimistic
      If rs.RecordCount <> 0 Then
         var_clave_usuario_traspasos = rs!vcha_usu_usuario_ID
         var_password_traspasos = rs!VCHA_USU_PASSWORD
         txtpassword.SetFocus
      Else
         MsgBox "La clave de usuario es incorrecta", vbCritical, "ATENCION"
         txtusuario.Text = ""
         txtusuario.SetFocus
         KeyAscii = 0
      End If
      rs.Close
      Exit Sub
salir:
   MsgBox "Existe un problema con la conexión del sistema", vbOKOnly, "ATENCION"
   End
End Sub

