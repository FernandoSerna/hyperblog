VERSION 5.00
Begin VB.Form frmpasswords2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirmación"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   Icon            =   "frmpasswords2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_usuario 
      Height          =   315
      Left            =   2685
      TabIndex        =   0
      Top             =   75
      Width           =   1725
   End
   Begin VB.TextBox txt_password 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2685
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   435
      Width           =   1725
   End
   Begin VB.TextBox txt_supervisor 
      Height          =   315
      Left            =   2685
      TabIndex        =   2
      Top             =   795
      Width           =   1725
   End
   Begin VB.TextBox txt_password_supervisor 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2685
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1155
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuario"
      Height          =   195
      Left            =   855
      TabIndex        =   7
      Top             =   135
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña:"
      Height          =   195
      Left            =   855
      TabIndex        =   6
      Top             =   495
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Supervisor:"
      Height          =   195
      Left            =   855
      TabIndex        =   5
      Top             =   855
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña Supervisor:"
      Height          =   195
      Left            =   855
      TabIndex        =   4
      Top             =   1215
      Width           =   1650
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "frmpasswords2.frx":08CA
      Top             =   525
      Width           =   480
   End
End
Attribute VB_Name = "frmpasswords2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_password As String
Dim var_clave_supervisor As String
Dim var_usuario_supervisor As String
Dim var_password_supervisor As String

Private Sub Form_Unload(Cancel As Integer)
   If var_activa_menu = True Then
      Frmmenu2.Enabled = True
   End If
   var_swpassword = False
   If sw_primera_validacion Then
      sw_primera_validacion = False
      sw_mostrar_forma = True
      Call menuvisible(Frmmenu2, True)
   Else
      sw_primera_validacion = True
      sw_mostrar_forma = False
      Call menuvisible(Frmmenu2, False)
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
         txt_supervisor.Enabled = True
         txt_supervisor.SetFocus
      Else
         si = MsgBox("Password incorrecto ¿Deseas volverlo a intentar?", vbYesNo, "ATENCION")
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

Private Sub txt_password2_KeyPress(KeyAscii As Integer)
End Sub

Private Sub txt_password_supervisor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If txt_password_supervisor = var_password_supervisor Then
         If var_opcion_seguridad = 1 Then
            Unload Me
            ejecuta_forma
            rsaux.Open "insert into tb_bitacora_seguridad (VCHA_BOT_FORMA, VCHA_BIT_PROCESO, VCHA_BIT_USUARIO, VCHA_BTI_SUPERVISOR, DTIM_BIT_FECHA) values ('" + var_accion_submenu + "', 'ENTRADA','" + var_clave_usuario_global + "', '" + var_clave_supervisor + "', GETDATE())", cnn, adOpenDynamic, adLockOptimistic
         End If
         If var_opcion_seguridad = 2 Then
            rsaux.Open "insert into tb_bitacora_seguridad (VCHA_BOT_FORMA, VCHA_BIT_PROCESO, VCHA_BIT_USUARIO, VCHA_BTI_SUPERVISOR, DTIM_BIT_FECHA) values ('" + var_accion_submenu + "', '" + var_cadena_seguridad + "','" + var_clave_usuario_global + "', '" + var_clave_supervisor + "', GETDATE())", cnn, adOpenDynamic, adLockOptimistic
            ejecuta_cambios
            Unload Me
         End If
      Else
         si = MsgBox("Password incorrecto, ¿Deseas volver a intentarlo?", vbYesNo, "ATENCION")
         If si = 6 Then
            txt_password_supervisor = ""
            txt_password_supervisor.SetFocus
         Else
            Unload Me
         End If
      End If
   End If
End Sub

Private Sub txt_supervisor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      rs.Open "select * from VW_SUPERVISORES where vcha_uor_unidad_id = '" + var_empresa_global + "' and vcha_blo_bloque_id = '" + var_bloque_global + "' and vcha_for_forma_id = '" + var_accion_submenu + "' and vcha_usu_usuario = '" + txt_supervisor + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_clave_supervisor = rs!vcha_usu_usuario_id
         var_password_supervisor = rs(2).Value
         txt_password_supervisor.SetFocus
      Else
        si = MsgBox("Supervisor incorrecto, ¿Deseas volver a intentarlo?", vbYesNo, "ATENCION")
         If si = 6 Then
            txt_supervisor = ""
            txt_supervisor.SetFocus
         Else
            Unload Me
         End If
      End If
      rs.Close
   End If
End Sub

Private Sub txt_usuario_KeyPress(KeyAscii As Integer)
   Dim var_clave_usuario As String
   If KeyAscii = 27 Then
      Unload Me
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      rs.Open "select * from tb_usuarios where vcha_usu_usuario = '" + txt_usuario + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_clave_usuario = rs(0).Value
         var_password = rs(4).Value
         If var_clave_usuario = var_clave_usuario_global Then
           txt_password.Enabled = True
           txt_password.SetFocus
         End If
      Else
         MsgBox "Usuario incorrecto", vbOKOnly, "ATENCION"
         txt_usuario.SetFocus
      End If
      rs.Close
   End If
End Sub
