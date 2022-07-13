VERSION 5.00
Begin VB.Form frmsupervisor1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supervisor"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_password 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1830
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   570
      Width           =   1815
   End
   Begin VB.TextBox txt_usuario 
      Height          =   315
      Left            =   1830
      TabIndex        =   0
      Top             =   210
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frmsupervisor1.frx":0000
      Top             =   210
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña:"
      Height          =   195
      Left            =   855
      TabIndex        =   3
      Top             =   615
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Clave:"
      Height          =   195
      Left            =   855
      TabIndex        =   2
      Top             =   270
      Width           =   450
   End
End
Attribute VB_Name = "frmsupervisor1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_password As String
Dim var_usuario_supervisor As String
Dim var_password_supervisor As String

Private Sub Form_Load()
   var_cadena_seguridad = ""
   var_posible_accion = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_supervisor1)
End Sub

Private Sub txt_password_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If txt_password = var_password_supervisor Then
         var_posible_accion = True
         Unload Me
      Else
         si = MsgBox("Password icorrecto, ¿Deseas volver a intentarlo?", vbYesNo, "ATENCION")
         If si = 6 Then
            txt_password = ""
            txt_password.SetFocus
            var_posible_accion = False
         Else
            Unload Me
            var_posible_accion = False
         End If
      End If
   End If
End Sub

Private Sub txt_usuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      var_global_supervisor_1 = ""
      var_global_supervisor_2 = ""
      rs.Open "select * from VW_SUPERVISORES where vcha_uor_unidad_id = '" + var_empresa_global + "' and vcha_blo_bloque_id = '" + var_bloque_global + "' and vcha_for_forma_id = '" + var_accion_submenu + "' and vcha_usu_usuario = '" + txt_usuario + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_password_supervisor = rs(2).Value
         var_global_supervisor_1 = rs!vcha_usu_usuario_id
         txt_password.SetFocus
      Else
        si = MsgBox("Supervisor incorrecto, ¿Deseas volver a intentarlo?", vbYesNo, "ATENCION")
         If si = 6 Then
            txt_usuario = ""
            txt_usuario.SetFocus
         Else
            Unload Me
         End If
      End If
      rs.Close
   End If
End Sub
