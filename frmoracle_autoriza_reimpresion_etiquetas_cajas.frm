VERSION 5.00
Begin VB.Form frmoracle_autoriza_reimpresion_etiquetas_cajas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autoriza reimpresión de etiqueta"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   135
      TabIndex        =   2
      Top             =   0
      Width           =   3255
      Begin VB.TextBox txt_usuario 
         Height          =   375
         Left            =   1365
         TabIndex        =   0
         Top             =   225
         Width           =   1740
      End
      Begin VB.TextBox txt_contraseña 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1365
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   645
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   255
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña:"
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   630
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmoracle_autoriza_reimpresion_etiquetas_cajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   If var_leyenda_reimpresion <> "" Then
      Me.Caption = var_leyenda_reimpresion
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.txt_usuario <> "" Then
      If Me.txt_contraseña <> "" Then
         rsaux.Open "SELECT * FROM TB_USUARIOS A, TB_ORACLE_USUARIOS_AUTORIZAN_REIMPRESION B WHERE A.VCHA_USU_USUARIO_ID = B.VCHA_USU_USUARIO_ID AND VCHA_USU_USUARIO = '" + Me.txt_usuario + "' AND VCHA_USU_PASSWORD = '" + Me.txt_contraseña + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_autoriza_REIMPRESION = 1
            var_usuario_reimpresion = rsaux!vcha_usu_usuario_id
         Else
            var_autoriza_REIMPRESION = 0
         End If
         rsaux.Close
      Else
         var_autoriza_REIMPRESION = 0
      End If
   Else
      var_autoriza_REIMPRESION = 0
   End If
End Sub

Private Sub txt_contraseña_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_contraseña <> "" Then
         Unload Me
      Else
         MsgBox "Contraseña incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_usuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_usuario <> "" Then
         Me.txt_contraseña.SetFocus
      Else
         MsgBox "Usuario incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub
