VERSION 5.00
Begin VB.Form frmoracle_desbloquear_cajas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desbloquear caja"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2115
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   2115
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   1995
      Begin VB.TextBox txt_caja 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   135
         TabIndex        =   1
         Top             =   195
         Width           =   1740
      End
   End
End
Attribute VB_Name = "frmoracle_desbloquear_cajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Top = 3000
   Left = 4500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_caja_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "select * from tb_oracle_cajas_unicas_embarques where caja = '" + Me.txt_caja + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_si = MsgBox("¿Desea eliminar la caja?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rsaux1.Open "DELETE from tb_oracle_cajas_unicas_embarques where caja = '" + Me.txt_caja + "'", cnn, adOpenDynamic, adLockOptimistic
            rsaux1.Open "INSERT INTO TB_ORACLE_BITACORA_CAJAS_ELIMINADAS (CAJA, USUARIO, FECHA_HORA) VALUES ('" + Me.txt_caja + "','" + var_clave_usuario_global + "',GETDATE())", cnn, adOpenDynamic, adLockOptimistic
            MsgBox "La caja a sido eliminada", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "La caja no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub
