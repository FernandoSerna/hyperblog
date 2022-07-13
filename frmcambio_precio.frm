VERSION 5.00
Begin VB.Form frmcambio_precio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de precios"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Ejecutar cambio de precio"
      Height          =   675
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4680
   End
End
Attribute VB_Name = "frmcambio_precio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   var_si = MsgBox("¿Ejecuta el cambio de precio?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_si = MsgBox("Confirmar el cambio de precios", vbYesNo, "ATENCION")
      If var_si = 6 Then
         'rs.Open "EXEC CAMBIO_PRECIO", cnn, adOpenDynamic, adLockOptimistic
         MsgBox "Se a terminado el cambio de precio", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub Form_Load()
   Top = 1000
   Left = 2000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_packing_list)
End Sub
