VERSION 5.00
Begin VB.Form frmoracle_orden_a_depurar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orden a depurar"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_orden 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   570
      Left            =   165
      TabIndex        =   0
      Top             =   180
      Width           =   2445
   End
End
Attribute VB_Name = "frmoracle_orden_a_depurar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Top = 3300
   Left = 4500
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_orden_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_orden) Then
         var_tipo_depurado = 0
         var_cadena_pedidos_global = Me.txt_orden
         frmoracle_depurar_pedidos.Show 1
      End If
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub
