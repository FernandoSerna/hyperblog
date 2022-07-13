VERSION 5.00
Begin VB.Form frmcerrado_pedidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cerrado de Pedidos"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4005
   Begin VB.TextBox txt_pedido 
      Height          =   345
      Left            =   1875
      TabIndex        =   0
      Top             =   510
      Width           =   1815
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3600
      Picture         =   "frmcerrado_pedidos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_generar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmcerrado_pedidos.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Generar Ordenes de Surtido Alt + G"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   0
      TabIndex        =   1
      Top             =   255
      Width           =   3930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número de Pedido:"
      Height          =   195
      Left            =   210
      TabIndex        =   4
      Top             =   585
      Width           =   1365
   End
End
Attribute VB_Name = "frmcerrado_pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_generar_Click()
   Dim var_estatus As String
   Dim var_si As Integer
   If IsNumeric(txt_pedido) Then
      rs.Open "select * from tb_encabezado_pedidos where inte_ped_numero = " + txt_pedido, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_si = MsgBox("¿Deseas cancelar el pedido?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar el cerrado del pedido", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_estatus = IIf(IsNull(rs!CHAR_PED_ESTATUS), "", rs!CHAR_PED_ESTATUS)
               If var_estatus = "I" Or var_estatus = "S" Then
                  rsaux.Open "update tb_encabezado_pedidos set CHAR_PED_ESTATUS = 'E' where inte_ped_numero = " + txt_pedido, cnn, adOpenDynamic, adLockOptimistic
                  MsgBox "El pedido a sido cerrado", vbOKOnly, "ATENCION"
               Else
                  MsgBox "El pedido ya no puede ser cerrado", vbOKOnly, "ATENCION"
               End If
            End If
         End If
      Else
         MsgBox "El pedido no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      MsgBox "Número de pedido Incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3850
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_cerrado_pedidos)
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub
