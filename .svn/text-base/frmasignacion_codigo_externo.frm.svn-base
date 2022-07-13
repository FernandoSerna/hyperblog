VERSION 5.00
Begin VB.Form frmasignacion_codigo_externo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de códigos externos"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   1125
      Left            =   60
      TabIndex        =   6
      Top             =   390
      Width           =   6990
      Begin VB.TextBox txt_codigo 
         Height          =   345
         Left            =   1035
         TabIndex        =   3
         Top             =   270
         Width           =   1545
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   345
         Left            =   2595
         TabIndex        =   4
         Top             =   270
         Width           =   4305
      End
      Begin VB.TextBox txt_externo 
         Height          =   345
         Left            =   1035
         TabIndex        =   5
         Top             =   630
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   345
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cantia:"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   705
         Width           =   495
      End
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmasignacion_codigo_externo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmasignacion_codigo_externo.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6690
      Picture         =   "frmasignacion_codigo_externo.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   120
      Left            =   0
      TabIndex        =   9
      Top             =   255
      Width           =   7050
   End
End
Attribute VB_Name = "frmasignacion_codigo_externo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
   Me.txt_codigo = ""
   Me.txt_descripcion = ""
   Me.txt_externo = ""
   Me.txt_codigo.SetFocus
End Sub

Private Sub cmd_cancelar_pedidos_Click()
   If Me.txt_codigo <> "" Then
      If Me.txt_externo <> "" Then
         var_si = MsgBox("Confirmar el cambio de código externo", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rs.Open "UPDATE TB_aRTICULOS SET VCHA_aRT_CODIGO_EXTERNO = '" + Me.txt_externo + "' WHERE VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
      Else
         MsgBox "Código externo incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Código incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub txt_agente_Change()

End Sub

Private Sub txt_agente_LostFocus()
   
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3200
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_codigo_Change()
   Me.txt_descripcion = ""
   Me.txt_externo = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_descripcion.SetFocus
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   If Me.txt_codigo <> "" Then
      rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_descripcion = IIf(IsNull(rs!VCHA_aRT_NOMBRE_ESPAÑOL), "", rs!VCHA_aRT_NOMBRE_ESPAÑOL)
         Me.txt_externo = IIf(IsNull(rs!VCHA_aRT_CODIGO_EXTERNO), "", rs!VCHA_aRT_CODIGO_EXTERNO)
      Else
         MsgBox "Código incorrecto", vbOKOnly, "ATENCION"
         Me.txt_codigo = ""
         Me.txt_descripcion = ""
         Me.txt_externo = ""
      End If
      rs.Close
   Else
      Me.txt_descripcion = ""
      Me.txt_externo = ""
   End If
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_externo.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_externo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_cancelar_pedidos.SetFocus
   End If
End Sub
