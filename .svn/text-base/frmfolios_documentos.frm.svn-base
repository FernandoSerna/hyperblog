VERSION 5.00
Begin VB.Form frmfolios_documentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consecutivo "
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_cancelar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   375
      Picture         =   "frmfolios_documentos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancelar Alt + C"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmfolios_documentos.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   0
      TabIndex        =   9
      Top             =   345
      Width           =   3345
   End
   Begin VB.Frame Frame1 
      Caption         =   " Documentos "
      Height          =   1350
      Left            =   90
      TabIndex        =   2
      Top             =   390
      Width           =   3195
      Begin VB.TextBox txt_nota_cargo 
         Height          =   315
         Left            =   1515
         TabIndex        =   3
         Top             =   945
         Width           =   1455
      End
      Begin VB.TextBox txt_nota_credito 
         Height          =   315
         Left            =   1515
         TabIndex        =   1
         Top             =   570
         Width           =   1455
      End
      Begin VB.TextBox txt_factura 
         Height          =   315
         Left            =   1515
         TabIndex        =   0
         Top             =   195
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nota de Cargo:"
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   1005
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nota de Crédito:"
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " Factura:"
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   300
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmfolios_documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
   var_si = MsgBox("¿Desea cambiar los folios?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_si = MsgBox("Confirmar el cambio de folios", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rs.Open "UPDATE TB_SERIES SET INTE_SER_FACTURA = " + txt_factura + ", INTE_sER_NOTA_CARGO = " + txt_nota_cargo + ", INTE_SER_NOTA_CREDITO = " + txt_nota_credito + " WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
      End If
   End If
End Sub

Private Sub cmd_cancelar_pedidos_Click()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 4000
   rs.Open "select * from tb_series where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Me.txt_factura = IIf(IsNull(rs!inte_ser_factura), 0, rs!inte_ser_factura)
      Me.txt_nota_cargo = IIf(IsNull(rs!inte_ser_nota_Cargo), 0, rs!inte_ser_nota_Cargo)
      Me.txt_nota_credito = IIf(IsNull(rs!inte_ser_nota_credito), 0, rs!inte_ser_nota_credito)
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_reporte_comisiones)
End Sub

Private Sub txt_factura_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nota_cargo_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nota_credito_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub
