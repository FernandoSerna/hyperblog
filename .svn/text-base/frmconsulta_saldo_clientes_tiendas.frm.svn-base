VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmconsulta_saldo_clientes_tiendas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de saldos de clientes "
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   1965
      Left            =   1695
      TabIndex        =   10
      Top             =   -45
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1410
         Left            =   30
         TabIndex        =   11
         Top             =   450
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   2487
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7584
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   12
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos del cliente "
      Height          =   1680
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   8070
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   350
         Left            =   2535
         TabIndex        =   2
         Top             =   345
         Width           =   5400
      End
      Begin VB.TextBox txt_disponible 
         Alignment       =   1  'Right Justify
         Height          =   350
         Left            =   4545
         TabIndex        =   5
         Top             =   1125
         Width           =   2220
      End
      Begin VB.TextBox txt_real 
         Alignment       =   1  'Right Justify
         Height          =   350
         Left            =   1005
         TabIndex        =   4
         Top             =   1125
         Width           =   2220
      End
      Begin VB.TextBox txt_referencia 
         Height          =   350
         Left            =   1005
         TabIndex        =   3
         Top             =   735
         Width           =   2220
      End
      Begin VB.TextBox txt_cliente 
         Height          =   350
         Left            =   1005
         TabIndex        =   1
         Top             =   345
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Disponible:"
         Height          =   195
         Left            =   3705
         TabIndex        =   9
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Real:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Referencia:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   810
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   420
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmconsulta_saldo_clientes_tiendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Top = 2500
   Left = 1800
   If cnn_clientes_tiendas.State = 0 Then
      cnn_clientes_tiendas.Open var_conexion_pedidos_tiendas
      cnn_clientes_tiendas.CursorLocation = adUseClient
   End If
   Me.frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_cliente = lv_lista.selectedItem
         txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
         Me.txt_cliente.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE from VW_clientes WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 5 Then
         lv_lista.ColumnHeaders(2).Width = 4070.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4299.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_cliente_LostFocus()
   If Trim(Me.txt_cliente) <> "" Then
      rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         Me.txt_referencia = IIf(IsNull(rs!VCHA_CLI_REFERENCIA), "", rs!VCHA_CLI_REFERENCIA)
         rsaux.Open "select VCHA_SAL_REFERENCIA, NUMB_SAL_IMPORTE_DISPONIBLE, NUMB_SAL_IMPORTE from tb_saldo where vcha_sal_referencia = '" + Trim(Me.txt_referencia) + "'", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            Me.txt_disponible = Format(IIf(IsNull(rsaux!NUMB_SAL_IMPORTE_DISPONIBLE), 0, rsaux!NUMB_SAL_IMPORTE_DISPONIBLE), "###,###,##0.00")
            Me.txt_real = Format(IIf(IsNull(rsaux!NUMB_SAL_IMPORTE), 0, rsaux!NUMB_SAL_IMPORTE), "###,###,##0.00")
         Else
            Me.txt_disponible = ""
            Me.txt_real = ""
         End If
         rsaux.Close
      Else
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
         Me.txt_disponible = ""
         Me.txt_real = ""
         Me.txt_nombre_cliente = ""
      End If
      rs.Close
   Else
      Me.txt_disponible = ""
      Me.txt_real = ""
      Me.txt_nombre_cliente = ""
   End If
End Sub

Private Sub txt_disponible_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 2280
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE from VW_clientes VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 5 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_real_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_referencia_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub
