VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcambiar_descuentos_clientes_multibondeados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de descuento de clientes para Multibondeados"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2100
      Left            =   1170
      TabIndex        =   14
      Top             =   -75
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1575
         Left            =   30
         TabIndex        =   15
         Top             =   465
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   2778
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
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7057
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   16
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7050
      Picture         =   "frmcambiar_descuentos_clientes_multibondeados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmcambiar_descuentos_clientes_multibondeados.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Cliente "
      Height          =   1560
      Left            =   90
      TabIndex        =   8
      Top             =   390
      Width           =   7320
      Begin VB.TextBox txt_descuento_2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   4665
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txt_descuento_1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   1230
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txt_nombre_grupo 
         Height          =   350
         Left            =   2265
         TabIndex        =   3
         Top             =   570
         Width           =   4890
      End
      Begin VB.TextBox txt_clave_grupo 
         Height          =   350
         Left            =   795
         TabIndex        =   2
         Top             =   570
         Width           =   1455
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   350
         Left            =   2265
         TabIndex        =   1
         Top             =   195
         Width           =   4890
      End
      Begin VB.TextBox txt_clave_cliente 
         Height          =   350
         Left            =   795
         TabIndex        =   0
         Top             =   195
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descuento 2:"
         Height          =   195
         Left            =   3555
         TabIndex        =   13
         Top             =   1103
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descuento 1:"
         Height          =   195
         Left            =   165
         TabIndex        =   12
         Top             =   1103
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
         Height          =   195
         Left            =   165
         TabIndex        =   11
         Top             =   648
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   273
         Width           =   525
      End
   End
   Begin VB.Frame Frame2 
      Height          =   90
      Left            =   30
      TabIndex        =   9
      Top             =   285
      Width           =   7470
   End
End
Attribute VB_Name = "frmcambiar_descuentos_clientes_multibondeados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
   If Me.txt_descuento_1 = "" Then
      Me.txt_descuento_1 = "0"
   End If
   If Me.txt_descuento_2 = "" Then
      Me.txt_descuento_2 = "0"
   End If
   If IsNumeric(Me.txt_descuento_1) Then
      If IsNumeric(Me.txt_descuento_2) Then
         If Me.txt_clave_grupo <> "" Then
            rs.Open "select vcha_emp_Empresa_id from vw_clientes where vcha_gac_grupo_actual_id = '" + Me.txt_clave_grupo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_empresa_grupo = IIf(IsNull(rs!vcha_emp_empresa_id), "", rs!vcha_emp_empresa_id)
               If var_empresa_grupo = "16" Then
                  var_si = MsgBox("¿Desea cambiar los descuentos del cliente", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     var_si = MsgBox("Confirmar el cambio de descuentos del cliente", vbYesNo, "ATENCION")
                     If var_si = 6 Then
                        rsaux.Open "update tb_gruposactuales set floa_gac_Descuento_1 = " + Me.txt_descuento_1 + ", floa_gac_descuento_2 = " + Me.txt_descuento_2 + " where vcha_gac_grupo_actual_id = '" + Me.txt_clave_grupo + "'", cnn, adOpenDynamic, adLockOptimistic
                        MsgBox "Se han cambiado los descuentos", vbOKOnly, "ATENCION"
                     Else
                        MsgBox "Se a cancelado el cambio de descuentos", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "Se a cancelado el cambio de descuentos", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "El cliente no pertenece a multibondeados", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El cliente no existe", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "No se a seleccionado un cliente", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Descuento 2 incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Descuento 1 incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 2700
   Left = 2800
   Me.frm_lista.Visible = False
   If var_empresa <> "16" Then
      Me.txt_clave_cliente.Enabled = False
      Me.txt_nombre_cliente.Enabled = False
      Me.txt_clave_grupo.Enabled = False
      Me.txt_nombre_grupo.Enabled = False
      Me.txt_descuento_1.Enabled = False
      Me.txt_descuento_2.Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_clave_cliente = Me.lv_lista.selectedItem
      Me.txt_nombre_cliente = Me.lv_lista.selectedItem.SubItems(1)
      Me.txt_clave_grupo = ""
      Me.txt_nombre_grupo = ""
      Me.txt_descuento_1 = ""
      Me.txt_descuento_2 = ""
      Me.txt_clave_cliente.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_clave_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct VCHA_ClI_CLAVE_ID, VCHA_CLI_NOMBRE from VW_CLIENTES where VCHA_EMP_EMPRESA_ID = '16' ORDER BY VCHA_CLI_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CLI_CLAVE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clave_cliente_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_cliente_LostFocus()
   If Me.txt_clave_cliente <> "" Then
      rs.Open "SELECT * FROM VW_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + Me.txt_clave_cliente + "' and vcha_emp_empresa_id = '16'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         Me.txt_clave_grupo = IIf(IsNull(rs!vcha_gac_grupo_Actual_id), "", rs!vcha_gac_grupo_Actual_id)
         Me.txt_nombre_grupo = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
         Me.txt_descuento_1 = IIf(IsNull(rs!floa_gac_descuento_1), 0, rs!floa_gac_descuento_1)
         Me.txt_descuento_2 = IIf(IsNull(rs!floa_gac_descuento_2), 0, rs!floa_gac_descuento_2)
      Else
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
         Me.txt_clave_cliente = ""
         Me.txt_nombre_cliente = ""
         Me.txt_clave_grupo = ""
         Me.txt_nombre_grupo = ""
         Me.txt_descuento_1 = ""
         Me.txt_descuento_2 = ""
      End If
      rs.Close
   Else
      Me.txt_clave_cliente = ""
      Me.txt_clave_grupo = ""
      Me.txt_descuento_1 = ""
      Me.txt_descuento_2 = ""
      Me.txt_nombre_cliente = ""
      Me.txt_nombre_grupo = ""
   End If
End Sub

Private Sub txt_clave_grupo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct VCHA_ClI_CLAVE_ID, VCHA_CLI_NOMBRE from VW_CLIENTES where VCHA_EMP_EMPRESA_ID = '16' ORDER BY VCHA_CLI_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CLI_CLAVE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clave_grupo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_descuento_1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct VCHA_ClI_CLAVE_ID, VCHA_CLI_NOMBRE from VW_CLIENTES where VCHA_EMP_EMPRESA_ID = '16' ORDER BY VCHA_CLI_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CLI_CLAVE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_descuento_1_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_descuento_2_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct VCHA_ClI_CLAVE_ID, VCHA_CLI_NOMBRE from VW_CLIENTES where VCHA_EMP_EMPRESA_ID = '16' ORDER BY VCHA_CLI_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CLI_CLAVE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_descuento_2_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct VCHA_ClI_CLAVE_ID, VCHA_CLI_NOMBRE from VW_CLIENTES where VCHA_EMP_EMPRESA_ID = '16' ORDER BY VCHA_CLI_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CLI_CLAVE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_grupo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct VCHA_ClI_CLAVE_ID, VCHA_CLI_NOMBRE from VW_CLIENTES where VCHA_EMP_EMPRESA_ID = '16' ORDER BY VCHA_CLI_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CLI_CLAVE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_grupo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub
