VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmnotas_credito_no_aplicables 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Notas de crédito no aplicables"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   165
      Picture         =   "frmnotas_credito_no_aplicables.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   495
      Picture         =   "frmnotas_credito_no_aplicables.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7815
      Picture         =   "frmnotas_credito_no_aplicables.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   1695
      Left            =   2235
      TabIndex        =   3
      Top             =   150
      Width           =   4050
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1230
         Left            =   30
         TabIndex        =   4
         Top             =   405
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   2170
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
            Object.Width           =   1464
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5380
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   5
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.Frame frm_lista2 
      Height          =   2190
      Left            =   1425
      TabIndex        =   0
      Top             =   30
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista2 
         Height          =   1710
         Left            =   30
         TabIndex        =   1
         Top             =   435
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3016
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
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label lbl_lista2 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   2
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame3 
      Height          =   90
      Left            =   120
      TabIndex        =   22
      Top             =   285
      Width           =   8085
   End
   Begin VB.Frame Frame2 
      Caption         =   " Datos Generales "
      Height          =   1725
      Left            =   135
      TabIndex        =   6
      Top             =   420
      Width           =   8070
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   585
         Width           =   5190
      End
      Begin VB.TextBox txt_nombre_clase 
         Height          =   315
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   255
         Width           =   5190
      End
      Begin VB.TextBox txt_clase 
         Height          =   315
         Left            =   1050
         TabIndex        =   7
         Top             =   255
         Width           =   1725
      End
      Begin VB.TextBox txt_saldo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1050
         TabIndex        =   11
         Top             =   915
         Width           =   1725
      End
      Begin VB.TextBox txt_clave_cliente 
         Height          =   315
         Left            =   1050
         TabIndex        =   9
         Top             =   585
         Width           =   1725
      End
      Begin VB.TextBox txt_serie 
         Height          =   345
         Left            =   1050
         TabIndex        =   12
         Top             =   1245
         Width           =   810
      End
      Begin VB.TextBox txt_numero 
         Height          =   330
         Left            =   2760
         TabIndex        =   13
         Top             =   1245
         Width           =   1965
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento:"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   165
         TabIndex        =   17
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   975
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   165
         TabIndex        =   15
         Top             =   645
         Width           =   525
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   1995
         TabIndex        =   14
         Top             =   1320
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmnotas_credito_no_aplicables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_serie As String
Dim var_clave_moneda As String
Dim var_tipo_Cambio As Double
Dim var_agente As String
Dim var_grupo_actual As String
Dim var_grupo_real As String
Dim var_titular As String
Dim var_plazo As Integer
Dim var_iva As Double
Dim var_numero_renglones As Double






Private Sub cmb_series_Click()
   var_serie = cmb_series
End Sub

Private Sub cmd_aplicar_Click()
End Sub

Private Sub cmd_imprimir_Click()
   If Trim(Me.txt_clase) <> "" Then
      If Trim(Me.txt_clave_cliente) <> "" Then
         If IsNumeric(Me.txt_numero) Then
            If IsNumeric(Me.txt_saldo) Then
               rs.Open "SELECT * FROM TB_sERIES WHERE VCHA_SER_SERIE_ID = '" + Me.txt_serie + "'", cnn, adOpenDynamic, adLockOptimistic
               If rs.EOF Then
                  rsaux1.Open "select * from tb_encabezado_cartera where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_car_tipo_documento = 'NC' and vcha_car_documento = '" + Me.txt_clase.Text + "' and vcha_ser_serie_id = '" + Me.txt_serie + "' and inte_Car_numero = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                  If rsaux1.EOF Then
                     rsaux.Open "exec SP_NOTA_CREDITO_SIN_APLICAR '" + var_empresa + "', '" + Me.txt_clave_cliente + "','" + Me.txt_clase + "'," + Me.txt_numero + ", '" + Me.txt_serie + "', " + Me.txt_saldo + ",1,'" + var_clave_usuario_global + "', '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     MsgBox "La nota de crédito ya existe", vbOKOnly, "ATENCION"
                  End If
                  rsaux1.Close
               Else
                  MsgBox "La serie no puede ser " + Me.txt_serie, vbOKOnly, "ATENCION"
               End If
               rs.Close
            Else
               MsgBox "El importe es incorrecto", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Número de documento incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Tipo de movimiento incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   txt_clave_cliente.Enabled = True
   cmd_imprimir.Enabled = True
   txt_clave_cliente = ""
   txt_nombre_cliente = ""
   txt_saldo = ""
   txt_total_aplicado = ""
   txt_saldo.Enabled = True
   If Me.txt_clase.Enabled = True Then
      txt_clase.SetFocus
   Else
      txt_clave_cliente.SetFocus
   End If
   
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
End Sub

Private Sub Form_Load()
   Me.Top = 2500
   Me.Left = 1800
   var_cadena_seguridad = ""
   frm_lista2.Visible = False
   txt_clave_cliente.Enabled = True
   frm_lista.Visible = False
   txt_saldo.Enabled = True
   Top = 0
   Left = 0
   rs.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'"
   var_numero_renglones = rs!INTE_PRI_RENGLONES_NOTA_CREDITO
   rs.Close
   rs.Open "select * from tb_clases_Cartera where vcha_car_documento = 'BO' order by vcha_car_nombre ", cnn, adOpenDynamic, adLockBatchOptimistic
   If Not rs.EOF Then
      var_contador_movimiento = 0
      While Not rs.EOF
         var_contador_movimiento = var_contador_movimiento + 1
         rs.MoveNext
      Wend
      
      If var_contador_movimiento > 1 Then
         txt_nombre_clase.Enabled = True
         txt_clase.Enabled = True
      Else
         txt_nombre_clase.Enabled = False
         txt_clase.Enabled = False
      End If
      rs.MoveFirst
      txt_nombre_clase = rs!VCHA_CAR_NOMBRE
      txt_clase = rs!VCHA_CAR_CLASE_ID
   Else
      MsgBox "No se a indicado una clase de Bonificación", vbOKOnly, "ATENCION"
      txt_clase.Enabled = False
      cmb_clases.Enabled = False
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_facturas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_facturas, ColumnHeader)
End Sub

Private Sub lv_facturas_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F4 para indicar el importe a aplicar a la factura"
End Sub

Private Sub lv_facturas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 115 Then
      If Trim(txt_saldo) = "" Then
         txt_saldo = 0
      End If
      If txt_saldo > 0 Then
         frm_cantidad_aplicar.Visible = True
         txt_cantidad_aplicar = ""
         txt_cantidad_aplicar.SetFocus
      Else
         MsgBox "No se a indicado el importe de la bonificación", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub lv_facturas_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_clase = lv_lista.selectedItem
         txt_nombre_clase = lv_lista.selectedItem.SubItems(1)
      Else
         txt_clase = ""
         txt_nombre_clase = ""
      End If
      frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_lista2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista2, ColumnHeader)
End Sub

Private Sub lv_lista2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista2.ListItems.Count > 0 Then
         txt_clave_cliente = lv_lista2.selectedItem
         txt_nombre_cliente = lv_lista2.selectedItem.SubItems(1)
      Else
         txt_clave_cliente = ""
         txt_nombre_cliente = ""
      End If
      txt_clave_cliente.SetFocus
      frm_lista2.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista2.Visible = False
   End If
End Sub

Private Sub lv_lista2_LostFocus()
   frm_lista2.Visible = False
End Sub

Private Sub txt_cantidad_aplicar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If IsNumeric(txt_cantidad_aplicar) Then
         If (txt_cantidad_aplicar * 1) + (lv_facturas.selectedItem.SubItems(8) * 1) > (lv_facturas.selectedItem.SubItems(7) * 1) Then
            MsgBox "La cantidad a aplicar exede el importe del saldo de la factura", vbOKOnly, "ATENCIO"
         Else
            If (txt_total_aplicado * 1) + (txt_cantidad_aplicar * 1) <= (txt_saldo * 1) Then
               lv_facturas.selectedItem.SubItems(8) = Format(txt_cantidad_aplicar + (lv_facturas.selectedItem.SubItems(8) * 1), "###,##0.00")
               txt_total_aplicado = (txt_total_aplicado * 1) + (txt_cantidad_aplicar * 1)
            Else
               MsgBox "La cantidad a aplicar exede al importe del saldo del pago del cliente", vbOKOnly, "ATENCION"
            End If
         End If
      Else
         MsgBox "Importe Incorrecto", vbOKOnly, "ATENCION"
      End If
      frm_cantidad_aplicar.Visible = False
      lv_facturas.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_cantidad_aplicar.Visible = False
      If lv_facturas.ListItems.Count > 0 Then
         lv_facturas.SetFocus
      End If
   End If
End Sub

Private Sub txt_clase_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clase_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      Dim list_item As ListItem
      Dim var_contador_lista As Integer
      rs.Open "select vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where vcha_car_documento= 'BO' order by vcha_Car_nombre", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            var_contador_lista = var_contador_lista + 1
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CAR_CLASE_ID)
            list_item.SubItems(1) = Trim(IIf(IsNull(rs!VCHA_CAR_NOMBRE), "", rs!VCHA_CAR_NOMBRE))
            rs.MoveNext
         Wend
      End If
      rs.Close
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clase_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clase_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_clave_cliente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clave_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista2.ListItems.Clear
      rs.Open "select * from vw_clientes WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_cli_nombre ", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista2.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista2 = "CLIENTES"
      var_tipo_lista = 1
      frm_lista2.Visible = True
      lv_lista2.SetFocus
   End If
End Sub

Private Sub txt_clave_cliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_cliente = rs!vcha_cli_nombre
         rs.Close
      Else
         rs.Close
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
         txt_clave_cliente = ""
         txt_nombre_cliente = ""
      End If
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_cliente_LostFocus()
   Dim var_importe_factura As Double
   Dim var_importe_pago As Double
   Dim var_importe_saldo_pago As Double
   Dim var_importe_total As Double
   Dim var_fecha_pago As Date
   Dim var_fecha_factura As Date
   Dim var_contador_pagos As Double
   Dim var_contador_facturas As Double
   Dim var_descuento_agente As Double
   Dim var_descuento_sistema As Double
   Dim var_saldo As Double
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_clave_cliente) <> "" Then
      rs.Open "select * from vw_clientes where vcha_cli_clave_id ='" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_cliente = rs!vcha_cli_nombre
         var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
         var_agente = IIf(IsNull(rs!vcha_Age_agente_id), "", rs!vcha_Age_agente_id)
         var_grupo_actual = IIf(IsNull(rs!vcha_gac_grupo_Actual_id), "", rs!vcha_gac_grupo_Actual_id)
         var_grupo_real = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
         var_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
         var_plazo = IIf(IsNull(rs!inte_pla_dias), 0, rs!inte_pla_dias)
         var_iva = IIf(IsNull(rs!FLOA_TPE_IVA), 0, rs!FLOA_TPE_IVA)
         rs.Close
      Else
         rs.Close
         txt_clave_cliente = ""
         txt_nombre_cliente = ""
         txt_saldo = ""
         txt_total_aplicado = Format(0, "###,##0.00")
         lv_facturas.ListItems.Clear
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_descuento_Change()

End Sub

Private Sub txt_descuento_KeyPress(KeyAscii As Integer)
End Sub

Private Sub txt_nombre_clase_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      Dim list_item As ListItem
      Dim var_contador_lista As Integer
      rs.Open "select vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where vcha_car_documento= 'DV' order by vcha_Car_nombre", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            var_contador_lista = var_contador_lista + 1
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CAR_CLASE_ID)
            list_item.SubItems(1) = Trim(IIf(IsNull(rs!VCHA_CAR_NOMBRE), "", rs!VCHA_CAR_NOMBRE))
            rs.MoveNext
         Wend
      End If
      rs.Close
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_clase_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista2.ListItems.Clear
      rs.Open "select * from vw_clientes WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_cli_nombre ", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista2.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista2 = "CLIENTES"
      var_tipo_lista = 1
      frm_lista2.Visible = True
      lv_lista2.SetFocus
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub txt_saldo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

