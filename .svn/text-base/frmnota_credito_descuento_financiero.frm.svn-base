VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmnota_credito_descuento_financiero 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota de Crédito por Descuento Financiero"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   8490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_nota_credito_electronica 
      Appearance      =   0  'Flat
      Caption         =   "NC Electronica"
      Enabled         =   0   'False
      Height          =   315
      Left            =   825
      Picture         =   "frmnota_credito_descuento_financiero.frx":0000
      TabIndex        =   26
      Top             =   30
      Width           =   1485
   End
   Begin VB.Frame frm_lista 
      Height          =   2565
      Left            =   1290
      TabIndex        =   23
      Top             =   2040
      Width           =   6390
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2100
         Left            =   45
         TabIndex        =   24
         Top             =   405
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   3704
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
            Text            =   "Descripción"
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   25
         Top             =   120
         Width           =   6315
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Notas de Crédito "
      Height          =   630
      Left            =   105
      TabIndex        =   17
      Top             =   1695
      Width           =   8205
      Begin VB.TextBox txt_a 
         Height          =   315
         Left            =   4350
         TabIndex        =   21
         Top             =   225
         Width           =   1320
      End
      Begin VB.TextBox txt_de 
         Height          =   315
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   225
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   195
         Left            =   4155
         TabIndex        =   20
         Top             =   285
         Width           =   150
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   2040
         TabIndex        =   19
         Top             =   285
         Width           =   255
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7950
      Picture         =   "frmnota_credito_descuento_financiero.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   " Importes de descuentos no aplicados "
      Height          =   4785
      Left            =   120
      TabIndex        =   9
      Top             =   2340
      Width           =   8205
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6195
         TabIndex        =   14
         Top             =   4335
         Width           =   1845
      End
      Begin MSComctlLib.ListView lv_importes 
         Height          =   4095
         Left            =   90
         TabIndex        =   10
         Top             =   240
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   7223
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
         NumItems        =   23
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Factura"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Moneda "
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Subimporte      "
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Descuento"
            Object.Width           =   1799
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Importe       "
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Almacen"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Grupo Actual"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Grupo Real"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Titular"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Establecimiento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "IVA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Moneda"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Tipo Cambio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Serie"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Saldo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Descuento Aplicado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Aplicado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Cheque"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Partida"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Documento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Linea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "Saldo total"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   5205
         TabIndex        =   15
         Top             =   4425
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos Generales "
      Height          =   1245
      Left            =   105
      TabIndex        =   3
      Top             =   435
      Width           =   8205
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   3435
         TabIndex        =   22
         Top             =   840
         Width           =   4590
      End
      Begin VB.TextBox txt_relacion_cobranza 
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   180
         Width           =   1770
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   315
         Left            =   3435
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   510
         Width           =   4590
      End
      Begin VB.ComboBox cmb_clientes 
         Height          =   315
         Left            =   3435
         TabIndex        =   8
         Top             =   180
         Visible         =   0   'False
         Width           =   4605
      End
      Begin VB.TextBox txt_clave_cliente 
         Height          =   315
         Left            =   2100
         TabIndex        =   6
         Top             =   840
         Width           =   1320
      End
      Begin VB.TextBox txt_clave_agente 
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   510
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Relación de Cobranza:"
         Height          =   195
         Left            =   375
         TabIndex        =   13
         Top             =   240
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   375
         TabIndex        =   7
         Top             =   900
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   390
         TabIndex        =   5
         Top             =   570
         Width           =   555
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   495
      Picture         =   "frmnota_credito_descuento_financiero.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   165
      Picture         =   "frmnota_credito_descuento_financiero.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   90
      Left            =   0
      TabIndex        =   2
      Top             =   300
      Width           =   8400
   End
End
Attribute VB_Name = "frmnota_credito_descuento_financiero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_numero_notas As Integer
Dim var_numero_renglones As Integer
Dim var_numero_nota As Double
Dim var_numero_nota_anterior As Double
Dim var_serie As String
Dim var_tolerancia_saldo As Double
Private Sub cmb_clientes_Click()
   txt_clave_cliente = Obtener_llave(cnn, rs, "TB_CLIENTES", "VCHA_CLI_NOMBRE", cmb_clientes, 0, "T")
End Sub

Private Sub cmb_clientes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim list_item As ListItem
      Dim var_importe As Double
      Dim var_descuento As Double
      Dim var_importe_total As Double
      Dim var_total As Double
      Dim var_contador As Double
      Dim var_contador_notas As Double
      rs.Open "select distinct vcha_cli_nombre from VW_NOTA_CREDITO_RELACION_COBRANZA with (nolock) where vcha_rco_folio = '" + txt_relacion_cobranza + "' and vcha_cli_clave_ID = '" + txt_clave_cliente + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_clave_cliente.Enabled = False
         cmb_clientes.Enabled = False
         rsaux2.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'"
         var_tolerancia_saldo = rsaux2!FLOA_PRI_TOLERANCIA_SALDOS
         rsaux2.Close
         rsaux2.Open "select * from VW_NOTA_CREDITO_RELACION_COBRANZA where vcha_rco_folio = '" + txt_relacion_cobranza + "' and vcha_cli_clave_ID = '" + txt_clave_cliente + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
         var_serie = rsaux2!vcha_Ser_Serie_id
         var_numero_notas = rsaux2.RecordCount
         rsaux3.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
         var_serie = IIf(IsNull(rsaux3!vcha_Ser_Serie_id), "", rsaux3!vcha_Ser_Serie_id)
         If var_serie = "" Then
            If var_empresa = "02" Then
               var_serie = "XM"
            End If
            If var_empresa = "03" Then
               var_serie = "SI"
            End If
         End If
         rsaux3.Close
         rsaux3.Open "select inte_ser_nota_credito from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
         var_numero_nota = (IIf(IsNull(rsaux3!inte_ser_nota_credito), 1, rsaux3!inte_ser_nota_credito))
         var_numero_nota_anterior = (IIf(IsNull(rsaux3!inte_ser_nota_credito), 1, rsaux3!inte_ser_nota_credito))
         rsaux3.Close
         var_total = 0
         txt_de = var_numero_nota
         var_contador = 0
         var_contador_notas = 0
         While Not rsaux2.EOF
            Set list_item = lv_importes.ListItems.Add(, , rsaux2!inte_Car_numero)
            list_item.SubItems(1) = Format(IIf(IsNull(rsaux2!dtim_Car_fecha), "", rsaux2!dtim_Car_fecha), "Short Date")
            list_item.SubItems(2) = IIf(IsNull(rsaux2!vcha_mon_nombre_plural), "", rsaux2!vcha_mon_nombre_plural)
            var_importe = 0
            var_descuento = 0
            var_importe = (IIf(IsNull(rsaux2!floa_rco_importe), 0, rsaux2!floa_rco_importe)) * (1 + var_descuento / 100)
            var_descuento = IIf(IsNull(rsaux2!floa_Rco_descuento_aplicar), "", rsaux2!floa_Rco_descuento_aplicar)
            list_item.SubItems(3) = Format((var_importe / ((100 - var_descuento) / 100)), "###,##0.00")
            list_item.SubItems(4) = Format(IIf(IsNull(rsaux2!floa_Rco_descuento_recomendado), 0, rsaux2!floa_Rco_descuento_recomendado), "###,##0.00")
            var_importe_total = ((var_importe / ((100 - var_descuento) / 100)) * (var_descuento / 100))
            list_item.SubItems(5) = Format(var_importe_total, "###,##0.00")
            list_item.SubItems(6) = IIf(IsNull(rsaux2!VCHA_ALM_ALMACEN_ID), "", rsaux2!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(7) = IIf(IsNull(rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID)
            list_item.SubItems(8) = IIf(IsNull(rsaux2!vcha_gre_grupo_real_id), "", rsaux2!vcha_gre_grupo_real_id)
            list_item.SubItems(9) = IIf(IsNull(rsaux2!vcha_tit_titular_id), "", rsaux2!vcha_tit_titular_id)
            list_item.SubItems(10) = IIf(IsNull(rsaux2!vcha_ESB_ESTABLECIMIENTO_id), "", rsaux2!vcha_ESB_ESTABLECIMIENTO_id)
            list_item.SubItems(11) = IIf(IsNull(rsaux2!floa_rco_iva), 0, rsaux2!floa_rco_iva)
            list_item.SubItems(12) = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
            list_item.SubItems(13) = IIf(IsNull(rsaux2!floa_rco_tipo_cambio), 1, rsaux2!floa_rco_tipo_cambio)
            list_item.SubItems(14) = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), "", rsaux2!vcha_Ser_Serie_id)
            rsaux3.Open "select floa_sal_importe from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_car_numero = " + Str(rsaux2!inte_Car_numero) + " and vcha_ser_serie_id = '" + rsaux2!vcha_Ser_Serie_id + "'", cnn, adOpenDynamic, adLockOptimistic
            list_item.SubItems(15) = Format(IIf(IsNull(rsaux3!FLOA_sAL_IMPORTE), 0, rsaux3!FLOA_sAL_IMPORTE) - var_importe_total, "###,##0.00")
            list_item.SubItems(16) = Format(IIf(IsNull(rsaux2!floa_Rco_descuento_aplicar), 0, rsaux2!floa_Rco_descuento_aplicar), "###,##0.00")
            list_item.SubItems(18) = IIf(IsNull(rsaux2!VCHA_rCO_CHEQUE), "", rsaux2!VCHA_rCO_CHEQUE)
            list_item.SubItems(19) = IIf(IsNull(rsaux2!inte_rco_partida), 0, rsaux2!inte_rco_partida)
            rsaux3.Close
            var_tipo_Cambio = (lv_importes.selectedItem.SubItems(13) * 1)
            var_saldo = lv_importes.selectedItem.SubItems(15) * 1
            var_tolerancia_saldo = var_tolerancia_saldo / var_tipo_Cambio
            var_importe = 0
            'If var_saldo < var_tolerancia_saldo Then
            '   var_importe = var_importe_total + var_saldo
            'Else
            '   var_importe = var_importe_total * 1
            'End If
            var_importe_total = var_importe
            list_item.SubItems(5) = Format(var_importe, "###,##0.00")
            var_total = var_total + var_importe_total
            var_contador = var_contador + 1
            If var_contador > var_numero_renglones Then
               var_contador_notas = var_contador_notas + 1
               var_contador = 0
            End If
            rsaux2.MoveNext
         Wend
         txt_a = txt_de + var_contador_notas
         txt_importe = Format(var_total, "###,##0.00")
         rsaux2.Close
      Else
         MsgBox "El cliente no tiene descuentos por aplicar o no se encuentra en la relación de cobranza", vbOKOnly, "ATENCION"
         txt_clave_cliente = ""
         cmb_clientes = ""
      End If
      rs.Close
   End If
End Sub

Private Sub cmd_imprimir_Click()
   Dim var_descuento_otorgado As Double
   Dim var_descuento_aplicado As Double
   Dim var_almacen As String
   Dim var_grupo_actual As String
   Dim var_grupo_real As String
   Dim var_cliente As String
   Dim var_titular As String
   Dim var_establecimiento As String
   Dim var_clave_moneda As String
   Dim var_agente As String
   Dim var_imprimir As Boolean
   Dim var_contador As Integer
   Dim var_contador_notas As Integer
   Dim var_contador_lineas As Integer
   Dim var_tipo_Cambio As Double
   Dim var_iva As Double
   Dim var_importe_total As Double
   Dim var_importe_iva As Double
   Dim var_subimporte As Double
   Dim var_importe As Double
   Dim si, i, n As Integer
   Dim var_saldo As Double
   Dim var_serie_cargo As String
   Dim var_numero_nota_inicio As Double
   Dim var_factura As Double
   Dim var_k As Double
   Dim var_descuentos As Double
   Dim var_desc_otorgado_str As String
   Dim var_desc_apilcado_str As String
   var_iva_pasado = 0
   var_posible_iva = 1
   For var_j = 1 To Me.lv_importes.ListItems.Count
       Me.lv_importes.ListItems.item(var_j).Selected = True
       If var_iva_pasado = 0 Then
          var_iva_pasado = CDbl(Me.lv_importes.selectedItem.SubItems(11))
       Else
          If var_iva_pasado <> CDbl(Me.lv_importes.selectedItem.SubItems(11)) Then
             var_posible_iva = 0
          End If
       End If
   Next var_j
   If var_posible_iva = 1 Then
   var_iva = var_iva_pasado
   
   
   Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
   Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
   var_descuentos = lv_importes.ListItems.Count
   If CDbl(Me.txt_importe) <= 0 Then
      MsgBox "No es posible imprimir el importe de esta nota de crédito", vbOKOnly, "ATENCION"
   Else
   If var_descuentos > 0 Then
      si = MsgBox("¿Deseas Imprimir las Notas de Crédito", vbYesNo, "ATENCION")
      If si = 6 Then
         si = MsgBox("Confirmar la impresión de las Notas de Crédito", vbYesNo, "ATENCION")
         If si = 6 Then
            cnn.BeginTrans
            rs.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            var_tolerancia_saldo = rs!FLOA_PRI_TOLERANCIA_SALDOS
            rs.Close
            var_almacen = lv_importes.selectedItem.SubItems(6)
            var_grupo_actual = lv_importes.selectedItem.SubItems(7)
            var_grupo_real = lv_importes.selectedItem.SubItems(8)
            var_titular = lv_importes.selectedItem.SubItems(9)
            var_agente = txt_clave_agente
            var_cliente = txt_clave_cliente
            var_establecimiento = lv_importes.selectedItem.SubItems(10)
            'var_iva = (lv_importes.selectedItem.SubItems(11) * 1)
            'var_iva = 15
            var_clave_moneda = lv_importes.selectedItem.SubItems(12)
            var_tipo_Cambio = (lv_importes.selectedItem.SubItems(13) * 1)
            rs.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            var_serie = IIf(IsNull(rs!vcha_Ser_Serie_id), "", rs!vcha_Ser_Serie_id)
            If var_serie = "" Then
               If var_empresa = "02" Then
                  var_serie = "XM"
               End If
               If var_empresa = "03" Then
                  var_serie = "SI"
               End If
            End If
            rs.Close
            var_insertar = False
            n = lv_importes.ListItems.Count
            var_imprimir = False
            var_contador = 0
            var_contador_notas = 0
            var_tolerancia_saldo = var_tolerancia_saldo / var_tipo_Cambio
            var_numero_nota_inicio = var_numero_nota
            For i = 1 To n
               lv_importes.ListItems.item(i).Selected = True
               var_descuento_otorgado = lv_importes.selectedItem.SubItems(4) * 1
               var_descuento_aplicado = lv_importes.selectedItem.SubItems(16) * 1
               var_saldo = lv_importes.selectedItem.SubItems(15) * 1
               If var_saldo < var_tolerancia_saldo Then
                  'var_importe = var_importe + (lv_importes.selectedItem.SubItems(5) * 1) + var_saldo
                  var_importe = var_importe + (lv_importes.selectedItem.SubItems(5) * 1)
               Else
                  var_importe = var_importe + (lv_importes.selectedItem.SubItems(5) * 1)
               End If
               var_contador = var_contador + 1
               If (var_contador = var_numero_renglones) Or (i = n) Then
                  var_contador = 0
                  var_imprimir = True
               End If
               'MsgBox "update tb_relacion_cobranza set INTE_RCO_NUMERO_DESCUENTO_FINANCIERO = " + Str(var_numero_nota) + ", DTIM_RCO_FECHA_DESCUENTO_FINANCIERO = getdate() where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + txt_relacion_cobranza + "' and inte_car_numero = " + lv_importes.selectedItem + " and vcha_rco_cheque = '" + lv_importes.selectedItem.SubItems(18) + "' and inte_rco_partida = " + lv_importes.selectedItem.SubItems(19)
               rs.Open "update tb_relacion_cobranza set INTE_RCO_NUMERO_DESCUENTO_FINANCIERO = " + Str(var_numero_nota) + ", DTIM_RCO_FECHA_DESCUENTO_FINANCIERO = getdate() where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + txt_relacion_cobranza + "' and inte_car_numero = " + lv_importes.selectedItem + " and vcha_rco_cheque = '" + lv_importes.selectedItem.SubItems(18) + "' and inte_rco_partida = " + lv_importes.selectedItem.SubItems(19), cnn, adOpenDynamic, adLockOptimistic
               If var_imprimir = True Then
                  var_subimporte = var_importe / (1 + (var_iva / 100))
                  var_importe_iva = var_importe - var_subimporte
                  var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "NC", "DF", "DF", var_numero_nota, "-", var_almacen, "", 0, Date, var_agente, var_grupo_actual, var_grupo_real, var_titular, var_cliente, var_establecimiento, 0, var_iva, 0, 0, 0, 0, 0, var_importe * var_tipo_Cambio, var_importe_iva * var_tipo_Cambio, 0, 0, 0, 0, 0, var_subimporte * var_tipo_Cambio, var_importe * var_tipo_Cambio, "", var_clave_usuario_global, "", Date, 0, Date, Date, var_clave_moneda, var_tipo_Cambio, var_serie, "")
                  rsaux3.Open "insert into tb_secuencia_notas_credito (vcha_emp_empresa_id, vcha_Ser_serie_id, inte_snc_numero_anterior, inte_snc_numero_actual) values ('" + var_empresa + "', '" + var_serie + "', " + CStr(var_numero_nota) + ", " + CStr(var_numero_nota) + ")", cnn, adOpenDynamic, adLockOptimistic
                  var_numero_nota = var_numero_nota + 1
                  var_contador_notas = var_contador_notas + 1
                  var_importe = 0
                  var_subimporte = 0
                  var_importe_iva = 0
               End If
               var_imprimir = False
            Next i
            var_numero_nota = var_numero_nota_anterior
            var_importe = 0
            For i = 1 To n
               lv_importes.ListItems.item(i).Selected = True
               var_saldo = lv_importes.selectedItem.SubItems(15) * 1
               var_serie_cargo = lv_importes.selectedItem.SubItems(14)
               If var_saldo < var_tolerancia_saldo Then
                  'var_importe = (lv_importes.selectedItem.SubItems(5) * 1) + var_saldo
                  var_importe = (lv_importes.selectedItem.SubItems(5) * 1)
               Else
                  var_importe = (lv_importes.selectedItem.SubItems(5) * 1)
               End If
               var_contador = var_contador + 1
               If (var_contador = var_numero_renglones) Or (i = n) Then
                  var_contador = 0
                  var_imprimir = True
               End If
               var_factura = lv_importes.selectedItem
               var_descuento_otorgado = lv_importes.selectedItem.SubItems(4) * 1
               var_descuento_aplicado = lv_importes.selectedItem.SubItems(16) * 1
               'var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, var_serie_cargo, "FA", lv_importes.selectedItem, var_serie, "DF", var_numero_nota, 0, (var_importe * var_tipo_Cambio))
               var_cadena = "insert into tb_estado_cuenta (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_ecu_movimiento_cargo, vcha_ecu_serie_cargo, inte_Ecu_numero_cargo, floa_ecu_importe_cargo, vcha_ecu_movimiento_abono, vcha_ecu_serie_abono, inte_ecu_numero_abono, floa_Ecu_importe_abono) "
               var_cadena = var_cadena + " values ('" + var_empresa + "','" + var_unidad_organizacional + "','FA','" + var_serie_cargo + "'," + Trim(CStr(CDbl(lv_importes.selectedItem))) + ",0,'DF','" + var_serie + "'," + CStr(var_numero_nota) + "," + CStr(Round((var_importe * var_tipo_Cambio), 2)) + ")"
               rsaux8.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               'rsaux9.Open "select * from tb_estado_cuenta where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ecu_movimiento_cargo = 'FA' and vcha_ecu_Serie_cargo = '" + var_serie_cargo + "' and inte_Ecu_numero_cargo = " + Trim(CStr(CDbl(lv_importes.selectedItem))) + " and vcha_ecu_movimiento_abono = 'DF' and vcha_ecu_serie_abono = '" + var_serie + "' and inte_Ecu_numero_abono = " + CStr(var_numero_nota), cnn, adOpenDynamic, adLockOptimistic
               'If rsaux9.EOF Then
               '   rsaux10.Open "select floa_sal_importe from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_Serie_id = '" + var_serie_cargo + "' and inte_Car_numero = " + Trim(CStr(CDbl(lv_importes.selectedItem))), cnn, adOpenDynamic, adLockOptimistic
               '   If Not rsaux10.EOF Then
               '      var_saldo_factura = rsaux10(0).Value
               '   End If
               '   rsaux10.Close
               '   var_diferencia_saldo_pago = var_saldo_factura - Round((var_importe * var_tipo_Cambio), 2)
               '   If var_diferencia_saldo_pago = -0.01 Then
               '      var_importe = var_importe - 0.01
               '   End If
               '   var_cadena = "insert into tb_estado_cuenta (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_ecu_movimiento_cargo, vcha_ecu_serie_cargo, inte_Ecu_numero_cargo, floa_ecu_importe_cargo, vcha_ecu_movimiento_abono, vcha_ecu_serie_abono, inte_ecu_numero_abono, floa_Ecu_importe_abono) "
               '   var_cadena = var_cadena + " values ('" + var_empresa + "','" + var_unidad_organizacional + "','FA','" + var_serie_cargo + "'," + Trim(CStr(CDbl(lv_importes.selectedItem))) + ",0,'DF','" + var_serie + "'," + CStr(var_numero_nota) + "," + CStr(Round((var_importe * var_tipo_Cambio), 2)) + ")"
               '   rsaux8.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               'End If
               'rsaux9.Close
               rsaux3.Open "Insert into TB_DETALLE_DESCUENTOS_FINANCIEROS (vcha_emp_empresa_id, vcha_car_documento, vcha_ser_serie_id, vcha_car_clase_id, inte_car_numero, vcha_ddf_concepto, floa_ddf_importe, inte_ddf_factura, floa_ddf_iva, floa_ddf_descuento_otorgado, floa_ddf_descuento_aplicado) values ('" + var_empresa + "', 'DF', '" + var_serie + "','DF'," + Str(var_numero_nota) + ",'', " + Str((var_importe * var_tipo_Cambio)) + ", " + Str(var_factura) + ", " + CStr(var_iva) + ", " + CStr(var_descuento_otorgado) + ", " + CStr(var_descuento_aplicado) + " )", cnn, adOpenDynamic, adLockOptimistic
               If var_imprimir = True Then
                  var_subimporte = var_importe / (1 + (var_iva / 100))
                  var_importe_iva = var_importe - var_subimporte
                  var_numero_nota = var_numero_nota + 1
                  var_importe = 0
                  var_importe_iva = 0
               End If
               var_imprimir = False
            Next i
            rs.Open "update tb_series set inte_ser_nota_credito = inte_ser_nota_credito + " + Str(var_contador_notas) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            If var_empresa = "02" Or var_empresa = "03" Then
'''''''''''''' se imprime la nota de credito
               For var_k = var_numero_nota_inicio To var_numero_nota
                   rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DF' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_k), cnn, adOpenDynamic, adLockOptimistic
                   If Not rs.EOF Then
                      Open (App.Path & "\nc" + Trim(Str(rs!inte_Car_numero)) + ".txt") For Output As #1
                      'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                      Print #1, Chr(27) + Chr(64)
                      Print #1, Spc(92); Str(rs!inte_Car_numero)
                      Print #1, ""
                      Print #1, Spc(92); "       "; Format(rs!dtim_Car_fecha, "Short Date")
                      var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                      For var_j = 1 + Len(Trim(var_cliente)) To 83
                           var_cliente = var_cliente + " "
                      Next var_j
                      var_cliente = var_cliente + "AGUASCALIENTES, AGS."
                      Print #1, ""
                      Print #1, Spc(12); var_cliente
                      var_domicilio = Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION)) + " COL.: " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                      For var_j = 1 + Len(Trim(var_domicilio)) To 83
                           var_domicilio = var_domicilio + " "
                      Next var_j
                      var_agente = ""
                      var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                      For var_j = 1 + Len(Trim(var_agente)) To 8
                          var_agente = var_agente + " "
                      Next var_j
                      var_agente = var_agente + Mid(IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE), 1, 30)
                      var_domicilio = var_domicilio
                      Print #1, Spc(12); var_domicilio
                      var_ciudad = ""
                      var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                      For var_j = 1 + Len(Trim(var_ciudad)) To 37
                          var_ciudad = var_ciudad + " "
                      Next var_j
                      var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                      For var_j = 1 + Len(Trim(var_estado)) To 46
                          var_estado = var_estado + " "
                      Next var_j
                      var_ciudad = var_ciudad + var_estado
                                 
                      For var_j = 1 + Len(Trim(var_ciudad)) To 14
                           var_ciudad = var_ciudad + " "
                      Next var_j
                           
                      var_ciudad = var_ciudad + var_agente
                               
                      Print #1, Spc(12); var_ciudad
                      var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                      var_rfc = "      " + var_rfc
                      For var_j = 1 + Len(Trim(var_rfc)) To 89
                          var_rfc = var_rfc + " "
                      Next var_j
                      var_rfc = var_rfc
                      Print #1, Spc(6); var_rfc
                      Print #1, ""
                      Print #1, ""
                      Print #1, ""
                      Print #1, ""
                      rsaux3.Open "select * from TB_DETALLE_DESCUENTOS_FINANCIEROS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DF' and vcha_ser_serie_id  = '" + var_serie + "' and vcha_car_clase_id = 'DF' and inte_car_numero =  " + Str(var_k), cnn, adOpenDynamic, adLockOptimistic
                      var_contador_lineas = 0
                      var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                      While Not rsaux3.EOF
                         'var_linea = "DF" + Str(rs!inte_car_numero) + " " + rs!vcha_Car_nombre + " Factura " + Str(rsaux3!inte_ddf_factura) + " " + CStr(rsaux3!floa_ddf_descuento_otorgado) + "%"
                         var_linea = "DF" + Str(rs!inte_Car_numero) + " " + rs!vcha_Car_nombre + " Factura " + Str(rsaux3!inte_ddf_factura)
                         'If Round(rsaux3!floa_ddf_descuento_otorgado, 2) <> Round(rsaux3!floa_ddf_descuento_aplicado, 2) Then
                         '   var_linea = var_linea + " (" + Format(rsaux3!floa_ddf_descuento_aplicado, "###,###,##0.0000") + "%)"
                         'End If
                         If Len(Trim(var_linea)) < 120 Then
                            For var_j = 1 + Len(Trim(var_linea)) To 120
                                var_linea = var_linea + " "
                            Next var_j
                         End If
                         If Len(Trim(var_rfc)) = 0 Then
                            var_importe_str = Format((IIf(IsNull(rsaux3!FLOA_DDF_IMPORTE), 0, rsaux3!FLOA_DDF_IMPORTE)), "###,###,##0.00")
                            If Len(Trim(var_importe_str)) < 14 Then
                               For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                   var_importe_str = " " + var_importe_str
                               Next var_j
                            End If
                         Else
                            var_importe_str = Format((IIf(IsNull(rsaux3!FLOA_DDF_IMPORTE), 0, rsaux3!FLOA_DDF_IMPORTE) / (1 + (var_iva / 100))), "###,###,##0.00")
                            If Len(Trim(var_importe_str)) < 14 Then
                               For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                   var_importe_str = " " + var_importe_str
                               Next var_j
                            End If
                         End If
                         var_linea = var_linea + var_importe_str
                         Print #1, var_linea
                         rsaux3.MoveNext
                         var_contador_lineas = var_contador_lineas + 1
                      Wend
                      If var_contador_lineas < var_numero_renglones Then
                         For var_l = var_contador_lineas To var_numero_renglones
                             Print #1, ""
                         Next var_l
                      End If
                      rsaux3.Close
                      var_cantidad_letra = rs!vcha_car_importe_letra
                      var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                      If Len(Trim(var_linea)) < 105 Then
                         For var_j = 1 + Len(Trim(var_linea)) To 105
                             var_linea = var_linea + " "
                         Next var_j
                      End If
                      If Len(Trim(var_rfc)) = 0 Then
                         var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                         If Len(Trim(var_subimporte_str)) < 14 Then
                            For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                var_subimporte_str = " " + var_subimporte_str
                            Next var_j
                         End If
                         '1
                         var_iva_str = "-"
                         For var_j = 1 + Len(Trim(var_iva_str)) To 14
                             var_iva_str = " " + var_iva_str
                         Next var_j
                      Else
                         var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                         If Len(Trim(var_subimporte_str)) < 14 Then
                            For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                var_subimporte_str = " " + var_subimporte_str
                            Next var_j
                         End If
                         var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                         If Len(Trim(var_iva_str)) < 14 Then
                            For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                var_iva_str = " " + var_iva_str
                            Next var_j
                         End If
                      End If
                      var_linea = var_linea + "           " + var_subimporte_str
                      Print #1, Spc(4); var_linea
                      Print #1, Spc(120); var_iva_str
                      var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                      If Len(Trim(var_importe_str)) < 14 Then
                         For var_j = 1 + Len(Trim(var_importe_str)) To 14
                             var_importe_str = " " + var_importe_str
                         Next var_j
                      End If
                      Print #1, Spc(120); var_importe_str
                      Print #1, ""
                      Print #1, ""
                      Print #1, ""
                      Print #1, Spc(85); "SISTEMAS"
                      Print #1, ""
                      Print #1, ""
                      Print #1, ""
                      Print #1, ""
                      Print #1, ""
                      Print #1, ""
                      Print #1, ""
                      Close #1
                      
                      Open (App.Path & "\nc" + Trim(Str(rs!inte_Car_numero)) + ".bat") For Output As #2
                      var_Archivo = App.Path & "\nc" + Trim(Str(rs!inte_Car_numero)) + ".bat"
                      Print #2, "copy " + App.Path & "\nc" + Trim(Str(rs!inte_Car_numero)) + ".txt lpt1"
                      Close #2
                      x = Shell(var_Archivo, vbHide)
                   End If
                    rs.Close
               Next var_k
            Else
               If var_empresa = "16" Then
''''' nota de credito para otras empresas
                  var_Archivo = App.Path & "\nota_credito" + Trim(Str(var_numero_nota_inicio)) + ".bat"
                  Open (App.Path & "\nota_credito" + Trim(Str(var_numero_nota_inicio)) + ".bat") For Output As #2
                  For var_k = var_numero_nota_inicio To Val(txt_a) + 1
                      rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DF' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_k), cnn, adOpenDynamic, adLockOptimistic
                      If Not rs.EOF Then
                         Open (App.Path & "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".txt") For Output As #1
                         Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                         Print #1, Chr(27) + Chr(64)
                         Print #1, Spc(92); Str(rs!inte_Car_numero)
                         Print #1, ""
                         Print #1, Spc(92); "       "; Format(rs!dtim_Car_fecha, "Short Date")
                         var_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                         For var_j = 1 + Len(Trim(var_cliente)) To 63
                             var_cliente = var_cliente + " "
                         Next var_j
                         var_cliente = var_cliente + " "
                         Print #1, ""
                         Print #1, Spc(12); var_cliente
                         var_domicilio = Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION))
                         var_j = 1 + Len(Trim(var_domicilio))
                         For var_j = var_j To 70
                             var_domicilio = var_domicilio + " "
                         Next var_j
                         var_domicilio = var_domicilio + " AGUASCALIENTES, AGS"
                         var_j = Len(var_domicilio)
                         var_agente = ""
                         var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                         For var_j = 1 + Len(Trim(var_agente)) To 8
                             var_agente = var_agente + " "
                         Next var_j
                         var_agente = var_agente
                         var_domicilio = var_domicilio
                         Print #1, Spc(12); var_domicilio
                         Print #1, Spc(12); IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
                         var_ciudad = ""
                         var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                         
                         For var_j = 1 + Len(Trim(var_ciudad)) To 14
                             var_ciudad = var_ciudad + " "
                         Next var_j
                              
                         var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                         var_ciudad = var_ciudad
                         
                         For var_j = 1 + Len(Trim(var_rfc)) To 79
                             var_rfc = var_rfc + " "
                         Next var_j
                         var_rfc = var_rfc + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                         For var_j = 1 + Len(Trim(var_rfc)) To 103
                             var_rfc = var_rfc + " "
                         Next var_j
                         var_rfc = var_rfc + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                         Print #1, Spc(12); var_ciudad
                         Print #1, Spc(12); var_rfc
                         Print #1, ""
                         Print #1, ""
                         var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                         
                         rsaux3.Open "select * from TB_DETALLE_DESCUENTOS_FINANCIEROS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DF' and vcha_ser_serie_id  = '" + var_serie + "' and vcha_car_clase_id = 'DF' and inte_car_numero =  " + Str(var_k), cnn, adOpenDynamic, adLockOptimistic
                         var_contador_lineas = 0
                         var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                         While Not rsaux3.EOF
                               'var_linea = "DF" + Str(rs!inte_car_numero) + " " + rs!vcha_Car_nombre + " Factura " + Str(rsaux3!inte_ddf_factura) + " " + CStr(rsaux3!floa_ddf_descuento_otorgado) + "%"
                               var_linea = "DF" + Str(rs!inte_Car_numero) + " " + rs!vcha_Car_nombre + " Factura " + Str(rsaux3!inte_ddf_factura)
                               'If Round(rsaux3!floa_ddf_descuento_otorgado, 2) <> Round(rsaux3!floa_ddf_descuento_aplicado, 2) Then
                               '   var_linea = var_linea + " (" + Format(rsaux3!floa_ddf_descuento_aplicado, "###,###,##0.0000") + "%)"
                               'End If
                               If Len(Trim(var_linea)) < 105 Then
                                  For var_j = 1 + Len(Trim(var_linea)) To 105
                                      var_linea = var_linea + " "
                                  Next var_j
                               End If
                               If Len(Trim(var_rfc)) = 0 Then
                                  var_importe_str = Format((IIf(IsNull(rsaux3!FLOA_DDF_IMPORTE), 0, rsaux3!FLOA_DDF_IMPORTE)), "###,###,##0.00")
                                  If Len(Trim(var_importe_str)) < 14 Then
                                     For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                         var_importe_str = " " + var_importe_str
                                     Next var_j
                                  End If
                               Else
                                  var_importe_str = Format((IIf(IsNull(rsaux3!FLOA_DDF_IMPORTE), 0, rsaux3!FLOA_DDF_IMPORTE) / (1 + (var_iva / 100))), "###,###,##0.00")
                                  If Len(Trim(var_importe_str)) < 14 Then
                                     For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                         var_importe_str = " " + var_importe_str
                                     Next var_j
                                  End If
                               End If
                               var_linea = var_linea + var_importe_str
                               Print #1, Spc(4); var_linea
                               rsaux3.MoveNext
                               var_contador_lineas = var_contador_lineas + 1
                          Wend
                          If var_contador_lineas < var_numero_renglones Then
                             For var_l = var_contador_lineas To var_numero_renglones - 1
                                 Print #1, Spc(4); ""
                             Next var_l
                          End If
                          rsaux3.Close
                          
                          var_cantidad_letra = rs!vcha_car_importe_letra
                          var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                          If Len(Trim(var_linea)) < 91 Then
                             For var_j = 1 + Len(Trim(var_linea)) To 91
                                 var_linea = var_linea + " "
                             Next var_j
                          End If
                          
                          Print #1, ""
                          
                          If Len(Trim(var_rfc)) = 0 Then
                             var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                             If Len(Trim(var_subimporte_str)) < 14 Then
                                For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                    var_subimporte_str = " " + var_subimporte_str
                                Next var_j
                             End If
                             var_iva_str = "-"
                             For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                 var_iva_str = " " + var_iva_str
                             Next var_j
                          Else
                            var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                             If Len(Trim(var_subimporte_str)) < 14 Then
                                For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                   var_subimporte_str = " " + var_subimporte_str
                               Next var_j
                             End If
                             var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                             If Len(Trim(var_iva_str)) < 14 Then
                                For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                    var_iva_str = " " + var_iva_str
                                Next var_j
                             End If
                          End If
                          var_linea = var_linea
                          Print #1, ""
                          Print #1, ""
                          Print #1, Spc(8); var_linea
                          Print #1, ""
                          Print #1, Spc(110); var_subimporte_str
                          Print #1, ""
                          Print #1, Spc(110); var_iva_str
                          Print #1, ""
                          var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                          If Len(Trim(var_importe_str)) < 14 Then
                             For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                 var_importe_str = " " + var_importe_str
                             Next var_j
                          End If
                          Print #1, Spc(110); var_importe_str
                          Print #1, ""
                          Print #1, ""
                          Print #1, ""
                          Print #1, ""
                          Close #1
                          Print #2, "copy " + App.Path + "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".txt lpt1"
                       End If
                       rs.Close
                  Next var_k
                  Close #2
                  x = Shell(var_Archivo, vbHide)
''''''''''  fin de la nota de credito para otras empresas
               Else
                   var_Archivo = App.Path & "\nota_credito" + Trim(Str(var_numero_nota_inicio)) + ".bat"
                   Open (App.Path & "\nota_credito" + Trim(Str(var_numero_nota_inicio)) + ".bat") For Output As #2
                   For var_k = var_numero_nota_inicio To Val(txt_a) + 1
                       rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DF' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_k), cnn, adOpenDynamic, adLockOptimistic
                      If Not rs.EOF Then
                         Open (App.Path & "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".txt") For Output As #1
                          'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                         Print #1, Chr(27) + Chr(64)
                         Print #1, Spc(92); Str(rs!inte_Car_numero)
                         Print #1, ""
                         'Print #1, Spc(92); "       "; Format(rs!DTIM_CAR_FECHA, "Short Date")
                         var_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                         For var_j = 1 + Len(Trim(var_cliente)) To 63
                             var_cliente = var_cliente + " "
                         Next var_j
                         var_cliente = var_cliente + " " + Format(rs!dtim_Car_fecha, "Short Date")
                         Print #1, ""
                         Print #1, Spc(12); var_cliente
                         var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                         For var_j = 1 + Len(Trim(var_domicilio)) To 83
                             var_domicilio = var_domicilio + " "
                         Next var_j
                         var_agente = ""
                         var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                         For var_j = 1 + Len(Trim(var_agente)) To 8
                             var_agente = var_agente + " "
                         Next var_j
                         var_agente = var_agente
                         var_domicilio = var_domicilio
                         Print #1, Spc(12); var_domicilio
                         var_ciudad = ""
                         var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                         var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                         If Len(Trim(var_estado)) > 0 Then
                            var_ciudad = var_ciudad + ", " + var_estado
                         End If
                         For var_j = 1 + Len(Trim(var_ciudad)) To 14
                             var_ciudad = var_ciudad + " "
                         Next var_j
                            
                         var_ciudad = var_ciudad
                           
                         var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                         var_ciudad = var_ciudad + " " + var_rfc
                          
                         For var_j = 1 + Len(Trim(var_ciudad)) To 79
                             var_ciudad = var_ciudad + " "
                         Next var_j
                         var_ciudad = var_ciudad + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                         For var_j = 1 + Len(Trim(var_ciudad)) To 103
                             var_ciudad = var_ciudad + " "
                         Next var_j
                         var_ciudad = var_ciudad + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                         
                         Print #1, Spc(12); var_ciudad
                         
                         var_rfc = "      " + var_rfc
                         For var_j = 1 + Len(Trim(var_rfc)) To 89
                             var_rfc = var_rfc + " "
                         Next var_j
                         var_rfc = var_rfc
                         'Print #1, Spc(6); var_rfc
                         Print #1, ""
                         Print #1, ""
                         var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                         
                         rsaux3.Open "select * from TB_DETALLE_DESCUENTOS_FINANCIEROS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DF' and vcha_ser_serie_id  = '" + var_serie + "' and vcha_car_clase_id = 'DF' and inte_car_numero =  " + Str(var_k), cnn, adOpenDynamic, adLockOptimistic
                         var_contador_lineas = 0
                         var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                         While Not rsaux3.EOF
                               'var_linea = "DF" + Str(rs!inte_car_numero) + " " + rs!vcha_Car_nombre + " Factura " + Str(rsaux3!inte_ddf_factura) + " " + CStr(rsaux3!floa_ddf_descuento_otorgado) + "%"
                               var_linea = "DF" + Str(rs!inte_Car_numero) + " " + rs!vcha_Car_nombre + " Factura " + Str(rsaux3!inte_ddf_factura)
                               'If Round(rsaux3!floa_ddf_descuento_otorgado, 2) <> Round(rsaux3!floa_ddf_descuento_aplicado, 2) Then
                               '   var_linea = var_linea + " (" + Format(rsaux3!floa_ddf_descuento_aplicado, "###,###,##0.0000") + "%)"
                               'End If
                               If Len(Trim(var_linea)) < 105 Then
                                  For var_j = 1 + Len(Trim(var_linea)) To 105
                                      var_linea = var_linea + " "
                                  Next var_j
                               End If
                               If Len(Trim(var_rfc)) = 0 Then
                                  var_importe_str = Format((IIf(IsNull(rsaux3!FLOA_DDF_IMPORTE), 0, rsaux3!FLOA_DDF_IMPORTE)), "###,###,##0.00")
                                  If Len(Trim(var_importe_str)) < 14 Then
                                     For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                         var_importe_str = " " + var_importe_str
                                     Next var_j
                                  End If
                               Else
                                  var_importe_str = Format((IIf(IsNull(rsaux3!FLOA_DDF_IMPORTE), 0, rsaux3!FLOA_DDF_IMPORTE) / (1 + (var_iva / 100))), "###,###,##0.00")
                                  If Len(Trim(var_importe_str)) < 14 Then
                                     For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                         var_importe_str = " " + var_importe_str
                                     Next var_j
                                  End If
                               End If
                               var_linea = var_linea + var_importe_str
                               Print #1, var_linea
                               rsaux3.MoveNext
                               var_contador_lineas = var_contador_lineas + 1
                          Wend
                          If var_contador_lineas < var_numero_renglones Then
                             For var_l = var_contador_lineas To var_numero_renglones - 1
                                 Print #1, ""
                             Next var_l
                          End If
                          rsaux3.Close
                           
                          var_cantidad_letra = rs!vcha_car_importe_letra
                          var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                          If Len(Trim(var_linea)) < 91 Then
                             For var_j = 1 + Len(Trim(var_linea)) To 91
                                 var_linea = var_linea + " "
                             Next var_j
                          End If
                          
                          Print #1, ""
                          If Len(Trim(var_rfc)) = 0 Then
                             var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                             If Len(Trim(var_subimporte_str)) < 14 Then
                                For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                    var_subimporte_str = " " + var_subimporte_str
                                Next var_j
                             End If
                             var_iva_str = "-"
                             For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                 var_iva_str = " " + var_iva_str
                              Next var_j
                           Else
                             var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                             If Len(Trim(var_subimporte_str)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                    var_subimporte_str = " " + var_subimporte_str
                                Next var_j
                              End If
                              var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                              If Len(Trim(var_iva_str)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                     var_iva_str = " " + var_iva_str
                                 Next var_j
                              End If
                           End If
                           var_linea = var_linea + "           " + var_subimporte_str
                           Print #1, Spc(4); var_linea
                           Print #1, Spc(106); var_iva_str
                           var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                           If Len(Trim(var_importe_str)) < 14 Then
                              For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                  var_importe_str = " " + var_importe_str
                              Next var_j
                           End If
                           Print #1, Spc(106); var_importe_str
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, Spc(45); "SISTEMAS"
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Close #1
                           
                           Print #2, "copy " + App.Path + "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".txt lpt1"
                        End If
                        rs.Close
                     Next var_k
                     Close #2
                     x = Shell(var_Archivo, vbHide)
                End If
            End If
   ''''''''''''''''
            MsgBox "Se a terminado de generar la nota de crédito", vbOKOnly, "ATENCION"
         End If
      End If
      cmd_imprimir.Enabled = False
   Else
      MsgBox "No se a seleccionado algun cliente", vbOKOnly, "ATENCION"
   End If
   End If
   Else
      MsgBox "No puede mezclar facturas con distintos tipos de IVA", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nota_credito_electronica_Click()
   Dim var_descuento_otorgado As Double
   Dim var_descuento_aplicado As Double
   Dim var_almacen As String
   Dim var_grupo_actual As String
   Dim var_grupo_real As String
   Dim var_cliente As String
   Dim var_titular As String
   Dim var_establecimiento As String
   Dim var_clave_moneda As String
   Dim var_agente As String
   Dim var_imprimir As Boolean
   Dim var_contador As Integer
   Dim var_contador_notas As Integer
   Dim var_contador_lineas As Integer
   Dim var_tipo_Cambio As Double
   Dim var_iva As Double
   Dim var_importe_total As Double
   Dim var_importe_iva As Double
   Dim var_subimporte As Double
   Dim var_importe As Double
   Dim si, i, n As Integer
   Dim var_saldo As Double
   Dim var_serie_cargo As String
   Dim var_numero_nota_inicio As Double
   Dim var_factura As Double
   Dim var_k As Double
   Dim var_descuentos As Double
   Dim var_desc_otorgado_str As String
   Dim var_desc_apilcado_str As String
   
   var_iva_pasado = 0
   var_posible_iva = 1
   For var_j = 1 To Me.lv_importes.ListItems.Count
       Me.lv_importes.ListItems.item(var_j).Selected = True
       If var_iva_pasado = 0 Then
          var_iva_pasado = CDbl(Me.lv_importes.selectedItem.SubItems(11))
       Else
          If var_iva_pasado <> CDbl(Me.lv_importes.selectedItem.SubItems(11)) Then
             var_posible_iva = 0
          End If
       End If
   Next var_j
   If var_posible_iva = 1 Then
   var_iva = var_iva_pasado
   
   
   Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
   Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
   var_descuentos = lv_importes.ListItems.Count
   If Not IsNumeric(Me.txt_importe) Then
      MsgBox "No es posible imprimir el importe de esta nota de crédito", vbOKOnly, "ATENCION"
   Else
   If var_descuentos > 0 Then
      si = MsgBox("¿Deseas Imprimir las Notas de Crédito", vbYesNo, "ATENCION")
      If si = 6 Then
         si = MsgBox("Confirmar la impresión de las Notas de Crédito", vbYesNo, "ATENCION")
         If si = 6 Then
            cnn.BeginTrans
            var_numero_renglones = 10000
            rs.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            var_tolerancia_saldo = rs!FLOA_PRI_TOLERANCIA_SALDOS
            rs.Close
            var_almacen = lv_importes.selectedItem.SubItems(6)
            var_grupo_actual = lv_importes.selectedItem.SubItems(7)
            var_grupo_real = lv_importes.selectedItem.SubItems(8)
            var_titular = lv_importes.selectedItem.SubItems(9)
            var_agente = txt_clave_agente
            var_cliente = txt_clave_cliente
            var_establecimiento = lv_importes.selectedItem.SubItems(10)
            'var_iva = (lv_importes.selectedItem.SubItems(11) * 1)
            'var_iva = 15
            var_clave_moneda = lv_importes.selectedItem.SubItems(12)
            var_tipo_Cambio = (lv_importes.selectedItem.SubItems(13) * 1)
            rs.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            var_serie = IIf(IsNull(rs!vcha_Ser_Serie_id), "", rs!vcha_Ser_Serie_id)
            If var_empresa = "02" Then
               If var_unidad_organizacional = "23" Then
                  var_serie = "NCEFT"
               Else
                  var_serie = "NCEMX"
               End If
            End If
            If var_empresa = "03" Then
               var_serie = "NCEVII"
            End If
            If var_empresa = "18" Then
               var_serie = "NCEVXX"
            End If
            
            rs.Close
            var_insertar = False
            n = lv_importes.ListItems.Count
            var_imprimir = False
            var_contador = 0
            var_contador_notas = 0
            var_tolerancia_saldo = var_tolerancia_saldo / var_tipo_Cambio
            var_numero_nota_inicio = var_numero_nota
            For i = 1 To n
               lv_importes.ListItems.item(i).Selected = True
               var_descuento_otorgado = lv_importes.selectedItem.SubItems(4) * 1
               var_descuento_aplicado = lv_importes.selectedItem.SubItems(16) * 1
               var_saldo = lv_importes.selectedItem.SubItems(15) * 1
               If var_saldo < var_tolerancia_saldo Then
                  'var_importe = var_importe + (lv_importes.selectedItem.SubItems(5) * 1) + var_saldo
                  var_importe = var_importe + (lv_importes.selectedItem.SubItems(5) * 1)
               Else
                  var_importe = var_importe + (lv_importes.selectedItem.SubItems(5) * 1)
               End If
               var_contador = var_contador + 1
               If (var_contador = var_numero_renglones) Or (i = n) Then
                  var_contador = 0
                  var_imprimir = True
               End If
               rs.Open "update tb_relacion_cobranza set INTE_RCO_NUMERO_DESCUENTO_FINANCIERO = " + Str(var_numero_nota) + ", DTIM_RCO_FECHA_DESCUENTO_FINANCIERO = getdate() where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + txt_relacion_cobranza + "' and inte_car_numero = " + lv_importes.selectedItem + " and vcha_rco_cheque = '" + lv_importes.selectedItem.SubItems(18) + "' and inte_rco_partida = " + lv_importes.selectedItem.SubItems(19), cnn, adOpenDynamic, adLockOptimistic
               If var_imprimir = True Then
                  var_subimporte = var_importe / (1 + (var_iva / 100))
                  var_importe_iva = var_importe - var_subimporte
                  var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "NC", "DF", "DF", var_numero_nota, "-", var_almacen, "", 0, Date, var_agente, var_grupo_actual, var_grupo_real, var_titular, var_cliente, var_establecimiento, 0, var_iva, 0, 0, 0, 0, 0, var_importe * var_tipo_Cambio, var_importe_iva * var_tipo_Cambio, 0, 0, 0, 0, 0, var_subimporte * var_tipo_Cambio, var_importe * var_tipo_Cambio, "", var_clave_usuario_global, "", Date, 0, Date, Date, var_clave_moneda, var_tipo_Cambio, var_serie, "")
                  rsaux3.Open "insert into tb_secuencia_notas_credito (vcha_emp_empresa_id, vcha_Ser_serie_id, inte_snc_numero_anterior, inte_snc_numero_actual) values ('" + var_empresa + "', '" + var_serie + "', " + CStr(var_numero_nota) + ", " + CStr(var_numero_nota) + ")", cnn, adOpenDynamic, adLockOptimistic
                  var_numero_nota = var_numero_nota + 1
                  var_contador_notas = var_contador_notas + 1
                  var_importe = 0
                  var_subimporte = 0
                  var_importe_iva = 0
               End If
               var_imprimir = False
            Next i
            var_numero_nota = var_numero_nota_anterior
            var_importe = 0
            For i = 1 To n
               lv_importes.ListItems.item(i).Selected = True
               var_saldo = lv_importes.selectedItem.SubItems(15) * 1
               var_serie_cargo = lv_importes.selectedItem.SubItems(14)
               If var_saldo < var_tolerancia_saldo Then
                  'var_importe = (lv_importes.selectedItem.SubItems(5) * 1) + var_saldo
                  var_importe = (lv_importes.selectedItem.SubItems(5) * 1)
               Else
                  var_importe = (lv_importes.selectedItem.SubItems(5) * 1)
               End If
               var_contador = var_contador + 1
               If (var_contador = var_numero_renglones) Or (i = n) Then
                  var_contador = 0
                  var_imprimir = True
               End If
               var_factura = lv_importes.selectedItem
               var_descuento_otorgado = lv_importes.selectedItem.SubItems(4) * 1
               var_descuento_aplicado = lv_importes.selectedItem.SubItems(16) * 1
               'var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, var_serie_cargo, "FA", lv_importes.selectedItem, var_serie, "DF", var_numero_nota, 0, (var_importe * var_tipo_Cambio))
               var_cadena = "insert into tb_estado_cuenta (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_ecu_movimiento_cargo, vcha_ecu_serie_cargo, inte_Ecu_numero_cargo, floa_ecu_importe_cargo, vcha_ecu_movimiento_abono, vcha_ecu_serie_abono, inte_ecu_numero_abono, floa_Ecu_importe_abono) "
               var_cadena = var_cadena + " values ('" + var_empresa + "','" + var_unidad_organizacional + "','FA','" + var_serie_cargo + "'," + Trim(CStr(CDbl(lv_importes.selectedItem))) + ",0,'DF','" + var_serie + "'," + CStr(var_numero_nota) + "," + CStr(Round((var_importe * var_tipo_Cambio), 2)) + ")"
               rsaux8.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               'rsaux9.Open "select * from tb_estado_cuenta where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ecu_movimiento_cargo = 'FA' and vcha_ecu_Serie_cargo = '" + var_serie_cargo + "' and inte_Ecu_numero_cargo = " + Trim(CStr(CDbl(lv_importes.selectedItem))) + " and vcha_ecu_movimiento_abono = 'DF' and vcha_ecu_serie_abono = '" + var_serie + "' and inte_Ecu_numero_abono = " + CStr(var_numero_nota), cnn, adOpenDynamic, adLockOptimistic
               'If rsaux9.EOF Then
               '   rsaux10.Open "select floa_sal_importe from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_Serie_id = '" + var_serie_cargo + "' and inte_Car_numero = " + Trim(CStr(CDbl(lv_importes.selectedItem))), cnn, adOpenDynamic, adLockOptimistic
               '   If Not rsaux10.EOF Then
               '      var_saldo_factura = rsaux10(0).Value
               '   End If
               '   rsaux10.Close
               '   var_diferencia_saldo_pago = var_saldo_factura - Round((var_importe * var_tipo_Cambio), 2)
               '   If var_diferencia_saldo_pago = -0.01 Then
               '      var_importe = var_importe - 0.01
               '   End If
               '   var_cadena = "insert into tb_estado_cuenta (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_ecu_movimiento_cargo, vcha_ecu_serie_cargo, inte_Ecu_numero_cargo, floa_ecu_importe_cargo, vcha_ecu_movimiento_abono, vcha_ecu_serie_abono, inte_ecu_numero_abono, floa_Ecu_importe_abono) "
               '   var_cadena = var_cadena + " values ('" + var_empresa + "','" + var_unidad_organizacional + "','FA','" + var_serie_cargo + "'," + Trim(CStr(CDbl(lv_importes.selectedItem))) + ",0,'DF','" + var_serie + "'," + CStr(var_numero_nota) + "," + CStr(Round((var_importe * var_tipo_Cambio), 2)) + ")"
               '   rsaux8.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               'End If
               'rsaux9.Close
               rsaux3.Open "Insert into TB_DETALLE_DESCUENTOS_FINANCIEROS (vcha_emp_empresa_id, vcha_car_documento, vcha_ser_serie_id, vcha_car_clase_id, inte_car_numero, vcha_ddf_concepto, floa_ddf_importe, inte_ddf_factura, floa_ddf_iva, floa_ddf_descuento_otorgado, floa_ddf_descuento_aplicado) values ('" + var_empresa + "', 'DF', '" + var_serie + "','DF'," + Str(var_numero_nota) + ",'', " + Str((var_importe * var_tipo_Cambio)) + ", " + Str(var_factura) + ", " + CStr(var_iva) + ", " + CStr(var_descuento_otorgado) + ", " + CStr(var_descuento_aplicado) + " )", cnn, adOpenDynamic, adLockOptimistic
               If var_imprimir = True Then
                  var_subimporte = var_importe / (1 + (var_iva / 100))
                  var_importe_iva = var_importe - var_subimporte
                  var_numero_nota = var_numero_nota + 1
                  var_importe = 0
                  var_importe_iva = 0
               End If
               var_imprimir = False
            Next i
            rs.Open "update tb_series set inte_ser_nota_credito = inte_ser_nota_credito + " + Str(var_contador_notas) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
''''''''''' se imprime la nota de credito
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DF' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(Me.txt_de), cnn, adOpenDynamic, adLockOptimistic
            Open (var_ruta_documentos_electronicos & "\" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".fi") For Output As #1
            
            
            var_rfc_cliente_1 = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
            var_rfc_cliente = ""
            If var_rfc_cliente_1 = "" Then
               var_rfc_cliente = "XAXX010101000"
            Else
               For var_j = 1 To Len(var_rfc_cliente_1)
                   If Mid(var_rfc_cliente_1, var_j, 1) <> "-" Then
                      If Mid(var_rfc_cliente_1, var_j, 1) <> "" Then
                         If Mid(var_rfc_cliente_1, var_j, 1) <> " " Then
                            var_rfc_cliente = var_rfc_cliente + Mid(var_rfc_cliente_1, var_j, 1)
                         End If
                      End If
                   End If
               Next var_j
            End If
            
            
            
            var_cadena = "Outputmode=" + Chr(13) + "<Factura>" + Chr(13) + "<Comprobante>" + Chr(13) + "Version=2.0" + Chr(13) + "Serie=" + rs!vcha_Ser_Serie_id + Chr(13) + "folio=" + CStr(rs!inte_Car_numero) + Chr(13)
            var_año = CStr(Year(rs!dtim_Car_fecha))
            var_mes = CStr(Month(rs!dtim_Car_fecha))
            var_dia = CStr(Day(rs!dtim_Car_fecha))
            var_hora = CStr(Hour(rs!dtim_Car_fecha))
            var_minuto = CStr(Minute(rs!dtim_Car_fecha))
            var_segundo = CStr(Second(rs!dtim_Car_fecha))
            If Len(var_año) = 2 Then
               var_año = "20" + var_año
            End If
            If Len(var_mes) = 1 Then
               var_mes = "0" + var_mes
            End If
            If Len(var_dia) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(var_hora) = 1 Then
               var_hora = "0" + var_hora
            End If
            If Len(var_minuto) = 1 Then
               var_minuto = "0" + var_minuto
            End If
            If Len(var_segundo) = 1 Then
               var_segundo = "0" + var_segundo
            End If
            var_rfc_cliente_1 = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
            var_rfc_cliente = ""
            If var_rfc_cliente_1 = "" Then
               var_rfc_cliente = "XAXX010101000"
            Else
               For var_j = 1 To Len(var_rfc_cliente_1)
                   If Mid(var_rfc_cliente_1, var_j, 1) <> "-" Then
                      If Mid(var_rfc_cliente_1, var_j, 1) <> "" Then
                         If Mid(var_rfc_cliente_1, var_j, 1) <> " " Then
                            var_rfc_cliente = var_rfc_cliente + Mid(var_rfc_cliente_1, var_j, 1)
                         End If
                      End If
                   End If
               Next var_j
            End If
            
            var_cadena_fecha = var_año + "-" + var_mes + "-" + var_dia + "T" + var_hora + ":" + var_minuto + ":" + var_segundo
            var_cadena = var_cadena + "fecha=" + var_cadena_fecha + Chr(13)
            var_cadena = var_cadena + "noAprobacion=" + Chr(13)
            var_cadena = var_cadena + "anoAprobacion=" + Chr(13)
            var_cadena = var_cadena + "tipoDeComprobante=NOTA DE CREDITO" + Chr(13)
            var_cadena = var_cadena + "formaDePago=PAGO HECHO EN UNA SOLA EXHIBICION" + Chr(13)
            var_cadena = var_cadena + "condicionesDePago=" + Chr(13)
            If var_rfc_cliente = "XAXX010101000" Then
               var_cadena = var_cadena + "subtotal=" + Format(CStr(rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
            Else
               var_cadena = var_cadena + "subtotal=" + Format(CStr(rs!floa_car_subimporte / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
            End If
            var_cadena = var_cadena + "descuento=" + Chr(13)
            var_cadena = var_cadena + "descuento1=" + Chr(13)
            var_cadena = var_cadena + "descuento2=" + Chr(13)
            var_cadena = var_cadena + "conceptodescuento1=" + Chr(13)
            var_cadena = var_cadena + "conceptodescuento2=" + Chr(13)
            var_cadena = var_cadena + "tasadescuento1=" + Chr(13)
            var_cadena = var_cadena + "tasadescuento2=" + Chr(13)
            If rsaux2.State = 1 Then
               rsaux2.Close
            End If
            rsaux2.Open "select * from tb_empresa_FACTURA_ELECTRONICA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            var_certificado = rsaux2!vcha_emp_certificado
            var_expedido = rsaux2!vcha_emp_expedido
            If var_rfc_cliente = "XAXX010101000" Then
               var_cadena = var_cadena + "iva=" + Format(0, "###,###,##0.000000") + Chr(13)
            Else
               var_cadena = var_cadena + "iva=" + Format(CStr(rs!floa_car_importe_iva / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
            End If
            var_cadena = var_cadena + "total=" + Format(CStr(rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
            var_cadena = var_cadena + "retencion=" + Chr(13)
            var_cadena = var_cadena + "factorretencioniva=" + Chr(13)
            var_cadena = var_cadena + "</Comprobante>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<Emisor>" + Chr(13)
            var_cadena = var_cadena + "erfc=" + rsaux2!VCHA_eMP_RFC + Chr(13)
            var_cadena = var_cadena + "enombre=" + rsaux2!VCHA_EMP_NOMBRE + Chr(13)
            var_cadena = var_cadena + "</Emisor>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<DomicilioFiscal>" + Chr(13)
            var_cadena = var_cadena + "ecalle=" + rsaux2!VCHA_eMP_CALLE + Chr(13)
            var_cadena = var_cadena + "enoExterior=" + rsaux2!VCHA_eMP_exterior + Chr(13)
            var_cadena = var_cadena + "enoInterior=" + Chr(13)
            var_cadena = var_cadena + "ecolonia=" + rsaux2!VCHA_eMP_COLONIA + Chr(13)
            var_cadena = var_cadena + "elocalidad=" + rsaux2!VCHA_EMP_LOCALIDAD + Chr(13)
            var_cadena = var_cadena + "ereferencia=" + Chr(13)
            var_cadena = var_cadena + "emunicipio=" + rsaux2!VCHA_EMP_MUNICIPIO + Chr(13)
            var_cadena = var_cadena + "eestado=" + rsaux2!VCHA_EMP_ESTADO + Chr(13)
            var_cadena = var_cadena + "epais=" + rsaux2!VCHA_eMP_PAIS + Chr(13)
            var_cadena = var_cadena + "ecodigoPostal=" + rsaux2!VCHA_EMP_CODIGO_POSTAL + Chr(13)
            var_cadena = var_cadena + "etel=" + IIf(IsNull(rsaux2!VCHA_EMP_TELEFONO), "", rsaux2!VCHA_EMP_TELEFONO) + Chr(13)
            var_cadena = var_cadena + "eemail=" + IIf(IsNull(rsaux2!VCHA_EMP_EMAIL), "", rsaux2!VCHA_EMP_EMAIL) + Chr(13)
            var_cadena = var_cadena + "</DomicilioFiscal>" + Chr(13) + Chr(13)
            
            var_cadena = var_cadena + "<ExpedidoEn>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "ex_calle=" + rsaux2!VCHA_eMP_CALLE + Chr(13)
            var_cadena = var_cadena + "ex_noExterior=" + rsaux2!VCHA_eMP_exterior + Chr(13)
            var_cadena = var_cadena + "ex_noInterior=" + Chr(13)
            var_cadena = var_cadena + "ex_colonia=" + rsaux2!VCHA_eMP_COLONIA + Chr(13)
            var_cadena = var_cadena + "ex_localidad=" + rsaux2!VCHA_EMP_LOCALIDAD + Chr(13)
            var_cadena = var_cadena + "ex_referencia=" + Chr(13)
            var_cadena = var_cadena + "ex_municipio=" + rsaux2!VCHA_EMP_MUNICIPIO + Chr(13)
            var_cadena = var_cadena + "ex_estado=" + rsaux2!VCHA_EMP_ESTADO + Chr(13)
            var_cadena = var_cadena + "ex_pais=" + rsaux2!VCHA_eMP_PAIS + Chr(13)
            var_cadena = var_cadena + "ex_codigoPostal=" + rsaux2!VCHA_EMP_CODIGO_POSTAL + Chr(13)
            var_cadena = var_cadena + "</ExpedidoEn>"
            
            
            
            var_cadena = var_cadena + "<Receptor>" + Chr(13)
            var_cadena = var_cadena + "noCliente=" + rs!vcha_cli_clave_id + Chr(13)
            rsaux2.Close
                                         
            var_rfc_cliente_1 = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
            var_rfc_cliente = ""
            If var_rfc_cliente_1 = "" Then
               var_rfc_cliente = "XAXX010101000"
            Else
               For var_j = 1 To Len(var_rfc_cliente_1)
                   If Mid(var_rfc_cliente_1, var_j, 1) <> "-" Then
                      If Mid(var_rfc_cliente_1, var_j, 1) <> "" Then
                         If Mid(var_rfc_cliente_1, var_j, 1) <> " " Then
                            var_rfc_cliente = var_rfc_cliente + Mid(var_rfc_cliente_1, var_j, 1)
                         End If
                      End If
                   End If
               Next var_j
            End If
            If var_empresa = "03" Or var_empresa = "28" Then
               var_rfc_cliente = "XEXX010101000"
            End If
            var_cadena = var_cadena + "rfc=" + var_rfc_cliente + Chr(13)
            var_cadena = var_cadena + "nombre=" + rs!VCHA_CLI_NOMBRE + Chr(13)
            var_cadena = var_cadena + "</Receptor>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<Cliente>" + Chr(13)
            var_cadena = var_cadena + "domicilio=" + IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + Chr(13)
            var_cadena = var_cadena + "calle=" + Chr(13)
            var_cadena = var_cadena + "noExterior=" + Chr(13)
            var_cadena = var_cadena + "noInterior=" + Chr(13)
            var_cadena = var_cadena + "colonia=" + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre) + Chr(13)
            var_cadena = var_cadena + "localidad=" + IIf(IsNull(rs!vcha_ciu_nombre), "", rs!VCHA_CLI_NOMBRE) + Chr(13)
            rsaux2.Open "select * from vw_clientes where vcha_Cli_clave_id = '" + rs!vcha_cli_clave_id + "'"
            var_cadena = var_cadena + "referencia=" + Chr(13)
            var_cadena = var_cadena + "municipio=" + IIf(IsNull(rsaux2!vcha_mun_nombre), "", rsaux2!vcha_mun_nombre) + Chr(13)
            var_cadena = var_cadena + "estado=" + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre) + Chr(13)
            VAR_NOMBRE_PAIS = IIf(IsNull(rs!vcha_pai_nombre), "MEXICO", rs!vcha_pai_nombre)
            If Trim(VAR_NOMBRE_PAIS) = "" Then
               VAR_NOMBRE_PAIS = "MEXICO"
            End If
            var_cadena = var_cadena + "pais=" + VAR_NOMBRE_PAIS + Chr(13)
            var_cadena = var_cadena + Chr(13)
            var_cadena = var_cadena + "codigoPostal=" + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP) + Chr(13)
            var_cadena = var_cadena + "tel=" + Chr(13)
            var_cadena = var_cadena + "email=" + IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email) + Chr(13)
            var_cadena = var_cadena + "</Cliente>" + Chr(13) + Chr(13)
                                         
            var_cadena = var_cadena + "<EntregarEn>" + Chr(13)
            var_cadena = var_cadena + "endomicilio=" + Chr(13)
            var_cadena = var_cadena + "encalle=" + Chr(13)
            var_cadena = var_cadena + "ennoExterior=" + Chr(13)
            var_cadena = var_cadena + "ennoInterior=" + Chr(13)
            var_cadena = var_cadena + "encolonia=" + Chr(13)
            var_cadena = var_cadena + "enlocalidad=" + Chr(13)
            var_cadena = var_cadena + "enreferencia=" + Chr(13)
            var_cadena = var_cadena + "enmunicipio=" + Chr(13)
            var_cadena = var_cadena + "enestado=" + Chr(13)
            var_cadena = var_cadena + "enpais=" + Chr(13)
            var_cadena = var_cadena + "encodigoPostal=" + Chr(13)
            var_cadena = var_cadena + "entel=" + Chr(13)
            var_cadena = var_cadena + "enemail=" + Chr(13)
            var_cadena = var_cadena + "</EntregarEn>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<Concepto>" + Chr(13)
            
            
            var_i = 1
            rsaux3.Open "select * from TB_DETALLE_DESCUENTOS_FINANCIEROS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DF' and vcha_ser_serie_id  = '" + var_serie + "' and vcha_car_clase_id = 'DF' and inte_car_numero =  " + Str(Me.txt_de), cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux3.EOF
                  pxx = CStr(var_i)
                  If Len(pxx) = 1 Then
                     pxx = "0" + pxx
                  End If
                  var_cadena = var_cadena + "p" + pxx + "_cantidad=1" + Chr(13)
                  var_cadena = var_cadena + "p" + pxx + "_unidad=DF" + Chr(13)
                  var_cadena = var_cadena + "p" + pxx + "_noIdentificacion=" + Chr(13)
                  var_linea = "DF" + Str(rs!inte_Car_numero) + " " + rs!vcha_Car_nombre + " Factura " + Str(rsaux3!inte_ddf_factura)
                  'If Round(rsaux3!floa_ddf_descuento_otorgado, 2) <> Round(rsaux3!floa_ddf_descuento_aplicado, 2) Then
                  '   var_linea = var_linea + " (" + Format(rsaux3!floa_ddf_descuento_aplicado, "###,###,##0.0000") + "%)"
                  'End If
                  var_cadena = var_cadena + "p" + pxx + "_descripcion=" + var_linea + Chr(13)
                  If var_rfc_cliente = "XAXX010101000" Then
                     var_importe_str = IIf(IsNull(rsaux3!FLOA_DDF_IMPORTE), 0, rsaux3!FLOA_DDF_IMPORTE)
                  Else
                     var_importe_str = IIf(IsNull(rsaux3!FLOA_DDF_IMPORTE), 0, rsaux3!FLOA_DDF_IMPORTE) / (1 + (rs!floa_car_porcentaje_iva / 100))
                  End If
                     
                  var_cadena = var_cadena + "p" + pxx + "_valorUnitario=" + Format(CStr(var_importe_str), "###,###,##0.000000") + Chr(13)
                  var_cadena = var_cadena + "p" + pxx + "_importe=" + Format(CStr(var_importe_str), "###,###,##0.000000") + Chr(13)
                  rsaux3.MoveNext
                  var_i = var_i + 1
            Wend
            rsaux3.Close
            
            
            
            
            
            var_cadena = var_cadena + "</Concepto>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<Otros>" + Chr(13)
            var_cadena = var_cadena + "certificado=" + IIf(IsNull(var_certificado), "", var_certificado) + Chr(13)
            rs.MoveFirst
            var_cadena = var_cadena + "cant_letra=" + rs!vcha_car_importe_letra + Chr(13)
            var_cadena = var_cadena + "factoriva=" + CStr(rs!floa_car_porcentaje_iva) + "%" + Chr(13)
            rsaux1.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id) + "'", cnn, adOpenDynamic, adLockOptimistic
            var_cadena = var_cadena + "moneda=" + IIf(IsNull(rsaux1!vcha_mon_nombre_plural), "", rsaux1!vcha_mon_nombre_plural) + Chr(13)
            rsaux1.Close
            var_cadena = var_cadena + "tipodeCambio=" + CStr(IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)) + Chr(13)
            var_cadena = var_cadena + "pedido=" + Chr(13)
            var_cadena = var_cadena + "Embarque=" + Chr(13)
            var_referencia_Bancaria = ""
            var_cadena = var_cadena + "referenciabancaria=" + Chr(13)
            var_cadena = var_cadena + "fechaPedido=" + Chr(13)
            var_cadena = var_cadena + "expedicion=" + Chr(13)
            var_cadena = var_cadena + "observaciones=" + Chr(13)
            var_cadena = var_cadena + "conceptoExtra1=" + Chr(13)
            var_cadena = var_cadena + "montoconceptoExtra1=" + Chr(13)
            var_cadena = var_cadena + "conceptoExtra2=" + Chr(13)
            var_cadena = var_cadena + "montoconceptoExtra2=" + Chr(13)
            var_cadena = var_cadena + "tipoimpresion=2" + Chr(13)
            
            rsaux11.Open "select * from vw_clientes where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux11.EOF Then
               var_cadena = var_cadena + "agente=" + rsaux11!VCHA_AGE_AGENTE_ID + " " + rsaux11!VCHA_AGE_NOMBRE + Chr(13)
            End If
            rsaux11.Close
            
            If var_empresa = "16" Or var_empresa = "02" Or var_empresa = "03" Or var_empresa = "18" Or var_empresa = "17" Or var_empresa = "06" Then
               var_cadena = var_cadena + "formato=MHNCVTH_V01.dat" + Chr(13)
            End If
            If var_empresa = "07" Then
               var_cadena = var_cadena + "formato=MHNCARE_V01.dat" + Chr(13)
            End If
            If var_empresa = "31" Then
               var_cadena = var_cadena + "formato=MHNCCAN_V01.dat" + Chr(13)
            End If
            If var_empresa = "42" Then
               var_cadena = var_cadena + "formato=MHNCCMA_V01.dat" + Chr(13)
            End If
            If var_empresa = "41" Then
               var_cadena = var_cadena + "formato=MHNCCOP_V01.dat" + Chr(13)
            End If
            If var_empresa = "15" Then
               var_cadena = var_cadena + "formato=MHNCERE_V01.dat" + Chr(13)
            End If
            If var_empresa = "33" Then
               var_cadena = var_cadena + "formato=MHNCMPU_V01.dat" + Chr(13)
            End If
            If var_empresa = "34" Then
               var_cadena = var_cadena + "formato=MHNCMYG_V01.dat" + Chr(13)
            End If
            If var_empresa = "160000" Then
               var_cadena = var_cadena + "formato=MHNCMYG_V01.dat" + Chr(13)
            End If
            If var_empresa = "36" Then
               var_cadena = var_cadena + "formato=MHNCSME_V01.dat" + Chr(13)
            End If
            If var_empresa = "30" Then
               var_cadena = var_cadena + "formato=MHNCTUR_V01.dat" + Chr(13)
            End If
            If var_empresa = "44" Then
               var_cadena = var_cadena + "formato=MHNCUTV_V01.dat" + Chr(13)
            End If
            If var_empresa = "38" Then
               var_cadena = var_cadena + "formato=MHNCVIA_V01.dat" + Chr(13)
            End If
            If var_empresa = "40" Then
               var_cadena = var_cadena + "formato=MHNCVIN_V01.dat" + Chr(13)
            End If
            If var_empresa = "43" Then
               var_cadena = var_cadena + "formato=MHNCVOP_V01.dat" + Chr(13)
            End If
            
            var_cadena = var_cadena + "</Otros>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<addenda>" + Chr(13)
            var_cadena = var_cadena + "</addenda>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "</Factura>"
            Print #1, var_cadena
            Close #1
            var_Archivo = App.Path & "\renombra" + var_serie + Trim(Str(rs!inte_Car_numero)) + ".bat"
            Open (App.Path & "\renombra" + var_serie + Trim(Str(rs!inte_Car_numero)) + ".bat") For Output As #2
            Print #2, "ren " + var_ruta_documentos_electronicos + "\" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".fi " + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".ff"
            Close #2
            
            x = Shell(var_Archivo, vbHide)
             
            rs.Close
            
            
            
''''''''''''''''
            MsgBox "Se a terminado de generar la nota de crédito", vbOKOnly, "ATENCION"
         End If
      End If
      Me.cmd_nota_credito_electronica.Enabled = False
   Else
      MsgBox "No se a seleccionado algun cliente", vbOKOnly, "ATENCION"
   End If
   End If
   Else
      MsgBox "No puede mezclar facturas con distintos tipos de IVA", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   txt_clave_cliente = ""
   lv_importes.ListItems.Clear
   txt_clave_cliente = ""
   txt_nombre_cliente = ""
   cmb_clientes = ""
   txt_clave_cliente.Enabled = False
   cmb_clientes.Enabled = True
   If var_empresa = "16" Or var_empresa = "15" Or var_empresa = "02" Or var_empresa = "03" Or var_empresa = "18" Then
      Me.cmd_nota_credito_electronica.Enabled = True
      Me.cmd_imprimir.Enabled = False
   Else
      cmd_imprimir.Enabled = True
      Me.cmd_nota_credito_electronica.Enabled = False
   End If
   'Me.cmb_clientes.SetFocus
   Me.txt_clave_cliente.Enabled = True
   Me.txt_clave_cliente.SetFocus
   Me.txt_importe = ""
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   frm_lista.Visible = False
   var_cadena_seguridad = ""
   txt_clave_cliente.Enabled = False
   rs.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'"
   var_numero_renglones = rs!INTE_PRI_RENGLONES_NOTA_CREDITO
   var_tolerancia_saldo = rs!FLOA_PRI_TOLERANCIA_SALDOS
   rs.Close
   If var_empresa = "02" Or var_empresa = "03" Then
      var_numero_renglones = 38
   Else
      If var_empresa = "16" Then
         var_numero_renglones = 6
      Else
         var_numero_renglones = 9
      End If
   End If
   cmd_nuevo.Enabled = True
   cmd_imprimir.Enabled = False
   If var_empresa = "15" Or var_empresa = "16" Or var_empresa = "02" Or var_empresa = "03" Or var_empresa = "18" Then
      Me.cmd_imprimir.Enabled = False
      Me.cmd_nota_credito_electronica.Enabled = True
   Else
      Me.cmd_imprimir.Enabled = True
      Me.cmd_nota_credito_electronica.Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_nota_credito_descuento_financiero)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_lista.ListItems.Count > 0 Then
         Me.txt_clave_cliente = Me.lv_lista.selectedItem
         Me.txt_nombre_cliente = Me.lv_lista.selectedItem.SubItems(1)
         Me.txt_clave_cliente.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
  Me.frm_lista.Visible = False
End Sub

Private Sub txt_clave_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_clave_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_cli_clave_id, vcha_cli_nombre from VW_NOTA_CREDITO_RELACION_COBRANZA where vcha_rco_folio = '" + txt_relacion_cobranza + "' and vcha_emp_empresa_id = '" + var_empresa + "'order by vcha_cli_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      Dim list_item As ListItem
      Dim var_contador_lista As Integer
      If Not rs.EOF Then
         While Not rs.EOF
            var_contador_lista = var_contador_lista + 1
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
            rs.MoveNext
         Wend
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
      rs.Close
   End If
End Sub

Private Sub txt_clave_cliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      lv_importes.SetFocus
   End If
End Sub

Private Sub txt_clave_cliente_LostFocus()
Dim list_item As ListItem
Dim var_importe As Double
Dim var_descuento As Double
Dim var_importe_total As Double
Dim var_total As Double
Dim var_contador As Double
Dim var_contador_notas As Double
Dim var_saldo_relacion As Double
If Trim(txt_clave_cliente) <> "" Then
      rs.Open "select distinct vcha_cli_nombre from VW_NOTA_CREDITO_RELACION_COBRANZA with (nolock) where vcha_rco_folio = '" + txt_relacion_cobranza + "' and vcha_cli_clave_ID = '" + txt_clave_cliente + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_clave_cliente.Enabled = False
         cmb_clientes.Enabled = False
         rsaux2.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'"
         var_tolerancia_saldo = rsaux2!FLOA_PRI_TOLERANCIA_SALDOS
         var_tolerancia_saldo_2 = rsaux2!FLOA_PRI_TOLERANCIA_SALDOS
         rsaux2.Close
         rsaux2.Open "select * from VW_NOTA_CREDITO_RELACION_COBRANZA where vcha_rco_folio = '" + txt_relacion_cobranza + "' and vcha_cli_clave_ID = '" + txt_clave_cliente + "' and vcha_emp_empresa_id = '" + var_empresa + "' order by inte_Car_numero", cnn, adOpenDynamic, adLockOptimistic
         var_serie = rsaux2!vcha_Ser_Serie_id
         var_numero_notas = rsaux2.RecordCount
         rsaux3.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
         var_serie = IIf(IsNull(rsaux3!vcha_Ser_Serie_id), "", rsaux3!vcha_Ser_Serie_id)
         If var_empresa = "02" Then
            If var_unidad_organizacional = "23" Then
               var_serie = "NCEFT"
            Else
               var_serie = "NCEMX"
            End If
         End If
         If var_empresa = "03" Then
            var_serie = "NCEVII"
         End If
         
         If var_empresa = "18" Then
            var_serie = "NCEVXX"
         End If
         
         
         rsaux3.Close
         rsaux3.Open "select inte_ser_nota_credito from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
         var_numero_nota = (IIf(IsNull(rsaux3!inte_ser_nota_credito), 1, rsaux3!inte_ser_nota_credito))
         var_numero_nota_anterior = (IIf(IsNull(rsaux3!inte_ser_nota_credito), 1, rsaux3!inte_ser_nota_credito))
         rsaux3.Close
         var_total = 0
         txt_de = var_numero_nota
         var_contador = 0
         var_contador_notas = 0
         var_factura_anterior = 0
         var_factura_actual = 0
         var_contador_lineas = 0
         While Not rsaux2.EOF
            Set list_item = lv_importes.ListItems.Add(, , rsaux2!inte_Car_numero)
            list_item.SubItems(1) = Format(IIf(IsNull(rsaux2!dtim_Car_fecha), "", rsaux2!dtim_Car_fecha), "Short Date")
            list_item.SubItems(2) = IIf(IsNull(rsaux2!vcha_mon_nombre_plural), "", rsaux2!vcha_mon_nombre_plural)
            var_importe = 0
            var_descuento = 0
            var_importe = (IIf(IsNull(rsaux2!floa_rco_importe), 0, rsaux2!floa_rco_importe)) * (1 + var_descuento / 100)
            var_descuento = IIf(IsNull(rsaux2!floa_Rco_descuento_aplicar), "", rsaux2!floa_Rco_descuento_aplicar)
            'MsgBox CStr(rsaux2!floa_rco_importe)
            list_item.SubItems(3) = Format((var_importe / ((100 - var_descuento) / 100)), "###,##0.00")
            list_item.SubItems(4) = Format(IIf(IsNull(rsaux2!floa_Rco_descuento_recomendado), 0, rsaux2!floa_Rco_descuento_recomendado), "###,##0.00")
            var_importe_total = Round(((var_importe / ((100 - var_descuento) / 100)) * (var_descuento / 100)), 2)
            list_item.SubItems(5) = Format(var_importe_total, "###,##0.00")
            list_item.SubItems(6) = IIf(IsNull(rsaux2!VCHA_ALM_ALMACEN_ID), "", rsaux2!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(7) = IIf(IsNull(rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID)
            list_item.SubItems(8) = IIf(IsNull(rsaux2!vcha_gre_grupo_real_id), "", rsaux2!vcha_gre_grupo_real_id)
            list_item.SubItems(9) = IIf(IsNull(rsaux2!vcha_tit_titular_id), "", rsaux2!vcha_tit_titular_id)
            list_item.SubItems(10) = IIf(IsNull(rsaux2!vcha_ESB_ESTABLECIMIENTO_id), "", rsaux2!vcha_ESB_ESTABLECIMIENTO_id)
            list_item.SubItems(11) = IIf(IsNull(rsaux2!floa_rco_iva), 0, rsaux2!floa_rco_iva)
            list_item.SubItems(12) = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
            list_item.SubItems(13) = IIf(IsNull(rsaux2!floa_rco_tipo_cambio), 1, rsaux2!floa_rco_tipo_cambio)
            list_item.SubItems(14) = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), "", rsaux2!vcha_Ser_Serie_id)
            rsaux3.Open "select floa_sal_importe from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_car_numero = " + Str(rsaux2!inte_Car_numero) + " and vcha_ser_serie_id = '" + rsaux2!vcha_Ser_Serie_id + "' and vcha_car_documento = '" + IIf(IsNull(rsaux2!vcha_car_documento), "", rsaux2!vcha_car_documento) + "'", cnn, adOpenDynamic, adLockOptimistic
            var_saldo_total_factura = rsaux3!FLOA_sAL_IMPORTE
            list_item.SubItems(22) = IIf(IsNull(rsaux3!FLOA_sAL_IMPORTE), 0, rsaux3!FLOA_sAL_IMPORTE)
            list_item.SubItems(15) = Format(IIf(IsNull(rsaux3!FLOA_sAL_IMPORTE), 0, rsaux3!FLOA_sAL_IMPORTE) - var_importe_total, "###,##0.00")
            list_item.SubItems(16) = Format(IIf(IsNull(rsaux2!floa_Rco_descuento_aplicar), 0, rsaux2!floa_Rco_descuento_aplicar), "###,##0.00")
            list_item.SubItems(18) = IIf(IsNull(rsaux2!VCHA_rCO_CHEQUE), "", rsaux2!VCHA_rCO_CHEQUE)
            list_item.SubItems(19) = IIf(IsNull(rsaux2!inte_rco_partida), 0, rsaux2!inte_rco_partida)
            rsaux3.Close
            var_tipo_Cambio = (lv_importes.selectedItem.SubItems(13) * 1)
            var_saldo = list_item.SubItems(15) * 1
            var_tolerancia_saldo = var_tolerancia_saldo / var_tipo_Cambio
            var_importe = 0
            'aqui deberia ir la tolerancia
            'fin de la tolerancia, se movio cuando el contador de facturas sea 1 para que no a todas les de la tolerancia
            
            var_contador = var_contador + 1
            If var_contador > var_numero_renglones Then
               var_contador_notas = var_contador_notas + 1
               var_contador = 0
            End If
            var_factura_actual = rsaux2!inte_Car_numero
            rsaux2.MoveNext
            If var_factura_actual = var_factura_anterior Then
               var_contador_lineas = var_contador_lineas + 1
               var_importe = var_importe_total * 1
               var_importe_total = var_importe
            
               list_item.SubItems(5) = Format(var_importe, "###,##0.00")
               var_total = var_total + var_importe_total
               var_total_factura = var_total_factura + var_importe_total
            Else
               var_contador_lineas = 1
               If var_saldo < var_tolerancia_saldo Then
                  var_importe = var_importe_total + var_saldo
               Else
                  var_importe = var_importe_total * 1
               End If
               var_importe_total = var_importe
              
               list_item.SubItems(5) = Format(var_importe, "###,##0.00")
               var_total = var_total + var_importe_total
               var_total_factura = var_importe_total
            End If
            If var_total_factura > var_saldo_total_factura Then
               If var_total_factura - var_saldo_total_factura < 1 Then
                  var_importe = var_importe - (var_total_factura - var_saldo_total_factura)
                  list_item.SubItems(5) = Format(var_importe, "###,##0.00")
               End If
            End If
            var_factura_anterior = var_factura_actual
            list_item.SubItems(21) = var_contador_lineas
         Wend
         
         txt_importe = Format(var_total, "###,##0.00")
         var_tolerancia_saldo = var_tolerancia_saldo_2
         
         'For var_j = 1 To Me.lv_importes.ListItems.Count
         '    Me.lv_importes.ListItems.Item(var_j).Selected = True
         '    If CDbl(Me.lv_importes.selectedItem.SubItems(21)) = 1 Then
         '       rsaux.Open "SELECT SUM(FLOA_RCO_IMPORTE / ((100 - FLOA_RCO_DESCUENTO_APLICAR) / 100))- SUM(FLOA_RCO_IMPORTE) AS importe_saldo_Relacion, INTE_CAR_NUMERO, VCHA_SER_SERIE_ID From dbo.VW_NOTA_CREDITO_RELACION_COBRANZA WHERE  vcha_rco_folio = '" + txt_relacion_cobranza + "' and vcha_cli_clave_ID = '" + txt_clave_cliente + "' and vcha_emp_empresa_id = '" + var_empresa + "' and (INTE_CAR_NUMERO = " + Me.lv_importes.selectedItem + ") group by INTE_CAR_NUMERO, VCHA_SER_SERIE_ID", cnn, adOpenDynamic, adLockOptimistic
         '       If Not rsaux.EOF Then
         '          rsaux3.Open "select * from tb_Saldos where vcha_Car_documento = 'FA' and vcha_cli_clave_id ='" + Me.txt_clave_cliente + "' and inte_Car_numero = " + CStr(rsaux!INTE_CAR_NUMERO) + " and vcha_ser_Serie_id = '" + rsaux!VCHA_SER_SERIE_ID + "'", cnn, adOpenDynamic, adLockOptimistic
         '          If Not rsaux3.EOF Then
         '             var_saldo_relacion = Round(rsaux!importe_saldo_relacion, 2)
         '             var_saldo = rsaux3!FLOA_sAL_IMPORTE - Round(rsaux!importe_saldo_relacion, 2)
         '             var_tipo_Cambio = (lv_importes.selectedItem.SubItems(13) * 1)
         '             'var_saldo = list_item.SubItems(15) * 1
         '             var_tolerancia_saldo = var_tolerancia_saldo / var_tipo_Cambio
         '             var_importe = 0
         '             If var_saldo > 0 Then
         '                If var_saldo < var_tolerancia_saldo Then
         '                   Me.lv_importes.selectedItem.SubItems(5) = Format(CDbl(Me.lv_importes.selectedItem.SubItems(5)) + var_saldo, "###,##0.00")
         '                   Me.txt_importe = Format(CDbl(Me.txt_importe) + var_saldo, "###,##0.00")
         '                End If
         '             End If
         '          End If
         '          rsaux3.Close
         '       End If
         '       rsaux.Close
         '    End If
         'Next var_j
         
         
         txt_a = txt_de + var_contador_notas
         rsaux2.Close
      Else
         MsgBox "El cliente no tiene descuentos por aplicar o no se encuentra en la relación de cobranza", vbOKOnly, "ATENCION"
         txt_clave_cliente = ""
         cmb_clientes = ""
      End If
      rs.Close
   If var_empresa = "16" Or var_empresa = "15" Or var_empresa = "02" Or var_empresa = "03" Or var_empresa = "18" Then
      If CDbl(Me.txt_importe) = 0 Then
         Me.cmd_nota_credito_electronica.Enabled = False
      Else
         Me.cmd_nota_credito_electronica.Enabled = True
      End If
   Else
      If CDbl(Me.txt_importe) = 0 Then
         Me.cmd_imprimir = False
      Else
         Me.cmd_imprimir.Enabled = True
      End If
   End If
End If
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_cli_clave_id, vcha_cli_nombre from VW_NOTA_CREDITO_RELACION_COBRANZA where vcha_rco_folio = '" + txt_relacion_cobranza + "' and vcha_emp_empresa_id = '" + var_empresa + "'order by vcha_cli_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      Dim list_item As ListItem
      Dim var_contador_lista As Integer
      If Not rs.EOF Then
         While Not rs.EOF
            var_contador_lista = var_contador_lista + 1
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
            rs.MoveNext
         Wend
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
      rs.Close
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_relacion_cobranza_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub
