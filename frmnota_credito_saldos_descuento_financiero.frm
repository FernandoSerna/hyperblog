VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmnota_credito_saldos_descuento_financiero 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nota de Crédito por Descuento Financiero"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmnota_credito_saldos_descuento_financiero.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos Generales "
      Height          =   1035
      Left            =   105
      TabIndex        =   10
      Top             =   435
      Width           =   8205
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   3435
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   585
         Width           =   4590
      End
      Begin VB.TextBox txt_clave_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   255
         Width           =   1320
      End
      Begin VB.TextBox txt_clave_cliente 
         Height          =   315
         Left            =   2100
         TabIndex        =   12
         Top             =   585
         Width           =   1320
      End
      Begin VB.TextBox txt_nombre_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3435
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   255
         Width           =   4590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   390
         TabIndex        =   15
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   375
         TabIndex        =   14
         Top             =   645
         Width           =   525
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Importes de descuentos no aplicados "
      Height          =   4785
      Left            =   120
      TabIndex        =   6
      Top             =   2310
      Width           =   8205
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6195
         TabIndex        =   7
         Top             =   4335
         Width           =   1845
      End
      Begin MSComctlLib.ListView lv_importes 
         Height          =   4095
         Left            =   75
         TabIndex        =   8
         Top             =   225
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
         NumItems        =   16
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
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   5205
         TabIndex        =   9
         Top             =   4425
         Width           =   570
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7950
      Picture         =   "frmnota_credito_saldos_descuento_financiero.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Notas de Crédito "
      Height          =   630
      Left            =   105
      TabIndex        =   0
      Top             =   1560
      Width           =   8205
      Begin VB.TextBox txt_de 
         Height          =   315
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1335
      End
      Begin VB.TextBox txt_a 
         Height          =   315
         Left            =   4350
         TabIndex        =   1
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   2040
         TabIndex        =   4
         Top             =   285
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   195
         Left            =   4155
         TabIndex        =   3
         Top             =   285
         Width           =   150
      End
   End
   Begin VB.Frame Frame3 
      Height          =   90
      Left            =   15
      TabIndex        =   17
      Top             =   300
      Width           =   8400
   End
End
Attribute VB_Name = "frmnota_credito_saldos_descuento_financiero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_numero_notas As Integer
Dim var_numero_renglones As Integer
Dim var_numero_nota As Double
Dim var_numero_nota_anterior As Integer
Dim var_serie As String
Dim var_tolerancia_saldo As Double

Private Sub cmb_clientes_LostFocus()
   If KeyAscii = 13 Then
      Dim list_item As ListItem
      Dim var_importe As Double
      Dim var_descuento As Double
      Dim var_importe_total As Double
      Dim var_total As Double
      Dim var_contador As Double
      Dim var_contador_notas As Double
      Dim var_importe_descuento As Double
      rs.Open "select distinct vcha_cli_nombre from VW_NOTA_CREDITO_SALDO_DESCUENTO_FINANCIERO where vcha_cli_clave_ID = '" + txt_clave_cliente + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_clave_cliente.Enabled = False
         cmb_clientes.Enabled = False
         rsaux2.Open "select distinct vcha_cli_nombre from VW_NOTA_CREDITO_SALDO_DESCUENTO_FINANCIERO where vcha_cli_clave_ID = '" + txt_clave_cliente + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
         var_numero_notas = rsaux2.RecordCount
         rsaux3.Open "select inte_pri_nota_credito from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
         var_numero_nota = (IIf(IsNull(rsaux3!inte_pri_nota_credito), 0, rsaux3!inte_pri_nota_credito) + 1)
         var_numero_nota_anterior = (IIf(IsNull(rsaux3!inte_pri_nota_credito), 0, rsaux3!inte_pri_nota_credito) + 1)
         rsaux3.Close
         var_total = 0
         txt_de = var_numero_nota
         var_contador = 0
         var_contador_notas = 0
         While Not rsaux2.EOF
            Set list_item = lv_importes.ListItems.Add(, , rsaux2!INTE_CAR_NUMERO)
            list_item.SubItems(1) = Format(IIf(IsNull(rsaux2!dtim_Car_fecha), "", rsaux2!dtim_Car_fecha), "Short Date")
            list_item.SubItems(2) = IIf(IsNull(rsaux2!vcha_mon_nombre_plural), "", rsaux2!vcha_mon_nombre_plural)
            var_importe = 0
            var_descuento = 0
            var_descuento = IIf(IsNull(rsaux2!floa_sap_descuento_1), 0, rsaux2!floa_Rco_descuento_aplicar)
            var_importe_descuento = IIf(IsNull(rsaux2!importe_saldo), 0, rsaux2!importe_saldo)
            var_importe = ((var_importe_descuento * 100) / var_descuento)
            list_item.SubItems(3) = Format(var_importe, "###,##0.00")
            list_item.SubItems(4) = Format(IIf(IsNull(rsaux2!floa_sap_descuento_1), 0, rsaux2!floa_Rco_descuento_aplicar), "###,##0.00")
            var_importe_total = var_importe_descuento
            list_item.SubItems(5) = Format(var_importe_total, "###,##0.00")
            list_item.SubItems(6) = IIf(IsNull(rsaux2!VCHA_ALM_ALMACEN_ID), "", rsaux2!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(7) = IIf(IsNull(rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID)
            list_item.SubItems(8) = IIf(IsNull(rsaux2!vcha_gre_grupo_real_id), "", rsaux2!vcha_gre_grupo_real_id)
            list_item.SubItems(9) = IIf(IsNull(rsaux2!vcha_tit_titular_id), "", rsaux2!vcha_tit_titular_id)
            list_item.SubItems(10) = IIf(IsNull(rsaux2!vcha_ESB_ESTABLECIMIENTO_id), "", rsaux2!vcha_ESB_ESTABLECIMIENTO_id)
            list_item.SubItems(11) = IIf(IsNull(rsaux2!floa_rco_iva), 0, rsaux2!floa_rco_iva)
            list_item.SubItems(12) = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
            list_item.SubItems(13) = IIf(IsNull(rsaux2!floa_rco_tipo_cambio), 1, rsaux2!floa_rco_tipo_cambio)
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

Private Sub cmd_nuevo_Click()

End Sub

Private Sub cmd_imprimir_Click()
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
   Dim var_tipo_Cambio As Double
   Dim var_iva As Double
   Dim var_importe_total As Double
   Dim var_importe_iva As Double
   Dim var_subimporte As Double
   Dim var_importe As Double
   Dim var_serie_cargo As String
   Dim si, i, n As Integer
   Dim var_saldo As Double
   Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
   Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
   If lv_importes.ListItems.Count > 0 Then
      si = MsgBox("¿Deseas Imprimir las Notas de Crédito", vbYesNo, "ATENCION")
      If si = 6 Then
         si = MsgBox("Confirmar la impresión de las Notas de Crédito", vbYesNo, "ATENCION")
         If si = 6 Then
            cnn.BeginTrans
            var_almacen = lv_importes.selectedItem.SubItems(6)
            var_grupo_actual = lv_importes.selectedItem.SubItems(7)
            var_grupo_real = lv_importes.selectedItem.SubItems(8)
            var_titular = lv_importes.selectedItem.SubItems(9)
            var_agente = txt_clave_agente
            var_cliente = txt_clave_cliente
            var_establecimiento = lv_importes.selectedItem.SubItems(10)
            var_iva = (lv_importes.selectedItem.SubItems(11) * 1)
            var_clave_moneda = lv_importes.selectedItem.SubItems(12)
            var_tipo_Cambio = (lv_importes.selectedItem.SubItems(13) * 1)
            var_insertar = False
            n = lv_importes.ListItems.Count
            var_imprimir = False
            var_contador = 0
            var_contador_notas = 0
            For i = 1 To n
               lv_importes.ListItems.Item(i).Selected = True
               var_saldo = lv_importes.selectedItem.SubItems(15) * 1
               If var_saldo < var_tolerancia_saldo Then
                  var_importe = var_importe + (lv_importes.selectedItem.SubItems(5) * 1) + var_saldo
               Else
                  var_importe = var_importe + (lv_importes.selectedItem.SubItems(5) * 1)
               End If
               var_contador = var_contador + 1
               If (var_contador = var_numero_renglones) Or (i = n) Then
                  var_contador = 0
                  var_imprimir = True
               End If
               If rs.State = 1 Then
                  rs.Close
               End If
               rs.Open "update tb_relacion_cobranza set inte_rco_numero_descuento_financiero = " + Str(var_numero_nota) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + txt_relacion_cobranza + "' and inte_car_numero = " + lv_importes.selectedItem, cnn, adOpenDynamic, adLockOptimistic
               If var_imprimir = True Then
                  var_subimporte = var_importe / (1 + (var_iva / 100))
                  var_importe_iva = var_importe - var_subimporte
                  var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "NC", "DF", "DF", var_numero_nota, "-", var_almacen, "", 0, Date, var_agente, var_grupo_actual, var_grupo_real, var_titular, var_cliente, var_establecimiento, 0, var_iva, 0, 0, 0, 0, 0, var_importe * var_tipo_Cambio, var_importe_iva * var_tipo_Cambio, 0, 0, 0, 0, 0, var_subimporte * var_tipo_Cambio, var_importe * var_tipo_Cambio, "", var_clave_usuario_global, "", Date, 0, Date, Date, var_clave_moneda, var_tipo_Cambio, var_serie, "")
                  rsaux3.Open "insert into tb_secuencia_notas_credito (vcha_emp_empresa_id, vcha_Ser_serie_id, inte_snc_numero_anterior, inte_snc_numero_actual) values ('" + var_empresa + "', '" + var_serie + "', " + CStr(var_numero_nota) + ", " + CStr(var_numero_nota) + ")", cnn, adOpenDynamic, adLockOptimistic
               End If
               var_imprimir = False
            Next i
            var_numero_nota = var_numero_nota_anterior
            For i = 1 To n
               lv_importes.ListItems.Item(i).Selected = True
               var_serie_cargo = lv_importes.selectedItem.SubItems(14)
               var_saldo = lv_importes.selectedItem.SubItems(15) * 1
               If var_saldo < var_tolerancia_saldo Then
                  var_importe = var_importe + (lv_importes.selectedItem.SubItems(5) * 1) + var_saldo
               Else
                  var_importe = var_importe + (lv_importes.selectedItem.SubItems(5) * 1)
               End If
               var_contador = var_contador + 1
               If (var_contador = var_numero_renglones) Or (i = n) Then
                  var_contador = 0
                  var_imprimir = True
               End If
               var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, var_serie_cargo, "FA", lv_importes.selectedItem, var_serie, "DF", var_numero_nota, 0, ((lv_importes.selectedItem.SubItems(5) * 1) * var_tipo_Cambio))
               
               rsaux3.Open "Insert into TB_DETALLE_DESCUENTOS_FINANCIEROS (vcha_emp_empresa_id, vcha_car_documento, vcha_ser_serie_id, vcha_car_clase_id, inte_car_numero, vcha_ddf_concepto, floa_ddf_importe, inte_ddf_factura, floa_ddf_iva, floa_ddf_descuento_otorgado, floa_ddf_descuento_aplicado) values ('" + var_empresa + "', 'DF', '" + var_serie + "','DF'," + Str(var_numero_nota) + ",'', " + Str(((lv_importes.selectedItem.SubItems(5) * 1) * var_tipo_Cambio)) + ", " + lv_importes.selectedItem + ", " + CStr(var_iva) + ", " + Me.lv_importes.selectedItem.SubItems(4) + ", " + Me.lv_importes.selectedItem.SubItems(4) + " )", cnn, adOpenDynamic, adLockOptimistic
               If var_imprimir = True Then
                  var_subimporte = var_importe / (1 + (var_iva / 100))
                  var_importe_iva = var_importe - var_subimporte
                  var_numero_nota = var_numero_nota
               End If
               var_imprimir = False
            Next i
            rs.Open "update tb_series set inte_ser_nota_credito = inte_ser_nota_credito + 1 where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            rsaux2.Open "update tb_saldos_aplicar set char_sap_estatus = '*' where vcha_emp_empresa_id = '" + var_empresa + "'  and vcha_cli_clave_id = '" + txt_clave_cliente + "' and char_sap_estatus <> '*'", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
'''''''''  '
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DF' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_numero_nota), cnn, adOpenDynamic, adLockOptimistic
            Me.lv_importes.ListItems(1).Selected = True
            If Not rs.EOF Then
'''''   '''''''''  IMPRESION DE LA NOTA DE CARGO
               Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt") For Output As #1
               'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
               Print #1, Chr(27) + Chr(64)
               Print #1, Spc(92); Str(rs!INTE_CAR_NUMERO)
               Print #1, ""
               Print #1, Spc(92); "       "; Format(rs!dtim_Car_fecha, "Short Date")
               Print #1, ""
               var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
               For var_j = 1 + Len(Trim(var_cliente)) To 83
                   var_cliente = var_cliente + " "
               Next var_j
               var_cliente = var_cliente + "AGUASCALIENTES, AGS."
               Print #1, Spc(12); var_cliente
               var_domicilio = Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION)) + " COL.: " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre) + "  C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
               For var_j = 1 + Len(Trim(var_domicilio)) To 83
                   var_domicilio = var_domicilio + " "
               Next var_j
               var_agente = ""
               var_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
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
               var_linea = "DF" + Str(rs!INTE_CAR_NUMERO) + " " + rs!vcha_Car_nombre + " " + lv_importes.selectedItem + " " + Me.lv_importes.selectedItem.SubItems(4)
               If Len(Trim(var_linea)) < 120 Then
                  For var_j = 1 + Len(Trim(var_linea)) To 120
                      var_linea = var_linea + " "
                  Next var_j
               End If
               var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
               If Len(Trim(var_importe_str)) < 14 Then
                  For var_j = 1 + Len(Trim(var_importe_str)) To 14
                      var_importe_str = " " + var_importe_str
                  Next var_j
               End If
               var_linea = var_linea + var_importe_str
               Print #1, var_linea
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               var_cantidad_letra = rs!vcha_car_importe_letra
                      
               var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
               If Len(Trim(var_linea)) < 105 Then
                  For var_j = 1 + Len(Trim(var_linea)) To 105
                      var_linea = var_linea + " "
                  Next var_j
               End If
                      
               var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                       
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
                      
              Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat") For Output As #2
              var_Archivo = App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat"
              Print #2, "copy " + App.Path + "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt lpt1"
              Close #2
              x = Shell(var_Archivo, vbHide)
''''''''''''''
           End If
           rs.Close
''''''''''
           MsgBox "Se a terminado de generar la nota de crédito", vbOKOnly, "ATENCION"
        End If
     End If
  Else
     MsgBox "El cliente no cuenta con ningún importe", vbOKOnly, "ATENCION"
  End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
Dim list_item As ListItem
Dim var_importe As Double
Dim var_descuento As Double
Dim var_importe_total As Double
Dim var_total As Double
Dim var_contador As Double
Dim var_contador_notas As Double
Dim var_importe_descuento As Double
Dim var_tipo_Cambio As Double
   var_cadena_seguridad = ""
   Top = 0
   Left = 1500
   txt_clave_cliente = frmasigna_pagos_no_aplicados.txt_clave_cliente
   txt_nombre_cliente = frmasigna_pagos_no_aplicados.txt_nombre_cliente
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open "select vcha_age_agente_id, vcha_age_nombre from vw_clientes_1 where vcha_cli_clave_id = '" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
   txt_clave_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
   txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
   rs.Close
   txt_clave_cliente.Enabled = False
   txt_nombre_cliente.Enabled = False
   rs.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'"
   var_numero_renglones = rs!INTE_PRI_RENGLONES_NOTA_CREDITO
   var_tolerancia_saldo = rs!FLOA_PRI_TOLERANCIA_SALDOS
   rs.Close
   rs.Open "select distinct vcha_cli_nombre from VW_NOTA_CREDITO_SALDO_DESCUENTO_FINANCIERO where vcha_cli_clave_ID = '" + txt_clave_cliente + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   
   If Not rs.EOF Then
      txt_clave_cliente.Enabled = False
      txt_nombre_cliente.Enabled = False
      rsaux2.Open "select * from VW_NOTA_CREDITO_SALDO_DESCUENTO_FINANCIERO where vcha_cli_clave_ID = '" + txt_clave_cliente + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      var_serie = rsaux2!VCHA_SER_SERIE_ID
      var_numero_notas = rsaux2.RecordCount
      rsaux3.Open "select inte_ser_nota_credito from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
      var_numero_nota = IIf(IsNull(rsaux3!inte_ser_nota_credito), 1, rsaux3!inte_ser_nota_credito)
      var_numero_nota_anterior = (IIf(IsNull(rsaux3!inte_ser_nota_credito), 1, rsaux3!inte_ser_nota_credito))
      rsaux3.Close
      var_total = 0
      txt_de = var_numero_nota
      var_contador = 0
      var_contador_notas = 0
      While Not rsaux2.EOF
         Set list_item = lv_importes.ListItems.Add(, , rsaux2!INTE_CAR_NUMERO)
         list_item.SubItems(1) = Format(IIf(IsNull(rsaux2!dtim_Car_fecha), "", rsaux2!dtim_Car_fecha), "Short Date")
         list_item.SubItems(2) = IIf(IsNull(rsaux2!vcha_mon_nombre_plural), "", rsaux2!vcha_mon_nombre_plural)
         var_importe = 0
         var_descuento = 0
         var_tipo_Cambio = IIf(IsNull(rsaux2!floa_sap_tipo_cambio), 1, rsaux2!floa_sap_tipo_cambio)
         var_descuento = IIf(IsNull(rsaux2!floa_sap_descuento_1), 0, rsaux2!floa_sap_descuento_1)
         var_importe_descuento = (IIf(IsNull(rsaux2!floa_sap_importe), 0, rsaux2!floa_sap_importe) / var_tipo_Cambio)
         var_importe = ((var_importe_descuento * 100) / var_descuento)
         list_item.SubItems(3) = Format(var_importe, "###,##0.00")
         list_item.SubItems(4) = Format(IIf(IsNull(rsaux2!floa_sap_descuento_1), 0, rsaux2!floa_sap_descuento_1), "###,##0.00")
         var_importe_total = var_importe_descuento
         list_item.SubItems(5) = Format(var_importe_total, "###,##0.00")
         list_item.SubItems(6) = IIf(IsNull(rsaux2!VCHA_ALM_ALMACEN_ID), "", rsaux2!VCHA_ALM_ALMACEN_ID)
         list_item.SubItems(7) = IIf(IsNull(rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID)
         list_item.SubItems(8) = IIf(IsNull(rsaux2!vcha_gre_grupo_real_id), "", rsaux2!vcha_gre_grupo_real_id)
         list_item.SubItems(9) = IIf(IsNull(rsaux2!vcha_tit_titular_id), "", rsaux2!vcha_tit_titular_id)
         list_item.SubItems(10) = IIf(IsNull(rsaux2!vcha_ESB_ESTABLECIMIENTO_id), "", rsaux2!vcha_ESB_ESTABLECIMIENTO_id)
         list_item.SubItems(11) = IIf(IsNull(rsaux2!floa_car_porcentaje_iva), 0, rsaux2!floa_car_porcentaje_iva)
         list_item.SubItems(12) = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
         list_item.SubItems(13) = IIf(IsNull(rsaux2!floa_sap_tipo_cambio), 1, rsaux2!floa_sap_tipo_cambio)
         list_item.SubItems(14) = IIf(IsNull(rsaux2!VCHA_SER_SERIE_ID), 1, rsaux2!VCHA_SER_SERIE_ID)
         rsaux3.Open "select floa_sal_importe from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_car_numero = " + Str(rsaux2!INTE_CAR_NUMERO) + " and vcha_ser_serie_id = '" + rsaux2!VCHA_SER_SERIE_ID + "'", cnn, adOpenDynamic, adLockOptimistic
         list_item.SubItems(15) = Format(IIf(IsNull(rsaux3!FLOA_sAL_IMPORTE), 0, rsaux3!FLOA_sAL_IMPORTE) - var_importe_total, "###,##0.00")
         rsaux3.Close
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
      MsgBox "El cliente no tiene descuentos por aplicar", vbOKOnly, "ATENCION"
      txt_clave_cliente = ""
      txt_nombre_cliente = ""
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_nota_credito_saldos_descuento_financiero)
End Sub

Private Sub txt_clave_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_clave_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub
