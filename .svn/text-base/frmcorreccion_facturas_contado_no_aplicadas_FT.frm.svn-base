VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcorreccion_facturas_contado_no_aplicadas_FT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Correccion de facturas de contado no aplicadas"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11145
      Picture         =   "frmcorreccion_facturas_contado_no_aplicadas_FT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmcorreccion_facturas_contado_no_aplicadas_FT.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmcorreccion_facturas_contado_no_aplicadas_FT.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Actualizar"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   5100
      Left            =   45
      TabIndex        =   0
      Top             =   390
      Width           =   11505
      Begin MSComctlLib.ListView lv_facturas 
         Height          =   4875
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   11340
         _ExtentX        =   20003
         _ExtentY        =   8599
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Agente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre Agente"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nombre Cliente"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Factura"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Saldo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Referencia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Subimporte"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   45
      TabIndex        =   5
      Top             =   270
      Width           =   11460
   End
End
Attribute VB_Name = "frmcorreccion_facturas_contado_no_aplicadas_FT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim list_item As ListItem

Private Sub cmd_aceptar_pedidos_Click()
   lv_facturas.ListItems.Clear
   rs.Open "select * from VW_FACTURAS_CONTADO_TIENDAS_SIN_PAGO", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_facturas.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
         list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
         list_item.SubItems(2) = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
         list_item.SubItems(3) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         list_item.SubItems(4) = IIf(IsNull(rs!inte_Car_numero), "", rs!inte_Car_numero)
         list_item.SubItems(5) = Format(IIf(IsNull(rs!dtim_car_fecha), "", rs!dtim_car_fecha), "Short Date")
          'list_item.SubItems(6) = Format(IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto), "###,###,##0.00")
          list_item.SubItems(6) = IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto)
         list_item.SubItems(7) = Format(IIf(IsNull(rs!FLOA_sAL_IMPORTE), "", rs!FLOA_sAL_IMPORTE), "###,###,##0.00")
         list_item.SubItems(8) = IIf(IsNull(rs!vcha_Cli_referencia), "", rs!vcha_Cli_referencia)
         list_item.SubItems(9) = IIf(IsNull(rs!floa_car_subimporte), "", rs!floa_car_subimporte)
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   var_si = MsgBox("¿Desea aplicar la corrección?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_si = MsgBox("Confirmar la aplicación de la corrección?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         If lv_facturas.ListItems.Count > 0 Then
            cnn.CommandTimeout = 360
            'cnn.RollbackTrans
            cnn.BeginTrans
            'rs.Open "select max(inte_car_numero) as maximo_numero from tb_encabezado_cartera where vcha_car_tipo_documento = 'PA'", cnn, adOpenDynamic, adLockOptimistic
            'If rs.EOF Then
            '   var_numero_folio = 0
            'Else
            '   var_numero_folio = IIf(IsNull(rs!maximo_numero), 0, rs!maximo_numero)
            'End If
            'rs.Close
            'MsgBox cnn_sid_quezada
            rs.Open "SELECT INTE_MAX_MAXIMO_PAGO as maximo_numero FROM TB_MAXIMO_PAGO", cnn_sid_quezada, adOpenDynamic, adLockOptimistic
            If rs.EOF Then
               var_numero_folio = 0
            Else
               var_numero_folio = IIf(IsNull(rs!maximo_numero), 0, rs!maximo_numero)
            End If
            rs.Close
            rs.Open "UPDATE TB_MAXIMO_PAGO SET INTE_MAX_MAXIMO_PAGO = INTE_MAX_MAXIMO_PAGO + 1", cnn_sid_quezada, adOpenDynamic, adLockOptimistic

            
            var_agente_cliente_tienda = lv_facturas.selectedItem
            var_clave_cliente_tienda = lv_facturas.selectedItem.SubItems(2)
            var_subimporte = CDbl(lv_facturas.selectedItem.SubItems(9))
            var_importe_neto = CDbl(lv_facturas.selectedItem.SubItems(6))
            var_referencia_cliente_tienda = Trim(Me.lv_facturas.selectedItem.SubItems(8))
            rsaux9.Open "select vcha_gac_grupo_actual_id, vcha_gre_grupo_real_id,vcha_tit_titular_id, vcha_can_Canal_Venta_id from vw_clientes where vcha_cli_clave_id = '" + var_clave_cliente_tienda + "'", cnn, adOpenDynamic, adLockOptimistic
            var_grupo_actual_tienda = IIf(IsNull(rsaux9!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux9!VCHA_GAC_GRUPO_aCTUAL_ID)
            var_grupo_real_tienda = IIf(IsNull(rsaux9!vcha_gre_grupo_real_id), "", rsaux9!vcha_gre_grupo_real_id)
            var_titular_tienda = IIf(IsNull(rsaux9!vcha_tit_titular_id), "", rsaux9!vcha_tit_titular_id)
            var_canal_venta = IIf(IsNull(rsaux9!vcha_can_canal_venta_id), "", rsaux9!vcha_can_canal_venta_id)
            var_clave_moneda = "1"
            rsaux9.Close
            var_numero_folio = var_numero_folio + 1
            Cadena = "INSERT INTO TB_ENCABEZADO_CARTERA (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, "
            Cadena = Cadena + "FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3, FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, VCHA_SER_SERIE_ID, CHAR_CAR_ESTATUS, VCHA_CAR_CHEQUE_DEPOSITO, VCHA_CAR_BANCO_CHEQUE, VCHA_CAR_DEPOSITO, VCHA_CAR_BANCO_DEPOSITO, DTIM_CAR_FECHA_DEPOSITO) values ("
            Cadena = Cadena + "'02', '23', 'PA', 'PA', 'PA', " + CStr(var_numero_folio) + ", '-', '', '', 0, getdate(), '" + var_agente_cliente_tienda + "', '" + var_grupo_actual_tienda + "', '" + var_grupo_real_tienda + "', '" + var_titular_tienda + "', '" + var_clave_cliente_tienda + "', '', 0, 16, 0, 0, 0, 0, 0, " + CStr(var_importe_neto) + ", " + CStr(var_importe_neto - var_subimporte) + ", 0, 0, 0, 0, 0, " + CStr(var_subimporte) + ", " + CStr(var_importe_neto) + ", '', '"
            Cadena = Cadena + CStr(var_clave_usuario_global) + "', '', getdate(), 0, getdate(), getdate(), '" + var_clave_moneda + "', 1, 'FAEFT', 'I','', '', '','','')"
            Text1 = Cadena
            'MsgBox Cadena
            rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            rsaux7.Open "update tb_encabezado_cartera set inte_car_pedido_credito = 0 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Car_documento = 'FA' and vcha_Ser_serie_id = 'FT' and inte_Car_numero = " + CStr(CDbl(Me.lv_facturas.selectedItem.SubItems(4))), cnn, adOpenDynamic, adLockOptimistic
            Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
            var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, "FAEFT", "FA", CDbl(Me.lv_facturas.selectedItem.SubItems(4)), "FAEFT", "PA", CDbl(var_numero_folio), 0, CDbl(var_importe_neto))
            'MsgBox var_referencia_cliente_tienda
            var_importe_neto = var_importe_neto - 0.1
            'var_importe_neto = 2947
            rsaux7.Open "call SP_AGREGA_CARGO ('" + var_canal_venta + "','" + var_agente_cliente_tienda + "', " + CStr(CDbl(Me.lv_facturas.selectedItem.SubItems(4))) + ",'" + Trim(var_referencia_cliente_tienda) + "'," + CStr(CDbl(Round(var_importe_neto, 2))) + ", 0,TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'VA')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            
            Me.lv_facturas.ListItems.Clear
            lv_facturas.ListItems.Clear
            rs.Open "select * from VW_FACTURAS_CONTADO_TIENDAS_SIN_PAGO", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = lv_facturas.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
                  list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
                  list_item.SubItems(2) = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                  list_item.SubItems(3) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                  list_item.SubItems(4) = IIf(IsNull(rs!inte_Car_numero), "", rs!inte_Car_numero)
                  list_item.SubItems(5) = Format(IIf(IsNull(rs!dtim_car_fecha), "", rs!dtim_car_fecha), "Short Date")
                   'list_item.SubItems(6) = Format(IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto), "###,###,##0.00")
                  list_item.SubItems(6) = IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto)
                  list_item.SubItems(7) = Format(IIf(IsNull(rs!FLOA_sAL_IMPORTE), "", rs!FLOA_sAL_IMPORTE), "###,###,##0.00")
                  list_item.SubItems(8) = IIf(IsNull(rs!vcha_Cli_referencia), "", rs!vcha_Cli_referencia)
                  list_item.SubItems(9) = IIf(IsNull(rs!floa_car_subimporte), "", rs!floa_car_subimporte)
                  rs.MoveNext
            Wend
            rs.Close
            
            
            MsgBox "Se a terminado de aplicar la corrección", vbOKOnly, "ATENCION"
         Else
            MsgBox "No existen documentos por corregir", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub Form_Load()
   Left = 0
   Top = 1000
   If var_empresa <> "18" Then
      If var_unidad_organizacional = "23" Then
         If cnn_clientes_tiendas.State = 0 Then
            cnn_clientes_tiendas.Open var_conexion_pedidos_tiendas
            cnn_clientes_tiendas.CursorLocation = adUseClient
         End If
      Else
         Me.Command1.Enabled = False
         Me.cmd_aceptar_pedidos.Enabled = False
      End If
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

