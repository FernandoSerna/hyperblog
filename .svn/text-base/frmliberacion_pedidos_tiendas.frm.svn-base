VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmliberacion_pedidos_tiendas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liberación de pedidos de tiendas"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_pedido_resurtir 
      Height          =   1110
      Left            =   810
      TabIndex        =   3
      Top             =   195
      Width           =   2310
      Begin VB.TextBox txt_orden_surtido 
         Height          =   345
         Left            =   195
         TabIndex        =   4
         Top             =   555
         Width           =   1905
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   " Orden de Surtido"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   5
         Top             =   120
         Width           =   2235
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   " Saldo del cliente "
      Height          =   2685
      Left            =   7065
      TabIndex        =   11
      Top             =   495
      Width           =   4395
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Real:"
         Height          =   195
         Left            =   165
         TabIndex        =   27
         Top             =   1785
         Width           =   375
      End
      Begin VB.Label lbl_real 
         Alignment       =   1  'Right Justify
         Caption         =   "999,999,999.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   1020
         TabIndex        =   26
         Top             =   1665
         Width           =   2655
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Disponible:"
         Height          =   195
         Left            =   165
         TabIndex        =   25
         Top             =   1140
         Width           =   780
      End
      Begin VB.Label lbl_nombre 
         Height          =   375
         Left            =   855
         TabIndex        =   16
         Top             =   705
         Width           =   3300
      End
      Begin VB.Label lbl_referencia 
         Height          =   375
         Left            =   1095
         TabIndex        =   15
         Top             =   300
         Width           =   3180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   375
         Left            =   165
         TabIndex        =   14
         Top             =   705
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Referencia:"
         Height          =   375
         Left            =   165
         TabIndex        =   13
         Top             =   300
         Width           =   825
      End
      Begin VB.Label lbl_saldo 
         Alignment       =   1  'Right Justify
         Caption         =   "999,999,999.00"
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
         Height          =   435
         Left            =   1020
         TabIndex        =   12
         Top             =   1020
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3960
      Left            =   75
      TabIndex        =   9
      Top             =   3195
      Width           =   11415
      Begin MSComctlLib.ListView lv_saldos 
         Height          =   3120
         Left            =   75
         TabIndex        =   10
         Top             =   225
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   5503
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Orden Surtido"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Importe OS"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Liberada"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Pedido Credito"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Autorizado"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label11 
         Caption         =   "Pedidos de crédito sin surtir"
         Height          =   300
         Left            =   8760
         TabIndex        =   24
         Top             =   3555
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF0000&
         Height          =   285
         Left            =   8415
         TabIndex        =   23
         Top             =   3510
         Width           =   270
      End
      Begin VB.Label Label9 
         Caption         =   "Pedidos sin liberar y con saldo"
         Height          =   300
         Left            =   5865
         TabIndex        =   22
         Top             =   3555
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H0000FF00&
         Height          =   285
         Left            =   5520
         TabIndex        =   21
         Top             =   3510
         Width           =   270
      End
      Begin VB.Label Label7 
         Caption         =   "Pedidos liberados sin surtir"
         Height          =   300
         Left            =   3225
         TabIndex        =   20
         Top             =   3555
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H000000FF&
         Height          =   285
         Left            =   2880
         TabIndex        =   19
         Top             =   3510
         Width           =   270
      End
      Begin VB.Label Label5 
         Caption         =   "Pedidos sin saldo para liberar"
         Height          =   300
         Left            =   435
         TabIndex        =   18
         Top             =   3555
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000007&
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   3510
         Width           =   270
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Clientes "
      Height          =   2670
      Left            =   60
      TabIndex        =   7
      Top             =   495
      Width           =   6945
      Begin MSComctlLib.ListView lv_clientes 
         Height          =   2355
         Left            =   75
         TabIndex        =   8
         Top             =   225
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   4154
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
            Text            =   "Referencia"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   8643
         EndProperty
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmliberacion_pedidos_tiendas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11205
      Picture         =   "frmliberacion_pedidos_tiendas.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Left            =   45
      Picture         =   "frmliberacion_pedidos_tiendas.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cargar Pedidos"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   0
      TabIndex        =   6
      Top             =   285
      Width           =   11535
   End
End
Attribute VB_Name = "frmliberacion_pedidos_tiendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_fecha As Integer
Dim var_almacen As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_pedido_factura_ceros As Integer



Private Sub cmd_imprimir_Click()
   Me.frm_pedido_resurtir.Visible = True
   Me.txt_orden_surtido = ""
   Me.txt_orden_surtido.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   lv_saldos.ListItems.Clear
   rs.Open "select distinct vcha_cli_referencia, vcha_cli_nombre from vw_pedidos_tiendas where  char_ped_estatus <> 'E' and char_ped_estatus <> 'C' and inte_ped_autorizo = 1", cnn, adOpenDynamic, adLockOptimistic
   var_i = 0
   lv_clientes.ListItems.Clear
   While Not rs.EOF
         var_i = var_i + 1
         Set list_item = lv_clientes.ListItems.Add(, , Trim(rs!VCHA_CLI_REFERENCIA))
         list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         rs.MoveNext
   Wend
   rs.Close
   lbl_saldo = Format(0, "###,###,##0.00")
End Sub


Private Sub Form_Load()
   Top = 0
   Left = 0
   If cnn_clientes_tiendas.State = 0 Then
      cnn_clientes_tiendas.Open var_conexion_pedidos_tiendas
      cnn_clientes_tiendas.CursorLocation = adUseClient
   End If
   Me.frm_pedido_resurtir.Visible = False
   cnn.CommandTimeout = 360
   rs.Open "select distinct vcha_cli_referencia, vcha_cli_nombre from vw_pedidos_tiendas WITH (NOLOCK) where  char_ped_estatus <> 'E' and char_ped_estatus <> 'C' and inte_ped_autorizo = 1", cnn, adOpenDynamic, adLockOptimistic
   var_i = 0
   While Not rs.EOF
         var_i = var_i + 1
         Set list_item = lv_clientes.ListItems.Add(, , Trim(rs!VCHA_CLI_REFERENCIA))
         list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         rs.MoveNext
   Wend
   rs.Close
   lbl_saldo = Format(0, "###,###,##0.00")
   lbl_real = Format(0, "###,###,##0.00")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_clientes_ItemClick(ByVal item As MSComctlLib.ListItem)
'   lbl_saldo = Format(0, "###,###,##0.00")
'   Me.lv_saldos.ListItems.Clear
End Sub

Private Sub lv_clientes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.lbl_referencia = Trim(Me.lv_clientes.selectedItem)
      Me.lbl_nombre = Trim(Me.lv_clientes.selectedItem.SubItems(1))
      
      cnn.CommandTimeout = 6000
      cnn_clientes_tiendas.CommandTimeout = 6000
      'MsgBox cnn_clientes_tiendas.ConnectionString
      rs.Open "select VCHA_SAL_REFERENCIA, NUMB_SAL_IMPORTE_DISPONIBLE, NUMB_SAL_IMPORTE from tb_saldo where vcha_sal_referencia = '" + Trim(Me.lv_clientes.selectedItem) + "'", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         lbl_saldo = Format(IIf(IsNull(rs!NUMB_SAL_IMPORTE_DISPONIBLE), 0, rs!NUMB_SAL_IMPORTE_DISPONIBLE), "###,###,##0.00")
         lbl_real = Format(IIf(IsNull(rs!NUMB_SAL_IMPORTE), 0, rs!NUMB_SAL_IMPORTE), "###,###,##0.00")
      Else
         lbl_saldo = Format(0, "###,###,##0.00")
         lbl_real = Format(0, "###,###,##0.00")
      End If
      rs.Close
      rs.Open "select vcha_Cli_clave_id from tb_clientes where vcha_cli_referencia = '" + Trim(Me.lv_clientes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
      var_clave_cliente = ""
      If Not rs.EOF Then
         While Not rs.EOF
               If var_clave_cliente = "" Then
                  var_clave_cliente = " vcha_cli_clave_id = '" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'"
               Else
                  var_clave_cliente = var_clave_cliente + " or vcha_cli_clave_id = '" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'"
               End If
               rs.MoveNext
         Wend
      End If
      rs.Close
      rs.Open "select * from vw_pedidos_tiendas with (nolock) where char_ped_estatus <> 'E' and char_ped_estatus <> 'C' and (" + var_clave_cliente + ") and inte_ped_autorizo = 1", cnn, adOpenDynamic, adLockOptimistic
      var_i = 0
      lv_saldos.ListItems.Clear
      While Not rs.EOF
            rsaux9.Open "SELECT FLOA_ORS_CANTIDAD_SURTIR FROM VW_CANTIDAD_SURTIR_ORDEN_SURTIDO WITH (NOLOCK)  WHERE INTE_ORS_ORDEN_SURTIDO = " + CStr(rs!INTE_ORS_ORDEN_SURTIDO), cnn, adOpenDynamic, adLockOptimistic
            If rsaux9!FLOA_ORS_CANTIDAD_SURTIR > 0 Then
               var_i = var_i + 1
               Set list_item = lv_saldos.ListItems.Add(, , Trim(rs!inte_ped_numero))
               list_item.SubItems(1) = IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), "", rs!INTE_ORS_ORDEN_SURTIDO)
               list_item.SubItems(2) = Format(IIf(IsNull(rs!importe_pedido), 0, rs!importe_pedido) + IIf(IsNull(rs!importe_seguro), 0, rs!importe_seguro) + IIf(IsNull(rs!importe_paqueteria), 0, rs!importe_paqueteria) + IIf(IsNull(rs!floa_paq_costo_referencia), 0, rs!floa_paq_costo_referencia), "###,###,##0.00")
               list_item.SubItems(3) = IIf(IsNull(rs!inte_ors_liberada), 0, rs!inte_ors_liberada)
               list_item.SubItems(4) = IIf(IsNull(rs!inte_ped_pedido_credito), 0, rs!inte_ped_pedido_credito)
               list_item.SubItems(5) = Format(IIf(IsNull(rs!inte_ped_autorizo), 0, rs!inte_ped_autorizo), "###,###,##0.00")
               If rs!inte_ors_liberada = 1 Then
                  lv_saldos.ListItems.item(var_i).Selected = True
                  lv_saldos.selectedItem.ForeColor = &HFF&
                  lv_saldos.ListItems.item(var_i).ListSubItems(1).ForeColor = &HFF&
                  lv_saldos.ListItems.item(var_i).ListSubItems(2).ForeColor = &HFF&
                  lv_saldos.ListItems.item(var_i).ListSubItems(3).ForeColor = &HFF&
                  lv_saldos.ListItems.item(var_i).ListSubItems(4).ForeColor = &HFF&
                  lv_saldos.ListItems.item(var_i).ListSubItems(5).ForeColor = &HFF&
                  If rs!inte_ped_pedido_credito = 1 Then
                     lv_saldos.ListItems.item(var_i).Selected = True
                     lv_saldos.selectedItem.ForeColor = &HFF0000
                     lv_saldos.ListItems.item(var_i).ListSubItems(1).ForeColor = &HFF0000
                     lv_saldos.ListItems.item(var_i).ListSubItems(2).ForeColor = &HFF0000
                     lv_saldos.ListItems.item(var_i).ListSubItems(3).ForeColor = &HFF0000
                     lv_saldos.ListItems.item(var_i).ListSubItems(4).ForeColor = &HFF0000
                     lv_saldos.ListItems.item(var_i).ListSubItems(5).ForeColor = &HFF0000
                  End If
               Else
                  If IIf(IsNull(rs!importe_pedido), 0, rs!importe_pedido) + IIf(IsNull(rs!importe_seguro), 0, rs!importe_seguro) + IIf(IsNull(rs!importe_paqueteria), 0, rs!importe_paqueteria) + IIf(IsNull(rs!floa_paq_costo_referencia), 0, rs!floa_paq_costo_referencia) <= CDbl(lbl_saldo) + 0.1 Then
                     lv_saldos.ListItems.item(var_i).Selected = True
                     lv_saldos.selectedItem.ForeColor = &HC000&
                     lv_saldos.ListItems.item(var_i).ListSubItems(1).ForeColor = &HC000&
                     lv_saldos.ListItems.item(var_i).ListSubItems(2).ForeColor = &HC000&
                     lv_saldos.ListItems.item(var_i).ListSubItems(3).ForeColor = &HC000&
                     lv_saldos.ListItems.item(var_i).ListSubItems(4).ForeColor = &HC000&
                     lv_saldos.ListItems.item(var_i).ListSubItems(5).ForeColor = &HC000&
                  End If
                  If rs!inte_ped_pedido_credito = 1 Then
                     lv_saldos.ListItems.item(var_i).Selected = True
                     lv_saldos.selectedItem.ForeColor = &HFF0000
                     lv_saldos.ListItems.item(var_i).ListSubItems(1).ForeColor = &HFF0000
                     lv_saldos.ListItems.item(var_i).ListSubItems(2).ForeColor = &HFF0000
                     lv_saldos.ListItems.item(var_i).ListSubItems(3).ForeColor = &HFF0000
                     lv_saldos.ListItems.item(var_i).ListSubItems(4).ForeColor = &HFF0000
                     lv_saldos.ListItems.item(var_i).ListSubItems(5).ForeColor = &HFF0000
                  End If
               End If
            End If
            rsaux9.Close
            rs.MoveNext
      Wend
      rs.Close
      If Me.lv_saldos.ListItems.Count > 0 Then
         Me.lv_saldos.ListItems.item(1).Selected = True
         Me.lv_saldos.SetFocus
      End If
   End If
End Sub

Private Sub lv_saldos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_saldos, ColumnHeader)
End Sub


Private Sub lv_saldos_KeyDown(KeyCode As Integer, Shift As Integer)
   'On Error GoTo salir:
   If KeyCode = 117 Then
      If Me.lv_saldos.ListItems.Count > 0 Then
         Me.txt_orden_surtido = Me.lv_saldos.selectedItem.SubItems(1)
         If IsNumeric(Me.txt_orden_surtido) Then
            cnn.BeginTrans
            rs.Open "select max(INTE_TEM_CONSECUTIVO) as consecutivo from tb_temp_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!consecutivo), 0, rs!consecutivo) + 1
            Else
               var_consecutivo = 1
            End If
            rs.Close
            rs.Open "insert into tb_temp_orden_surtido (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            Cadena = "INSERT INTO [TB_TEMP_ORDEN_SURTIDO] ([INTE_TEM_CONSECUTIVO], [CHAR_TPE_TIPO_PEDIDO], [INTE_PED_NUMERO], [VCHA_ALM_ALMACEN_ID], [INTE_ORS_ORDEN_SURTIDO], [DTIM_ORS_FECHA_CARGA], [DTIM_ORS_FECHA_CADUCA], [VCHA_ESB_NOMBRE], [VCHA_ESB_ESTABLECIMIENTO_ID], [VCHA_CLI_CLAVE_ID], [VCHA_CLI_NOMBRE], "
            Cadena = Cadena + "[CHAR_PRI_PRIORIDAD_ID], [FLOA_ORS_DESCUENTO_1], [FLOA_ORS_DESCUENTO_2], [VCHA_ART_ARTICULO_ID], [VCHA_ART_NOMBRE_ESPAÑOL], [FLOA_ORS_PRECIO], [FLOA_ORS_CANTIDAD_SURTIR], [VCHA_AGE_NOMBRE], [VCHA_RUT_NOMBRE], [INTE_PED_DIAS_CONDICIONES], [INTE_PED_DIAS_CADUCIDAD], [MONE_ART_COSTO_ESTANDAR], [FLOA_ORS_CANTIDAD_SURTIDA], "
            Cadena = Cadena + "[VCHA_TIT_TITULAR_ID], [VCHA_TIT_NOMBRE], [VCHA_RUT_RUTA_ID], [VCHA_AGE_AGENTE_ID], [CHAR_ORS_ESTATUS], [VCHA_MOV_MOVIMIENTO_ID], [VCHA_CLI_EMAIL], [FLOA_TPE_IVA], [ALMACEN_AGENTE], [INTE_TCL_EMPACADO], [VCHA_TCL_ALMACEN_EMPAQUE], [FLOA_ORS_CANTIDAD_EMPACADA], [VCHA_PLA_PAZO_ID], [INTE_CLI_AGRUPADOR], [VCHA_MON_MONEDA_ID],"
            Cadena = Cadena + "[VCHA_ALM_NOMBRE], [FLOA_ORS_PROMOCION_1], [FLOA_ORS_PROMOCION_2], [INTE_ORS_FACTURA_CEROS], [CHAR_PED_TIPO], [VCHA_UOR_UNIDAD_ID], [INTE_PED_REFERENCIA],  [VCHA_EXI_UBICACION], [FLOA_ORS_CANTIDAD_NEGADA], [VCHA_TEM_CODIGO_BARRAS], [INTE_PED_PEDIDO_CREDITO], [VCHA_CLI_TELEFONO]) "
            Cadena = Cadena + " select " + CStr(var_consecutivo) + ", CHAR_TPE_TIPO_PEDIDO_ID, INTE_PED_NUMERO, VCHA_ALM_ALMACEN_ID, INTE_ORS_ORDEN_SURTIDO, DTIM_ORS_FECHA_CARGA, DTIM_ORS_FECHA_CADUCA, VCHA_ESB_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, CHAR_PRI_PRIORIDAD_ID, FLOA_ORS_DESCUENTO_1, FLOA_ORS_DESCUENTO_2, VCHA_ART_ARTICULO_ID, "
            Cadena = Cadena + " VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ORS_PRECIO, FLOA_ORS_CANTIDAD_SURTIR, VCHA_AGE_NOMBRE, VCHA_RUT_NOMBRE, INTE_PED_DIAS_CONDICIONES, INTE_PED_DIAS_CADUCIDAD, MONE_ART_COSTO_ESTANDAR, FLOA_ORS_CANTIDAD_SURTIDA, VCHA_TIT_TITULAR_ID, VCHA_TIT_NOMBRE, VCHA_RUT_RUTA_ID, VCHA_AGE_AGENTE_ID, CHAR_ORS_ESTATUS, VCHA_MOV_MOVIMIENTO_ID, VCHA_CLI_EMAIL, FLOA_TPE_IVA,"
            Cadena = Cadena + " ALMACEN_AGENTE, INTE_TCL_EMPACADO, VCHA_TCL_ALMACEN_EMPAQUE, FLOA_ORS_CANTIDAD_EMPACADA, VCHA_PLA_PlAZO_ID, INTE_CLI_AGRUPADOR, VCHA_MON_MONEDA_ID, VCHA_ALM_NOMBRE, FLOA_ORS_PROMOCION_1, FLOA_ORS_PROMOCION_2, INTE_ORS_FACTURA_CEROS, CHAR_PED_TIPO, VCHA_UOR_UNIDAD_ID, INTE_PED_REFERENCIA, VCHA_EXI_UBICACION, FLOA_ORS_CANTIDAD_NEGADA, " + Format(Me.txt_orden_surtido, "##########") + ",ISNULL(INTE_PED_PEDIDO_CREDITO,0), VCHA_CLI_TELEFONO from vw_orden_surtido where inte_ors_orden_surtido = " + CStr(CDbl(Me.txt_orden_surtido))
            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            If var_empresa = "18" Then
               Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido_nueva_textilera.rpt")
            Else
               Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido_nueva_FT.rpt")
            End If
            reporte.RecordSelectionFormula = "{TB_TEMP_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO} = " + CStr(CDbl(Me.txt_orden_surtido)) + " AND {TB_TEMP_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {TB_TEMP_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR} > 0"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Orden de Surtido"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            rs.Open "delete from tb_temp_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "Número de orden de surtido incorrecta", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   If KeyCode = 116 Then
      Dim var_pedido_tienada As Double
     
      Dim var_importe_pedido_tienda As Double
      Dim var_importe_paqueteria_tienda As Double
      Dim var_importe_seguro_tienda As Double
      Dim var_importe_referencia_tienda As Double
      Dim var_importe_total_tienda As Double
      Dim var_numero_factura_tienda As Double
   
      Dim var_clave_cliente_tienda As String
      Dim var_referencia_cliente_tienda As String
      Dim var_agente_cliente_tienda As String
      Dim var_canal_cliente_tienda As String
      If Trim(Me.lbl_referencia) <> "" Then
         If CDbl(Me.lv_saldos.selectedItem.SubItems(2)) <= CDbl(Me.lbl_saldo) + 0.1 Then
            If CDbl(Me.lv_saldos.selectedItem.SubItems(2)) <= CDbl(Me.lbl_real) + 0.1 Then
               If CDbl(Me.lv_saldos.selectedItem.SubItems(3)) = 1 Then
                  MsgBox "La orden de surtido ya fue liberada con anterioridad", vbOKOnly, "ATENCION"
               Else
                  var_si = MsgBox("¿Desea liberar la orden de surtido?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     var_si = MsgBox("Confirmar la autorización de la orden de surtido", vbYesNo, "ATENCION")
                     If var_si = 6 Then
                        cnn.CommandTimeout = 6000
                        cnn.BeginTrans
                        rs.Open "UPDATE TB_ENC_ORDEN_SURTIDO SET INTE_ORS_LIBERADA = 1, DTIM_ORS_FECHA_LIBERACION = GETDATE(), VCHA_ORS_USUARIO_LIBERACION = '" + var_clave_usuario_global + "', VCHA_ORS_MAQUINA_LIBERACION = '" + fun_NombrePc + "' WHERE INTE_ORS_ORDEN_SURTIDO = " + Me.lv_saldos.selectedItem.SubItems(1), cnn, adOpenDynamic, adLockOptimistic
                        If rs.State = 1 Then
                           rs.Close
                        End If
                        rs.Open "select * from tb_enc_orden_surtido WHERE INTE_ORS_ORDEN_SURTIDO = " + Me.lv_saldos.selectedItem.SubItems(1), cnn, adOpenDynamic, adLockOptimistic
                        var_clave_cliente_tienda = rs!vcha_cli_clave_id
                        rs.Close
                        rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + var_clave_cliente_tienda + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_referencia_cliente_tienda = Trim(rs!VCHA_CLI_REFERENCIA)
                        var_agente_cliente_tienda = rs!VCHA_AGE_AGENTE_ID
                        var_canal_cliente_tienda = rs!vcha_can_canal_venta_id
                        rs.Close
                        If Round(CDbl(Me.lv_saldos.selectedItem.SubItems(2)), 2) > 0 Then
                           rsaux7.Open "call SP_AGREGA_CARGO ('" + var_canal_cliente_tienda + "','" + var_agente_cliente_tienda + "', " + CStr(Trim(Me.lv_saldos.selectedItem)) + ",'" + Trim(var_referencia_cliente_tienda) + "',0," + CStr(Round(CDbl(Me.lv_saldos.selectedItem.SubItems(2)), 2)) + ", TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'VA')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                           If CDbl(Me.lv_saldos.selectedItem.SubItems(2)) > CDbl(Me.lbl_saldo) Then
                              VAR_DIFERENCIA_CENTAVOS = CDbl(Me.lv_saldos.selectedItem.SubItems(2)) - CDbl(Me.lbl_saldo)
                              rsaux7.Open "CALL SP_AGREGA_ABONO('" + Trim(var_referencia_cliente_tienda) + "',0.00," + CStr(VAR_DIFERENCIA_CENTAVOS) + ",SYSDATE,SYSDATE,'" + CStr(Trim(Me.lv_saldos.selectedItem)) + "','','AD','')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                           End If
                        End If
                        rs.Open "select VCHA_SAL_REFERENCIA, NUMB_SAL_IMPORTE_DISPONIBLE from tb_saldo where vcha_sal_referencia = '" + Trim(Me.lv_clientes.selectedItem) + "'", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           lbl_saldo = Format(IIf(IsNull(rs!NUMB_SAL_IMPORTE_DISPONIBLE), 0, rs!NUMB_SAL_IMPORTE_DISPONIBLE), "###,###,##0.00")
                        Else
                           lbl_saldo = Format(0, "###,###,##0.00")
                        End If
                        rs.Close
                        cnn.CommitTrans
                        lv_saldos.ListItems.Clear
                        rs.Open "select * from vw_pedidos_tiendas with (nolock) where char_ped_estatus <> 'E' and char_ped_estatus <> 'C' and vcha_cli_clave_id = '" + var_clave_cliente_tienda + "' and inte_ped_autorizo = 1", cnn, adOpenDynamic, adLockOptimistic
                        var_i = 0
                        While Not rs.EOF
                              var_i = var_i + 1
                              Set list_item = lv_saldos.ListItems.Add(, , Trim(rs!inte_ped_numero))
                              list_item.SubItems(1) = IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), "", rs!INTE_ORS_ORDEN_SURTIDO)
                              list_item.SubItems(2) = Format(IIf(IsNull(rs!importe_pedido), 0, rs!importe_pedido) + IIf(IsNull(rs!importe_seguro), 0, rs!importe_seguro) + IIf(IsNull(rs!importe_paqueteria), 0, rs!importe_paqueteria) + IIf(IsNull(rs!floa_paq_costo_referencia), 0, rs!floa_paq_costo_referencia), "###,###,##0.00")
                              list_item.SubItems(3) = IIf(IsNull(rs!inte_ors_liberada), 0, rs!inte_ors_liberada)
                              list_item.SubItems(4) = IIf(IsNull(rs!inte_ped_pedido_credito), 0, rs!inte_ped_pedido_credito)
                              list_item.SubItems(5) = Format(IIf(IsNull(rs!inte_ped_autorizo), 0, rs!inte_ped_autorizo), "###,###,##0.00")
                              If rs!inte_ors_liberada = 1 Then
                                 lv_saldos.ListItems.item(var_i).Selected = True
                                 lv_saldos.selectedItem.ForeColor = &HFF&
                                 lv_saldos.ListItems.item(var_i).ListSubItems(1).ForeColor = &HFF&
                                 lv_saldos.ListItems.item(var_i).ListSubItems(2).ForeColor = &HFF&
                                 lv_saldos.ListItems.item(var_i).ListSubItems(3).ForeColor = &HFF&
                                 lv_saldos.ListItems.item(var_i).ListSubItems(4).ForeColor = &HFF&
                                 lv_saldos.ListItems.item(var_i).ListSubItems(5).ForeColor = &HFF&
                                 If rs!inte_ped_pedido_credito = 1 Then
                                    lv_saldos.ListItems.item(var_i).Selected = True
                                    lv_saldos.selectedItem.ForeColor = &HFF0000
                                    lv_saldos.ListItems.item(var_i).ListSubItems(1).ForeColor = &HFF0000
                                    lv_saldos.ListItems.item(var_i).ListSubItems(2).ForeColor = &HFF0000
                                    lv_saldos.ListItems.item(var_i).ListSubItems(3).ForeColor = &HFF0000
                                    lv_saldos.ListItems.item(var_i).ListSubItems(4).ForeColor = &HFF0000
                                    lv_saldos.ListItems.item(var_i).ListSubItems(5).ForeColor = &HFF0000
                                 End If
                              Else
                                 If IIf(IsNull(rs!importe_pedido), 0, rs!importe_pedido) + IIf(IsNull(rs!importe_seguro), 0, rs!importe_seguro) + IIf(IsNull(rs!importe_paqueteria), 0, rs!importe_paqueteria) + IIf(IsNull(rs!floa_paq_costo_referencia), 0, rs!floa_paq_costo_referencia) < CDbl(lbl_saldo) Then
                                    lv_saldos.ListItems.item(var_i).Selected = True
                                    lv_saldos.selectedItem.ForeColor = &HC000&
                                    lv_saldos.ListItems.item(var_i).ListSubItems(1).ForeColor = &HC000&
                                    lv_saldos.ListItems.item(var_i).ListSubItems(2).ForeColor = &HC000&
                                    lv_saldos.ListItems.item(var_i).ListSubItems(3).ForeColor = &HC000&
                                    lv_saldos.ListItems.item(var_i).ListSubItems(4).ForeColor = &HC000&
                                    lv_saldos.ListItems.item(var_i).ListSubItems(5).ForeColor = &HC000&
                                 End If
                                 If rs!inte_ped_pedido_credito = 1 Then
                                    lv_saldos.ListItems.item(var_i).Selected = True
                                    lv_saldos.selectedItem.ForeColor = &HFF0000
                                    lv_saldos.ListItems.item(var_i).ListSubItems(1).ForeColor = &HFF0000
                                    lv_saldos.ListItems.item(var_i).ListSubItems(2).ForeColor = &HFF0000
                                    lv_saldos.ListItems.item(var_i).ListSubItems(3).ForeColor = &HFF0000
                                    lv_saldos.ListItems.item(var_i).ListSubItems(4).ForeColor = &HFF0000
                                    lv_saldos.ListItems.item(var_i).ListSubItems(5).ForeColor = &HFF0000
                                End If
                             End If
                             rs.MoveNext
                        Wend
                        rs.Close
                        If Me.lv_saldos.ListItems.Count > 0 Then
                           Me.lv_saldos.SetFocus
                        End If
                  
                  
                  
                        cnn.BeginTrans
                        rs.Open "select max(INTE_TEM_CONSECUTIVO) as consecutivo from tb_temp_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_consecutivo = IIf(IsNull(rs!consecutivo), 0, rs!consecutivo) + 1
                        Else
                           var_consecutivo = 1
                        End If
                        rs.Close
                        rs.Open "insert into tb_temp_orden_surtido (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                        cnn.CommitTrans
                        Cadena = "INSERT INTO [TB_TEMP_ORDEN_SURTIDO] ([INTE_TEM_CONSECUTIVO], [CHAR_TPE_TIPO_PEDIDO], [INTE_PED_NUMERO], [VCHA_ALM_ALMACEN_ID], [INTE_ORS_ORDEN_SURTIDO], [DTIM_ORS_FECHA_CARGA], [DTIM_ORS_FECHA_CADUCA], [VCHA_ESB_NOMBRE], [VCHA_ESB_ESTABLECIMIENTO_ID], [VCHA_CLI_CLAVE_ID], [VCHA_CLI_NOMBRE], "
                        Cadena = Cadena + "[CHAR_PRI_PRIORIDAD_ID], [FLOA_ORS_DESCUENTO_1], [FLOA_ORS_DESCUENTO_2], [VCHA_ART_ARTICULO_ID], [VCHA_ART_NOMBRE_ESPAÑOL], [FLOA_ORS_PRECIO], [FLOA_ORS_CANTIDAD_SURTIR], [VCHA_AGE_NOMBRE], [VCHA_RUT_NOMBRE], [INTE_PED_DIAS_CONDICIONES], [INTE_PED_DIAS_CADUCIDAD], [MONE_ART_COSTO_ESTANDAR], [FLOA_ORS_CANTIDAD_SURTIDA], "
                        Cadena = Cadena + "[VCHA_TIT_TITULAR_ID], [VCHA_TIT_NOMBRE], [VCHA_RUT_RUTA_ID], [VCHA_AGE_AGENTE_ID], [CHAR_ORS_ESTATUS], [VCHA_MOV_MOVIMIENTO_ID], [VCHA_CLI_EMAIL], [FLOA_TPE_IVA], [ALMACEN_AGENTE], [INTE_TCL_EMPACADO], [VCHA_TCL_ALMACEN_EMPAQUE], [FLOA_ORS_CANTIDAD_EMPACADA], [VCHA_PLA_PAZO_ID], [INTE_CLI_AGRUPADOR], [VCHA_MON_MONEDA_ID],"
                        Cadena = Cadena + "[VCHA_ALM_NOMBRE], [FLOA_ORS_PROMOCION_1], [FLOA_ORS_PROMOCION_2], [INTE_ORS_FACTURA_CEROS], [CHAR_PED_TIPO], [VCHA_UOR_UNIDAD_ID], [INTE_PED_REFERENCIA],  [VCHA_EXI_UBICACION], [FLOA_ORS_CANTIDAD_NEGADA], [VCHA_TEM_CODIGO_BARRAS],[INTE_PED_PEDIDO_CREDITO]) "
                        Cadena = Cadena + " select " + CStr(var_consecutivo) + ", CHAR_TPE_TIPO_PEDIDO_ID, INTE_PED_NUMERO, VCHA_ALM_ALMACEN_ID, INTE_ORS_ORDEN_SURTIDO, DTIM_ORS_FECHA_CARGA, DTIM_ORS_FECHA_CADUCA, VCHA_ESB_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, CHAR_PRI_PRIORIDAD_ID, FLOA_ORS_DESCUENTO_1, FLOA_ORS_DESCUENTO_2, VCHA_ART_ARTICULO_ID, "
                        Cadena = Cadena + " VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ORS_PRECIO, FLOA_ORS_CANTIDAD_SURTIR, VCHA_AGE_NOMBRE, VCHA_RUT_NOMBRE, INTE_PED_DIAS_CONDICIONES, INTE_PED_DIAS_CADUCIDAD, MONE_ART_COSTO_ESTANDAR, FLOA_ORS_CANTIDAD_SURTIDA, VCHA_TIT_TITULAR_ID, VCHA_TIT_NOMBRE, VCHA_RUT_RUTA_ID, VCHA_AGE_AGENTE_ID, CHAR_ORS_ESTATUS, VCHA_MOV_MOVIMIENTO_ID, VCHA_CLI_EMAIL, FLOA_TPE_IVA,"
                        Cadena = Cadena + " ALMACEN_AGENTE, INTE_TCL_EMPACADO, VCHA_TCL_ALMACEN_EMPAQUE, FLOA_ORS_CANTIDAD_EMPACADA, VCHA_PLA_PlAZO_ID, INTE_CLI_AGRUPADOR, VCHA_MON_MONEDA_ID, VCHA_ALM_NOMBRE, FLOA_ORS_PROMOCION_1, FLOA_ORS_PROMOCION_2, INTE_ORS_FACTURA_CEROS, CHAR_PED_TIPO, VCHA_UOR_UNIDAD_ID, INTE_PED_REFERENCIA, VCHA_EXI_UBICACION, FLOA_ORS_CANTIDAD_NEGADA, " + Format(Me.lv_saldos.selectedItem.SubItems(1), "##########") + ",ISNULL(INTE_PED_PEDIDO_CREDITO,0) from vw_orden_surtido where inte_ors_orden_surtido = " + CStr(Me.lv_saldos.selectedItem.SubItems(1))
                        rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                        If var_empresa = "18" Then
                           Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido_nueva_textilera.rpt")
                        Else
                            Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido_nueva_FT.rpt")
                        End If
                        reporte.RecordSelectionFormula = "{TB_TEMP_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO} = " + CStr(Me.lv_saldos.selectedItem.SubItems(1)) + " AND {TB_TEMP_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {TB_TEMP_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR} > 0"
                        frmvistasprevias.cr.ReportSource = reporte
                        For ntablas = 1 To reporte.Database.Tables.Count
                            reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                        Next ntablas
                        frmvistasprevias.cr.ViewReport
                        frmvistasprevias.Caption = "Resurtido de Pedidos"
                        frmvistasprevias.Show 1
                        Set reporte = Nothing
                        rs.Open "delete from tb_temp_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                     End If
                  End If
               End If
            Else
               MsgBox "El saldo real del cliente es menor al importe del pedido", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El saldo disponible del cliente es menor al importe del pedido", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a seleccionado un cliente", vbOKOnly, "ATENCION"
      End If
   End If
   Exit Sub
salir:
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux7.State = 1 Then
      rsaux7.Close
   End If
   MsgBox "A surgido un error al liberar el pedido. Puede que el cliente no tenga suficiente saldo disponible", vbOKOnly, "ATENCION"
End Sub


Private Sub mes_LostFocus()
   mes.Visible = False
End Sub





Private Sub Timer1_Timer()

End Sub

Private Sub txt_orden_surtido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_orden_surtido) Then
         cnn.BeginTrans
         rs.Open "select max(INTE_TEM_CONSECUTIVO) as consecutivo from tb_temp_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_consecutivo = IIf(IsNull(rs!consecutivo), 0, rs!consecutivo) + 1
         Else
            var_consecutivo = 1
         End If
         rs.Close
         rs.Open "insert into tb_temp_orden_surtido (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
         Cadena = "INSERT INTO [TB_TEMP_ORDEN_SURTIDO] ([INTE_TEM_CONSECUTIVO], [CHAR_TPE_TIPO_PEDIDO], [INTE_PED_NUMERO], [VCHA_ALM_ALMACEN_ID], [INTE_ORS_ORDEN_SURTIDO], [DTIM_ORS_FECHA_CARGA], [DTIM_ORS_FECHA_CADUCA], [VCHA_ESB_NOMBRE], [VCHA_ESB_ESTABLECIMIENTO_ID], [VCHA_CLI_CLAVE_ID], [VCHA_CLI_NOMBRE], "
         Cadena = Cadena + "[CHAR_PRI_PRIORIDAD_ID], [FLOA_ORS_DESCUENTO_1], [FLOA_ORS_DESCUENTO_2], [VCHA_ART_ARTICULO_ID], [VCHA_ART_NOMBRE_ESPAÑOL], [FLOA_ORS_PRECIO], [FLOA_ORS_CANTIDAD_SURTIR], [VCHA_AGE_NOMBRE], [VCHA_RUT_NOMBRE], [INTE_PED_DIAS_CONDICIONES], [INTE_PED_DIAS_CADUCIDAD], [MONE_ART_COSTO_ESTANDAR], [FLOA_ORS_CANTIDAD_SURTIDA], "
         Cadena = Cadena + "[VCHA_TIT_TITULAR_ID], [VCHA_TIT_NOMBRE], [VCHA_RUT_RUTA_ID], [VCHA_AGE_AGENTE_ID], [CHAR_ORS_ESTATUS], [VCHA_MOV_MOVIMIENTO_ID], [VCHA_CLI_EMAIL], [FLOA_TPE_IVA], [ALMACEN_AGENTE], [INTE_TCL_EMPACADO], [VCHA_TCL_ALMACEN_EMPAQUE], [FLOA_ORS_CANTIDAD_EMPACADA], [VCHA_PLA_PAZO_ID], [INTE_CLI_AGRUPADOR], [VCHA_MON_MONEDA_ID],"
         Cadena = Cadena + "[VCHA_ALM_NOMBRE], [FLOA_ORS_PROMOCION_1], [FLOA_ORS_PROMOCION_2], [INTE_ORS_FACTURA_CEROS], [CHAR_PED_TIPO], [VCHA_UOR_UNIDAD_ID], [INTE_PED_REFERENCIA],  [VCHA_EXI_UBICACION], [FLOA_ORS_CANTIDAD_NEGADA], [VCHA_TEM_CODIGO_BARRAS], [INTE_PED_PEDIDO_CREDITO],[VCHA_CLI_TELEFONO]) "
         Cadena = Cadena + " select " + CStr(var_consecutivo) + ", CHAR_TPE_TIPO_PEDIDO_ID, INTE_PED_NUMERO, VCHA_ALM_ALMACEN_ID, INTE_ORS_ORDEN_SURTIDO, DTIM_ORS_FECHA_CARGA, DTIM_ORS_FECHA_CADUCA, VCHA_ESB_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, CHAR_PRI_PRIORIDAD_ID, FLOA_ORS_DESCUENTO_1, FLOA_ORS_DESCUENTO_2, VCHA_ART_ARTICULO_ID, "
         Cadena = Cadena + " VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ORS_PRECIO, FLOA_ORS_CANTIDAD_SURTIR, VCHA_AGE_NOMBRE, VCHA_RUT_NOMBRE, INTE_PED_DIAS_CONDICIONES, INTE_PED_DIAS_CADUCIDAD, MONE_ART_COSTO_ESTANDAR, FLOA_ORS_CANTIDAD_SURTIDA, VCHA_TIT_TITULAR_ID, VCHA_TIT_NOMBRE, VCHA_RUT_RUTA_ID, VCHA_AGE_AGENTE_ID, CHAR_ORS_ESTATUS, VCHA_MOV_MOVIMIENTO_ID, VCHA_CLI_EMAIL, FLOA_TPE_IVA,"
         Cadena = Cadena + " ALMACEN_AGENTE, INTE_TCL_EMPACADO, VCHA_TCL_ALMACEN_EMPAQUE, FLOA_ORS_CANTIDAD_EMPACADA, VCHA_PLA_PlAZO_ID, INTE_CLI_AGRUPADOR, VCHA_MON_MONEDA_ID, VCHA_ALM_NOMBRE, FLOA_ORS_PROMOCION_1, FLOA_ORS_PROMOCION_2, INTE_ORS_FACTURA_CEROS, CHAR_PED_TIPO, VCHA_UOR_UNIDAD_ID, INTE_PED_REFERENCIA, VCHA_EXI_UBICACION, FLOA_ORS_CANTIDAD_NEGADA, " + Format(Me.txt_orden_surtido, "##########") + ", ISNULL(INTE_PED_PEDIDO_CREDITO,0),VCHA_CLI_TELEFONO from vw_orden_surtido where inte_ors_orden_surtido = " + CStr(Me.txt_orden_surtido)
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If var_empresa = "18" Then
            Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido_nueva_textilera.rpt")
         Else
            Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido_nueva_FT.rpt")
         End If
         reporte.RecordSelectionFormula = "{TB_TEMP_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO} = " + CStr(Me.txt_orden_surtido) + " AND {TB_TEMP_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {TB_TEMP_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR} > 0"
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Orden de Surtido"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         rs.Open "delete from tb_temp_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      Else
         MsgBox "Número de orden de surtido incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_orden_surtido_LostFocus()
   Me.frm_pedido_resurtir.Visible = False
End Sub


