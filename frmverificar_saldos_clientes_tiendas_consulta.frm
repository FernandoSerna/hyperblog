VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmverificar_saldos_clientes_tiendas_consulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de surtido de tiendas"
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   65000
      Left            =   1215
      Top             =   135
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Left            =   45
      Picture         =   "frmverificar_saldos_clientes_tiendas_consulta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cargar Pedidos"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11205
      Picture         =   "frmverificar_saldos_clientes_tiendas_consulta.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmverificar_saldos_clientes_tiendas_consulta.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      Picture         =   "frmverificar_saldos_clientes_tiendas_consulta.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Buscar ordenes de surtido"
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame frm_pedido_resurtir 
      Height          =   1110
      Left            =   11520
      TabIndex        =   12
      Top             =   1575
      Width           =   2310
      Begin VB.TextBox txt_orden_surtido 
         Height          =   345
         Left            =   195
         TabIndex        =   13
         Top             =   555
         Width           =   1905
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   " Orden de Surtido"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   14
         Top             =   120
         Width           =   2235
      End
   End
   Begin MSComCtl2.MonthView mes 
      Height          =   2370
      Left            =   11595
      TabIndex        =   0
      Top             =   2160
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   69009409
      CurrentDate     =   38875
   End
   Begin VB.Frame frm_periodo 
      Height          =   1305
      Left            =   11535
      TabIndex        =   2
      Top             =   1470
      Width           =   4380
      Begin VB.CommandButton cmd_cancelar_pedidos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmverificar_saldos_clientes_tiendas_consulta.frx":0940
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   405
         Width           =   330
      End
      Begin VB.CommandButton cmd_aceptar_pedidos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   45
         Picture         =   "frmverificar_saldos_clientes_tiendas_consulta.frx":0A8A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   405
         Width           =   330
      End
      Begin VB.Frame Frame3 
         Height          =   120
         Left            =   30
         TabIndex        =   7
         Top             =   645
         Width           =   4245
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   855
         Width           =   1140
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   285
         TabIndex        =   10
         Top             =   915
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2370
         TabIndex        =   9
         Top             =   900
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   " Periodo"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   8
         Top             =   135
         Width           =   4305
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   0
      TabIndex        =   17
      Top             =   285
      Width           =   11535
   End
   Begin MSComctlLib.ListView lv_saldos 
      Height          =   6300
      Left            =   0
      TabIndex        =   18
      Top             =   450
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   11113
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Agente"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Vence"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Referencia"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Cliente"
         Object.Width           =   3616
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Pedido"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Orden S."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Importe OS"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Saldo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Liberada"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Pedido Credito"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lbl_actualizacion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4350
      TabIndex        =   28
      Top             =   15
      Width           =   4920
   End
   Begin VB.Label Label13 
      Caption         =   "Ultima actualización:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1725
      TabIndex        =   27
      Top             =   -15
      Width           =   2505
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000007&
      Height          =   285
      Left            =   420
      TabIndex        =   26
      Top             =   6885
      Width           =   270
   End
   Begin VB.Label Label5 
      Caption         =   "Pedidos sin saldo para liberar"
      Height          =   300
      Left            =   735
      TabIndex        =   25
      Top             =   6930
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Height          =   285
      Left            =   3180
      TabIndex        =   24
      Top             =   6885
      Width           =   270
   End
   Begin VB.Label Label7 
      Caption         =   "Pedidos liberados sin surtir"
      Height          =   300
      Left            =   3525
      TabIndex        =   23
      Top             =   6930
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FF00&
      Height          =   285
      Left            =   5820
      TabIndex        =   22
      Top             =   6885
      Width           =   270
   End
   Begin VB.Label Label9 
      Caption         =   "Pedidos sin liberar y con saldo"
      Height          =   300
      Left            =   6165
      TabIndex        =   21
      Top             =   6930
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FF0000&
      Height          =   285
      Left            =   8715
      TabIndex        =   20
      Top             =   6885
      Width           =   270
   End
   Begin VB.Label Label11 
      Caption         =   "Pedidos de crédito sin surtir"
      Height          =   300
      Left            =   9060
      TabIndex        =   19
      Top             =   6930
      Width           =   2175
   End
End
Attribute VB_Name = "frmverificar_saldos_clientes_tiendas_consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_fecha As Integer
Dim var_almacen As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_pedido_factura_ceros As Integer
Dim var_intervalo As Integer

Private Sub cmd_aceptar_pedidos_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
            lv_saldos.ListItems.Clear
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            rs.Open "delete from tb_saldos_clientes_tiendas", cnn, adOpenDynamic, adLockOptimistic
            'rs.Open "insert into tb_saldos_clientes_tiendas (vcha_Cli_referencia, inte_ped_numero, inte_ors_orden_surtido, floa_Sal_importe_orden, INTE_ORS_LIBERADA, INTE_PED_PEDIDO_CREDITO) select vcha_cli_referencia, inte_ped_numero, inte_ors_orden_surtido, importe, INTE_ORS_LIBERADA, INTE_PED_PEDIDO_CREDITO from vw_suma_pedidos_tiendas where DTIM_ORS_FECHA_CARGA >= " + var_fecha_inicio + " and dtim_ors_fecha_carga <= " + var_fecha_fin + " and isnull(inte_ped_pedido_credito,0) = 0", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "insert into tb_saldos_clientes_tiendas (vcha_Cli_referencia, inte_ped_numero, inte_ors_orden_surtido, floa_Sal_importe_orden, INTE_ORS_LIBERADA, INTE_PED_PEDIDO_CREDITO) select vcha_cli_referencia, inte_ped_numero, inte_ors_orden_surtido, importe, INTE_ORS_LIBERADA, INTE_PED_PEDIDO_CREDITO from vw_suma_pedidos_tiendas where DTIM_ORS_FECHA_CARGA >= " + var_fecha_inicio + " and dtim_ors_fecha_carga <= " + var_fecha_fin, cnn, adOpenDynamic, adLockOptimistic
            
            rs.Open "select VCHA_SAL_REFERENCIA, NUMB_SAL_IMPORTE_DISPONIBLE from tb_saldo", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux.Open "update tb_saldos_clientes_tiendas set floa_sal_importe_saldo = isnull(floa_sal_importe_saldo,0) + " + CStr(rs(1).Value) + " where vcha_cli_referencia = '" + IIf(IsNull(rs(0).Value), "", rs(0).Value) + "'", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            rs.Open "select * from VW_saldos_clientes_tiendas", cnn, adOpenDynamic, adLockOptimistic
            var_i = 0
            While Not rs.EOF
                  var_i = var_i + 1
                  Set list_item = lv_saldos.ListItems.Add(, , Trim(rs!VCHA_CLI_REFERENCIA))
                  list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                  list_item.SubItems(2) = IIf(IsNull(rs!inte_ped_numero), "", rs!inte_ped_numero)
                  list_item.SubItems(3) = IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), "", rs!INTE_ORS_ORDEN_SURTIDO)
                  list_item.SubItems(4) = Format(IIf(IsNull(rs!FLOA_sAL_IMPORTE_ORDEN), 0, rs!FLOA_sAL_IMPORTE_ORDEN) + IIf(IsNull(rs!importe_seguro), 0, rs!importe_seguro) + IIf(IsNull(rs!importe_paqueteria), 0, rs!importe_paqueteria) + IIf(IsNull(rs!floa_paq_costo_referencia), 0, rs!floa_paq_costo_referencia), "###,###,##0.00")
                  list_item.SubItems(5) = Format(IIf(IsNull(rs!FLOA_sAL_IMPORTE_SALDO), 0, rs!FLOA_sAL_IMPORTE_SALDO), "###,###,##0.00")
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!inte_ors_liberada), 0, rs!inte_ors_liberada), "###,###,##0.00")
                  list_item.SubItems(7) = Format(IIf(IsNull(rs!inte_ped_pedido_credito), 0, rs!inte_ped_pedido_credito), "###,###,##0.00")
                  If rs!inte_ors_liberada = 1 Then
                     lv_saldos.ListItems.Item(var_i).Selected = True
                     lv_saldos.selectedItem.ForeColor = &HFF&
                     lv_saldos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF&
                     lv_saldos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF&
                     lv_saldos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF&
                     lv_saldos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF&
                     lv_saldos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF&
                     lv_saldos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF&
                  Else
                     If CDbl(list_item.SubItems(5)) > CDbl(list_item.SubItems(4)) Then
                        lv_saldos.ListItems.Item(var_i).Selected = True
                        lv_saldos.selectedItem.ForeColor = &HC000&
                        lv_saldos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HC000&
                        lv_saldos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HC000&
                        lv_saldos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HC000&
                        lv_saldos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HC000&
                        lv_saldos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HC000&
                        lv_saldos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HC000&
                     End If
                     If rs!inte_ped_pedido_credito = 1 Then
                        lv_saldos.ListItems.Item(var_i).Selected = True
                        lv_saldos.selectedItem.ForeColor = &HFF0000
                        lv_saldos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF0000
                        lv_saldos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF0000
                        lv_saldos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF0000
                        lv_saldos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF0000
                        lv_saldos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF0000
                        lv_saldos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF0000
                     End If
                  End If
                  rs.MoveNext
            Wend
            rs.Close
            If lv_saldos.ListItems.Count > 30 Then
               lv_saldos.ColumnHeaders(5).Width = lv_saldos.ColumnHeaders(5).Width - 200
            End If
         Else
            MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
   frm_periodo.Visible = False
End Sub

Private Sub cmd_cancelar_pedidos_Click()
   frm_periodo.Visible = False
End Sub

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
   cnn.CommandTimeout = 360
   cnn.BeginTrans
   rs.Open "SELECT MAX(INTE_SAL_CONSECUTIVO) FROM TB_SALDOS_CLIENTES_TIENDAS", cnn, adOpenDynamic, adLockOptimistic
   If rs.EOF Then
      var_consecutivo = 1
   Else
      var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
   End If
   rs.Close
   var_consecutivo = var_consecutivo + 1
   rs.Open "INSERT INTO TB_SALDOS_CLIENTES_TIENDAS (INTE_SAL_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
   cnn.CommitTrans
   rs.Open "insert into tb_saldos_clientes_tiendas (vcha_Cli_referencia, inte_ped_numero, inte_ors_orden_surtido, floa_Sal_importe_orden, INTE_ORS_LIBERADA, INTE_PED_AUTORIZO, INTE_SAL_CONSECUTIVO) select vcha_cli_referencia, inte_ped_numero, inte_ors_orden_surtido, importe, INTE_ORS_LIBERADA, INTE_PED_AUTORIZO, " + CStr(var_consecutivo) + " from vw_suma_pedidos_tiendas where char_ped_estatus <> 'E' and char_ped_estatus <> 'C' and inte_ped_autorizo = 1", cnn, adOpenDynamic, adLockOptimistic
   rsaux8.Open "select vcha_cli_referencia  from tb_saldos_clientes_tiendas WHERE INTE_sAL_CONSECUTIVO = " + CStr(var_consecutivo) + " AND VCHA_CLI_REFERENCIA IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux8.EOF
         rs.Open "select VCHA_SAL_REFERENCIA, sum(NUMB_SAL_IMPORTE_DISPONIBLE)  from tb_saldo where vcha_sal_referencia = '" + Trim(rsaux8!VCHA_CLI_REFERENCIA) + "' group by VCHA_SAL_REFERENCIA", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux.Open "update tb_saldos_clientes_tiendas set floa_sal_importe_saldo = " + CStr(rs(1).Value) + " where vcha_cli_referencia = '" + IIf(IsNull(rs(0).Value), "", rs(0).Value) + "' AND INTE_SAL_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         End If
         rs.Close
         rsaux8.MoveNext
   Wend
   rsaux8.Close
   
   rs.Open "select * from VW_saldos_clientes_tiendas where char_ped_estatus = 'S' AND INTE_sAL_CONSECUTIVO = " + CStr(var_consecutivo) + " order by vcha_age_nombre, dtim_ped_fecha", cnn, adOpenDynamic, adLockOptimistic
   var_i = 0
   'Me.lv_saldos.ListItems.Clear
   'Call pro_ordena_listas(lv_saldos, lv_saldos.ColumnHeaders.Item(1))
   lbl_actualizacion = Now
   While Not rs.EOF
         var_i = var_i + 1
         Set list_item = lv_saldos.ListItems.Add(, , Trim(rs!VCHA_AGE_NOMBRE))
         list_item.SubItems(1) = IIf(IsNull(rs!dtim_ped_fecha), "", Trim(rs!dtim_ped_fecha))
         list_item.SubItems(2) = IIf(IsNull(rs!fecha_vencimiento), "", Trim(rs!fecha_vencimiento))
         list_item.SubItems(3) = IIf(IsNull(rs!VCHA_CLI_REFERENCIA), "", Trim(rs!VCHA_CLI_REFERENCIA))
         list_item.SubItems(4) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         list_item.SubItems(5) = IIf(IsNull(rs!inte_ped_numero), "", rs!inte_ped_numero)
         list_item.SubItems(6) = IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), "", rs!INTE_ORS_ORDEN_SURTIDO)
         list_item.SubItems(7) = Format(IIf(IsNull(rs!FLOA_sAL_IMPORTE_ORDEN), 0, rs!FLOA_sAL_IMPORTE_ORDEN) + IIf(IsNull(rs!importe_seguro), 0, rs!importe_seguro) + IIf(IsNull(rs!importe_paqueteria), 0, rs!importe_paqueteria) + IIf(IsNull(rs!floa_paq_costo_referencia), 0, rs!floa_paq_costo_referencia), "###,###,##0.00")
         list_item.SubItems(8) = Format(IIf(IsNull(rs!FLOA_sAL_IMPORTE_SALDO), 0, rs!FLOA_sAL_IMPORTE_SALDO), "###,###,##0.00")
         list_item.SubItems(9) = Format(IIf(IsNull(rs!inte_ors_liberada), 0, rs!inte_ors_liberada), "###,###,##0.00")
         
         If rs!inte_ors_liberada = 1 Then
            lv_saldos.ListItems.Item(var_i).Selected = True
            lv_saldos.selectedItem.ForeColor = &HFF&
            lv_saldos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF&
            lv_saldos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF&
            lv_saldos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF&
            lv_saldos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF&
            lv_saldos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF&
            lv_saldos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF&
            lv_saldos.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HFF&
            lv_saldos.ListItems.Item(var_i).ListSubItems(8).ForeColor = &HFF&
            lv_saldos.ListItems.Item(var_i).ListSubItems(9).ForeColor = &HFF&
            If rs!inte_ped_pedido_credito = 1 Then
               lv_saldos.ListItems.Item(var_i).Selected = True
               lv_saldos.selectedItem.ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(8).ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(9).ForeColor = &HFF0000
            End If
         Else
            If IIf(IsNull(rs!FLOA_sAL_IMPORTE_ORDEN), 0, rs!FLOA_sAL_IMPORTE_ORDEN) + IIf(IsNull(rs!importe_seguro), 0, rs!importe_seguro) + IIf(IsNull(rs!importe_paqueteria), 0, rs!importe_paqueteria) + IIf(IsNull(rs!floa_paq_costo_referencia), 0, rs!floa_paq_costo_referencia) <= IIf(IsNull(rs!FLOA_sAL_IMPORTE_SALDO), 0, rs!FLOA_sAL_IMPORTE_SALDO) Then
               lv_saldos.ListItems.Item(var_i).Selected = True
               lv_saldos.selectedItem.ForeColor = &HC000&
               lv_saldos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HC000&
               lv_saldos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HC000&
               lv_saldos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HC000&
               lv_saldos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HC000&
               lv_saldos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HC000&
               lv_saldos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HC000&
               lv_saldos.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HC000&
               lv_saldos.ListItems.Item(var_i).ListSubItems(8).ForeColor = &HC000&
               lv_saldos.ListItems.Item(var_i).ListSubItems(9).ForeColor = &HC000&
            End If
            If rs!inte_ped_pedido_credito = 1 Then
               lv_saldos.ListItems.Item(var_i).Selected = True
               lv_saldos.selectedItem.ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(8).ForeColor = &HFF0000
               lv_saldos.ListItems.Item(var_i).ListSubItems(9).ForeColor = &HFF0000
            End If
         End If
         
         rs.MoveNext
   Wend
   rs.Close
   rs.Open "delete from tb_saldos_clientes_tiendas WHERE INTE_SAL_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   If lv_saldos.ListItems.Count > 30 Then
      lv_saldos.ColumnHeaders(5).Width = lv_saldos.ColumnHeaders(5).Width - 200
   End If
End Sub

Private Sub Command2_Click()
   Me.txt_fin = Date
   Me.txt_inicio = Date
   Me.frm_periodo.Visible = True
   Me.txt_inicio.SetFocus
End Sub


Private Sub Form_Load()
   var_intevalo = 0
   frm_periodo.Visible = False
   mes.Visible = False
   Top = 0
   Left = 0
   If cnn_clientes_tiendas.State = 0 Then
      cnn_clientes_tiendas.Open var_conexion_pedidos_tiendas
      cnn_clientes_tiendas.CursorLocation = adUseClient
   End If
   Me.frm_pedido_resurtir.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_saldos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'   Call pro_ordena_listas(lv_saldos, ColumnHeader)
End Sub

Private Sub lv_saldos_GotFocus()
   frm_periodo.Visible = False
End Sub

Private Sub lv_saldos_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = 117 Then
'      If Me.lv_saldos.ListItems.Count > 0 Then
'         Me.txt_orden_surtido = Me.lv_saldos.selectedItem.SubItems(3)
'         If IsNumeric(Me.txt_orden_surtido) Then
'            cnn.BeginTrans
'            rs.Open "select max(INTE_TEM_CONSECUTIVO) as consecutivo from tb_temp_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
'            If Not rs.EOF Then
'               var_consecutivo = IIf(IsNull(rs!consecutivo), 0, rs!consecutivo) + 1
'            Else
'               var_consecutivo = 1
'            End If
'            rs.Close
'            rs.Open "insert into tb_temp_orden_surtido (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
'            cnn.CommitTrans
'            Cadena = "INSERT INTO [vianney].[dbo].[TB_TEMP_ORDEN_SURTIDO] ([INTE_TEM_CONSECUTIVO], [CHAR_TPE_TIPO_PEDIDO], [INTE_PED_NUMERO], [VCHA_ALM_ALMACEN_ID], [INTE_ORS_ORDEN_SURTIDO], [DTIM_ORS_FECHA_CARGA], [DTIM_ORS_FECHA_CADUCA], [VCHA_ESB_NOMBRE], [VCHA_ESB_ESTABLECIMIENTO_ID], [VCHA_CLI_CLAVE_ID], [VCHA_CLI_NOMBRE], "
'            Cadena = Cadena + "[CHAR_PRI_PRIORIDAD_ID], [FLOA_ORS_DESCUENTO_1], [FLOA_ORS_DESCUENTO_2], [VCHA_ART_ARTICULO_ID], [VCHA_ART_NOMBRE_ESPAÑOL], [FLOA_ORS_PRECIO], [FLOA_ORS_CANTIDAD_SURTIR], [VCHA_AGE_NOMBRE], [VCHA_RUT_NOMBRE], [INTE_PED_DIAS_CONDICIONES], [INTE_PED_DIAS_CADUCIDAD], [MONE_ART_COSTO_ESTANDAR], [FLOA_ORS_CANTIDAD_SURTIDA], "
'            Cadena = Cadena + "[VCHA_TIT_TITULAR_ID], [VCHA_TIT_NOMBRE], [VCHA_RUT_RUTA_ID], [VCHA_AGE_AGENTE_ID], [CHAR_ORS_ESTATUS], [VCHA_MOV_MOVIMIENTO_ID], [VCHA_CLI_EMAIL], [FLOA_TPE_IVA], [ALMACEN_AGENTE], [INTE_TCL_EMPACADO], [VCHA_TCL_ALMACEN_EMPAQUE], [FLOA_ORS_CANTIDAD_EMPACADA], [VCHA_PLA_PAZO_ID], [INTE_CLI_AGRUPADOR], [VCHA_MON_MONEDA_ID],"
'            Cadena = Cadena + "[VCHA_ALM_NOMBRE], [FLOA_ORS_PROMOCION_1], [FLOA_ORS_PROMOCION_2], [INTE_ORS_FACTURA_CEROS], [CHAR_PED_TIPO], [VCHA_UOR_UNIDAD_ID], [INTE_PED_REFERENCIA],  [VCHA_EXI_UBICACION], [FLOA_ORS_CANTIDAD_NEGADA], [VCHA_TEM_CODIGO_BARRAS]) "
'            Cadena = Cadena + " select " + CStr(var_consecutivo) + ", CHAR_TPE_TIPO_PEDIDO_ID, INTE_PED_NUMERO, VCHA_ALM_ALMACEN_ID, INTE_ORS_ORDEN_SURTIDO, DTIM_ORS_FECHA_CARGA, DTIM_ORS_FECHA_CADUCA, VCHA_ESB_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, CHAR_PRI_PRIORIDAD_ID, FLOA_ORS_DESCUENTO_1, FLOA_ORS_DESCUENTO_2, VCHA_ART_ARTICULO_ID, "
'            Cadena = Cadena + " VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ORS_PRECIO, FLOA_ORS_CANTIDAD_SURTIR, VCHA_AGE_NOMBRE, VCHA_RUT_NOMBRE, INTE_PED_DIAS_CONDICIONES, INTE_PED_DIAS_CADUCIDAD, MONE_ART_COSTO_ESTANDAR, FLOA_ORS_CANTIDAD_SURTIDA, VCHA_TIT_TITULAR_ID, VCHA_TIT_NOMBRE, VCHA_RUT_RUTA_ID, VCHA_AGE_AGENTE_ID, CHAR_ORS_ESTATUS, VCHA_MOV_MOVIMIENTO_ID, VCHA_CLI_EMAIL, FLOA_TPE_IVA,"
'            Cadena = Cadena + " ALMACEN_AGENTE, INTE_TCL_EMPACADO, VCHA_TCL_ALMACEN_EMPAQUE, FLOA_ORS_CANTIDAD_EMPACADA, VCHA_PLA_PlAZO_ID, INTE_CLI_AGRUPADOR, VCHA_MON_MONEDA_ID, VCHA_ALM_NOMBRE, FLOA_ORS_PROMOCION_1, FLOA_ORS_PROMOCION_2, INTE_ORS_FACTURA_CEROS, CHAR_PED_TIPO, VCHA_UOR_UNIDAD_ID, INTE_PED_REFERENCIA, VCHA_EXI_UBICACION, FLOA_ORS_CANTIDAD_NEGADA, " + Format(Me.txt_orden_surtido, "##########") + " from vw_orden_surtido where inte_ors_orden_surtido = " + CStr(CDbl(Me.txt_orden_surtido))
'            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
'            If var_empresa = "18" Then
'               Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido_nueva_textilera.rpt")
'            Else
'               Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido_nueva_FT.rpt")
'            End If
'            reporte.RecordSelectionFormula = "{TB_TEMP_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO} = " + CStr(CDbl(Me.txt_orden_surtido)) + " AND {TB_TEMP_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {TB_TEMP_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR} > 0"
'            frmvistasprevias.cr.ReportSource = reporte
'            For ntablas = 1 To reporte.Database.Tables.Count
'                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
'            Next ntablas
'            frmvistasprevias.cr.ViewReport
'            frmvistasprevias.Caption = "Orden de Surtido"
'            frmvistasprevias.Show 1
'            Set reporte = Nothing
'            rs.Open "delete from tb_temp_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
'         Else
'            MsgBox "Número de orden de surtido incorrecta", vbOKOnly, "ATENCION"
'         End If
'      End If
'   End If
End Sub

Private Sub mes_DateDblClick(ByVal DateDblClicked As Date)
   If var_tipo_fecha = 1 Then
      Me.txt_inicio = Me.mes.Value
      Me.txt_inicio.SetFocus
   End If
   If var_tipo_fecha = 2 Then
      Me.txt_fin = Me.mes.Value
      Me.txt_fin.SetFocus
   End If
End Sub

Private Sub mes_LostFocus()
   mes.Visible = False
End Sub

Private Sub Timer1_Timer()
'   If var_intervalo = 6 Then
'      var_intervalo = 0
'      Call Command1_Click
'   Else
'      var_intervalo = var_intervalo + 1
'   End If
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_tipo_fecha = 2
      If IsDate(Me.txt_fin) Then
         Me.mes = CDate(Me.txt_fin)
      Else
         mes = Date
      End If
      mes.Visible = True
      mes.SetFocus
   End If
End Sub

Private Sub txt_fin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_aceptar_pedidos.SetFocus
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_tipo_fecha = 1
      If IsDate(Me.txt_inicio) Then
         Me.mes = CDate(Me.txt_inicio)
      Else
         Me.mes = Date
      End If
      mes.Visible = True
      mes.SetFocus
   End If
End Sub

Private Sub txt_inicio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_fin.SetFocus
   End If
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
         Cadena = "INSERT INTO [vianney].[dbo].[TB_TEMP_ORDEN_SURTIDO] ([INTE_TEM_CONSECUTIVO], [CHAR_TPE_TIPO_PEDIDO], [INTE_PED_NUMERO], [VCHA_ALM_ALMACEN_ID], [INTE_ORS_ORDEN_SURTIDO], [DTIM_ORS_FECHA_CARGA], [DTIM_ORS_FECHA_CADUCA], [VCHA_ESB_NOMBRE], [VCHA_ESB_ESTABLECIMIENTO_ID], [VCHA_CLI_CLAVE_ID], [VCHA_CLI_NOMBRE], "
         Cadena = Cadena + "[CHAR_PRI_PRIORIDAD_ID], [FLOA_ORS_DESCUENTO_1], [FLOA_ORS_DESCUENTO_2], [VCHA_ART_ARTICULO_ID], [VCHA_ART_NOMBRE_ESPAÑOL], [FLOA_ORS_PRECIO], [FLOA_ORS_CANTIDAD_SURTIR], [VCHA_AGE_NOMBRE], [VCHA_RUT_NOMBRE], [INTE_PED_DIAS_CONDICIONES], [INTE_PED_DIAS_CADUCIDAD], [MONE_ART_COSTO_ESTANDAR], [FLOA_ORS_CANTIDAD_SURTIDA], "
         Cadena = Cadena + "[VCHA_TIT_TITULAR_ID], [VCHA_TIT_NOMBRE], [VCHA_RUT_RUTA_ID], [VCHA_AGE_AGENTE_ID], [CHAR_ORS_ESTATUS], [VCHA_MOV_MOVIMIENTO_ID], [VCHA_CLI_EMAIL], [FLOA_TPE_IVA], [ALMACEN_AGENTE], [INTE_TCL_EMPACADO], [VCHA_TCL_ALMACEN_EMPAQUE], [FLOA_ORS_CANTIDAD_EMPACADA], [VCHA_PLA_PAZO_ID], [INTE_CLI_AGRUPADOR], [VCHA_MON_MONEDA_ID],"
         Cadena = Cadena + "[VCHA_ALM_NOMBRE], [FLOA_ORS_PROMOCION_1], [FLOA_ORS_PROMOCION_2], [INTE_ORS_FACTURA_CEROS], [CHAR_PED_TIPO], [VCHA_UOR_UNIDAD_ID], [INTE_PED_REFERENCIA],  [VCHA_EXI_UBICACION], [FLOA_ORS_CANTIDAD_NEGADA], [VCHA_TEM_CODIGO_BARRAS]) "
         Cadena = Cadena + " select " + CStr(var_consecutivo) + ", CHAR_TPE_TIPO_PEDIDO_ID, INTE_PED_NUMERO, VCHA_ALM_ALMACEN_ID, INTE_ORS_ORDEN_SURTIDO, DTIM_ORS_FECHA_CARGA, DTIM_ORS_FECHA_CADUCA, VCHA_ESB_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, CHAR_PRI_PRIORIDAD_ID, FLOA_ORS_DESCUENTO_1, FLOA_ORS_DESCUENTO_2, VCHA_ART_ARTICULO_ID, "
         Cadena = Cadena + " VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ORS_PRECIO, FLOA_ORS_CANTIDAD_SURTIR, VCHA_AGE_NOMBRE, VCHA_RUT_NOMBRE, INTE_PED_DIAS_CONDICIONES, INTE_PED_DIAS_CADUCIDAD, MONE_ART_COSTO_ESTANDAR, FLOA_ORS_CANTIDAD_SURTIDA, VCHA_TIT_TITULAR_ID, VCHA_TIT_NOMBRE, VCHA_RUT_RUTA_ID, VCHA_AGE_AGENTE_ID, CHAR_ORS_ESTATUS, VCHA_MOV_MOVIMIENTO_ID, VCHA_CLI_EMAIL, FLOA_TPE_IVA,"
         Cadena = Cadena + " ALMACEN_AGENTE, INTE_TCL_EMPACADO, VCHA_TCL_ALMACEN_EMPAQUE, FLOA_ORS_CANTIDAD_EMPACADA, VCHA_PLA_PlAZO_ID, INTE_CLI_AGRUPADOR, VCHA_MON_MONEDA_ID, VCHA_ALM_NOMBRE, FLOA_ORS_PROMOCION_1, FLOA_ORS_PROMOCION_2, INTE_ORS_FACTURA_CEROS, CHAR_PED_TIPO, VCHA_UOR_UNIDAD_ID, INTE_PED_REFERENCIA, VCHA_EXI_UBICACION, FLOA_ORS_CANTIDAD_NEGADA, " + Format(Me.txt_orden_surtido, "##########") + " from vw_orden_surtido where inte_ors_orden_surtido = " + CStr(Me.txt_orden_surtido)
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

