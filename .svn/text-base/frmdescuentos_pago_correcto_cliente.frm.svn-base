VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcausas_no_otorgamiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Causas de no otorgamiento del descuento por pago correcto y puntual"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   Begin VB.Frame Frame2 
      Caption         =   " Agente "
      Height          =   675
      Left            =   4905
      TabIndex        =   16
      Top             =   510
      Width           =   6660
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   75
         TabIndex        =   7
         Top             =   225
         Width           =   1380
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   315
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   5100
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clientes"
      Height          =   2370
      Left            =   105
      TabIndex        =   14
      Top             =   1260
      Width           =   11445
      Begin MSComctlLib.ListView lv_clientes 
         Height          =   2010
         Left            =   60
         TabIndex        =   9
         Top             =   180
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   3545
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave Grupo Actual"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clave Grupo Real"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clave Titualar"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Clave Cliente"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nombre"
            Object.Width           =   9349
         EndProperty
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmdescuentos_pago_correcto_cliente.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11295
      Picture         =   "frmdescuentos_pago_correcto_cliente.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   "  Detalle del Calculo "
      Height          =   3570
      Left            =   105
      TabIndex        =   13
      Top             =   3690
      Width           =   11475
      Begin MSComctlLib.ListView lv_detalle 
         Height          =   3255
         Left            =   60
         TabIndex        =   10
         Top             =   195
         Width           =   11340
         _ExtentX        =   20003
         _ExtentY        =   5741
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
            Text            =   "Tipo"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Factura"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Fecha Factura  "
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Fecha Pago"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Importe Factura"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Importe Pago "
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "%DF"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "%BF"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "%SF"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "% Final"
            Object.Width           =   1676
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Dias "
            Object.Width           =   1411
         EndProperty
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   780
      Picture         =   "frmdescuentos_pago_correcto_cliente.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_aplicar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   450
      Picture         =   "frmdescuentos_pago_correcto_cliente.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Aplicar Pagos Alt + A"
      Top             =   15
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame frm_periodo 
      Caption         =   " Periodo "
      Height          =   675
      Left            =   90
      TabIndex        =   0
      Top             =   510
      Width           =   4770
      Begin VB.ComboBox cmb_meses 
         Height          =   315
         ItemData        =   "frmdescuentos_pago_correcto_cliente.frx":0A18
         Left            =   630
         List            =   "frmdescuentos_pago_correcto_cliente.frx":0A40
         TabIndex        =   5
         Top             =   240
         Width           =   2280
      End
      Begin VB.ListBox lst_años 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmdescuentos_pago_correcto_cliente.frx":0AA9
         Left            =   3630
         List            =   "frmdescuentos_pago_correcto_cliente.frx":0AEC
         TabIndex        =   6
         Top             =   255
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   300
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Left            =   3240
         TabIndex        =   11
         Top             =   300
         Width           =   330
      End
   End
   Begin VB.Frame Frame3 
      Height          =   90
      Left            =   75
      TabIndex        =   15
      Top             =   315
      Width           =   11565
   End
End
Attribute VB_Name = "frmcausas_no_otorgamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim fecha_inicio As Date
Dim fecha_fin As Date
Private Sub cmb_meses_Click()
   Dim fecha_anterior As Date
   Dim dia_anterior As Integer
   Dim mes_anterior As Integer
   Dim año_anterior As Integer
   Dim dia As Integer
   Dim mes As Integer
   Dim año As Integer
   Dim periodo As String
   
   If cmb_meses = "Enero" Then
      mes_anterior = 1
   End If
   If cmb_meses = "Febrero" Then
      mes_anterior = 2
   End If
   If cmb_meses = "Marzo" Then
      mes_anterior = 3
   End If
   If cmb_meses = "Abril" Then
      mes_anterior = 4
   End If
   If cmb_meses = "Mayo" Then
      mes_anterior = 5
   End If
   If cmb_meses = "Junio" Then
      mes_anterior = 6
   End If
   If cmb_meses = "Julio" Then
      mes_anterior = 7
   End If
   If cmb_meses = "Agosto" Then
      mes_anterior = 8
   End If
   If cmb_meses = "Septiembre" Then
      mes_anterior = 9
   End If
   If cmb_meses = "Octubre" Then
      mes_anterior = 10
   End If
   If cmb_meses = "Noviembre" Then
      mes_anterior = 11
   End If
   If cmb_meses = "Diciembre" Then
      mes_anterior = 12
   End If
   año_anterior = lst_años
   If mes_anterior = 1 Or mes_anterior = 3 Or mes_anterior = 5 Or mes_anterior = 7 Or mes_anterior = 8 Or mes_anterior = 10 Or mes_anterior = 12 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      fecha_fin = CDate("31/" + Str(mes_anterior) + "/" + Str(año_anterior))
      dia_anterior = 31
   End If
   If mes_anterior = 4 Or mes_anterior = 6 Or mes_anterior = 9 Or mes_anterior = 11 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      fecha_fin = CDate("30/" + Str(mes_anterior) + "/" + Str(año_anterior))
      dia_anterior = 30
   End If
   
   If mes_anterior = 2 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      If año_anterior = 2004 Or año_anterior = 2008 Or año_anterior = 2012 Or año_anterior = 2016 Or año_anterior = 2020 Or año_anterior = 2024 Then
         fecha_fin = CDate("29/" + Str(mes_anterior) + "/" + Str(año_anterior))
         dia_anterior = 29
      Else
         fecha_fin = CDate("28/" + Str(mes_anterior) + "/" + Str(año_anterior))
         dia_anterior = 28
      End If
   End If
End Sub

Private Sub cmb_meses_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      lst_años.SetFocus
   End If
End Sub

Private Sub cmd_imprimir_Click()
   rs.Open "SELECT * FROM [VW_DESCUENTOS_2%] Where dtim_dpc_periodo_inicio = '" + CStr(fecha_inicio) + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Set reporte = appl.OpenReport(App.Path + "\rep_descuento_pago_correcto.rpt")
      reporte.RecordSelectionFormula = "{VW_DESCUENTOS_2%.DTIM_DPC_PERIODO_INICIO} = date('" + CStr(fecha_inicio) + "')"
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte de descuentos por pago correcto y puntual"
      frmvistasprevias.Show
      Set reporte = Nothing
   Else
      MsgBox "No existe información para el periodo seleccionado", vbOKOnly, "ATENCION"
   End If
   rs.Close
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
   var_cadena_seguridad = ""
   Dim mes As Integer
   Dim año As Integer
   mes = Month(Date)
   año = Year(Date)
   If mes = 1 Then
      cmb_meses = "Enero"
   End If
   If mes = 2 Then
      cmb_meses = "Febrero"
   End If
   If mes = 3 Then
      cmb_meses = "Marzo"
   End If
   If mes = 4 Then
      cmb_meses = "Abril"
   End If
   If mes = 5 Then
      cmb_meses = "Mayo"
   End If
   If mes = 6 Then
      cmb_meses = "Junio"
   End If
   If mes = 7 Then
      cmb_meses = "Julio"
   End If
   If mes = 8 Then
      cmb_meses = "Agosto"
   End If
   If mes = 9 Then
      cmb_meses = "Septiembre"
   End If
   If mes = 10 Then
      cmb_meses = "Octubre"
   End If
   If mes = 11 Then
      cmb_meses = "Noviembre"
   End If
   If mes = 12 Then
      cmb_meses = "Diciembre"
   End If
   lst_años = año
   Top = 0
   Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro_causas_devolucion = False
   Call activa_forma(var_activa_forma_causas_no_otorgamiento)
End Sub

Private Sub lst_años_Click()
   Dim fecha_anterior As Date
   Dim dia_anterior As Integer
   Dim mes_anterior As Integer
   Dim año_anterior As Integer
   Dim dia As Integer
   Dim mes As Integer
   Dim año As Integer
   Dim periodo As String
   
   If cmb_meses = "Enero" Then
      mes_anterior = 1
   End If
   If cmb_meses = "Febrero" Then
      mes_anterior = 2
   End If
   If cmb_meses = "Marzo" Then
      mes_anterior = 3
   End If
   If cmb_meses = "Abril" Then
      mes_anterior = 4
   End If
   If cmb_meses = "Mayo" Then
      mes_anterior = 5
   End If
   If cmb_meses = "Junio" Then
      mes_anterior = 6
   End If
   If cmb_meses = "Julio" Then
      mes_anterior = 7
   End If
   If cmb_meses = "Agosto" Then
      mes_anterior = 8
   End If
   If cmb_meses = "Septiembre" Then
      mes_anterior = 9
   End If
   If cmb_meses = "Octubre" Then
      mes_anterior = 10
   End If
   If cmb_meses = "Noviembre" Then
      mes_anterior = 11
   End If
   If cmb_meses = "Diciembre" Then
      mes_anterior = 12
   End If
   año_anterior = lst_años
   If mes_anterior = 1 Or mes_anterior = 3 Or mes_anterior = 5 Or mes_anterior = 7 Or mes_anterior = 8 Or mes_anterior = 10 Or mes_anterior = 12 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      fecha_fin = CDate("31/" + Str(mes_anterior) + "/" + Str(año_anterior))
      dia_anterior = 31
   End If
   If mes_anterior = 4 Or mes_anterior = 6 Or mes_anterior = 9 Or mes_anterior = 11 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      fecha_fin = CDate("30/" + Str(mes_anterior) + "/" + Str(año_anterior))
      dia_anterior = 30
   End If
   
   If mes_anterior = 2 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      If año_anterior = 2004 Or año_anterior = 2008 Or año_anterior = 2012 Or año_anterior = 2016 Or año_anterior = 2020 Or año_anterior = 2024 Then
         fecha_fin = CDate("29/" + Str(mes_anterior) + "/" + Str(año_anterior))
         dia_anterior = 29
      Else
         fecha_fin = CDate("28/" + Str(mes_anterior) + "/" + Str(año_anterior))
         dia_anterior = 28
      End If
   End If
End Sub

Private Sub lst_años_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_agente.SetFocus
   End If
End Sub

Private Sub lv_clientes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_clientes, ColumnHeader)
End Sub

Private Sub lv_clientes_ItemClick(ByVal Item As MSComctlLib.ListItem)
   lv_detalle.ListItems.Clear
End Sub

Private Sub lv_clientes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "select * from TB_DESCUENTOS_PAGO_CORRECTO_ASIGNADO where vcha_age_agente_id = '" + txt_agente + "' and vcha_cli_clave_id = '" + lv_clientes.selectedItem.SubItems(3) + "' and inte_dpc_puntos >= 0 and dtim_dpc_periodo_inicio = '" + CStr(fecha_inicio) + "' and dtim_dpc_periodo_fin = '" + CStr(fecha_fin + 1) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Dim list_item As ListItem
         Dim numero_items_detalle As Double
         numero_items_detalle = 0
         lv_detalle.ListItems.Clear
            While Not rs.EOF
               Set list_item = lv_detalle.ListItems.Add(, , rs!vcha_dpc_tipo)
               list_item.SubItems(1) = IIf(IsNull(rs!inte_Car_numero), "", rs!inte_Car_numero)
               list_item.SubItems(2) = IIf(IsNull(rs!dtim_Car_fecha), "", Format(rs!dtim_Car_fecha, "Short Date"))
               list_item.SubItems(3) = IIf(IsNull(rs!DTIM_RCO_FECHA_APLICACION), "", Format(rs!DTIM_RCO_FECHA_APLICACION, "Short Date"))
               list_item.SubItems(4) = Format(IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto), "###,##0.00")
               list_item.SubItems(5) = Format(IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe), "###,##0.00")
               list_item.SubItems(6) = Format(IIf(IsNull(rs!FLOA_DPC_PORCENTAJE_DESCUENTO_FINANCIERO), 0, rs!FLOA_DPC_PORCENTAJE_DESCUENTO_FINANCIERO), "###,##0.00")
               list_item.SubItems(7) = Format(IIf(IsNull(rs!FLOA_DPC_PORCENTAJE_BONIFICACION_FINANCIERA), 0, rs!FLOA_DPC_PORCENTAJE_BONIFICACION_FINANCIERA), "###,##0.00")
               list_item.SubItems(8) = Format(IIf(IsNull(rs!FLOA_DPC_SALDO_FINANCIERO), 0, rs!FLOA_DPC_SALDO_FINANCIERO), "###,##0.00")
               list_item.SubItems(9) = Format(IIf(IsNull(rs!FLOA_DPC_DESCUENTO_PONDERADO), 0, rs!FLOA_DPC_DESCUENTO_PONDERADO), "###,##0.00")
               list_item.SubItems(10) = IIf(IsNull(rs!INTE_DPC_PUNTOS), 0, rs!INTE_DPC_PUNTOS)
               rs.MoveNext
            Wend
      End If
      rs.Close
   End If
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_nombre_agente.SetFocus
   End If
End Sub

Private Sub txt_agente_LostFocus()
   If Trim(txt_agente) <> "" Then
      Dim contador As Integer
      Dim numero_items_clientes As Double
      txt_grupo = ""
      txt_nombre_grupo = ""
      txt_cobranza = ""
      txt_descuento_aplicado = ""
      txt_descuento_aplicar = ""
      txt_causa = ""
      rs.Open "select * from tb_Agentes where vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
         Dim list_item As ListItem
         Dim var_grupo_actual As String
         numero_items_clientes = 0
         lv_clientes.ListItems.Clear
         lv_detalle.ListItems.Clear
         rsaux2.Open "select distinct vcha_gac_grupo_actual_id, vcha_gre_grupo_real_id, vcha_tit_titular_id, vcha_cli_clave_id, vcha_cli_nombre from vw_clientes where  vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux2.EOF Then
            While Not rsaux2.EOF
               var_grupo_actual = IIf(IsNull(rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID)
               Set list_item = lv_clientes.ListItems.Add(, , var_grupo_actual)
               list_item.SubItems(1) = IIf(IsNull(rsaux2!vcha_gre_grupo_real_id), "", rsaux2!vcha_gre_grupo_real_id)
               list_item.SubItems(2) = IIf(IsNull(rsaux2!vcha_tit_titular_id), "", rsaux2!vcha_tit_titular_id)
               list_item.SubItems(3) = IIf(IsNull(rsaux2!vcha_cli_clave_id), "", rsaux2!vcha_cli_clave_id)
               list_item.SubItems(4) = IIf(IsNull(rsaux2!VCHA_CLI_NOMBRE), "", rsaux2!VCHA_CLI_NOMBRE)
               numero_items_clientes = numero_items_clientes + 1
               rsaux2.MoveNext
            Wend
            If numero_items_clientes > 7 Then
               lv_clientes.ColumnHeaders(5).Width = 5090.22
            Else
               lv_clientes.ColumnHeaders(5).Width = 5300.22
            End If
         Else
            MsgBox "El agente no tiene "
         End If
         rsaux2.Close
      End If
      rs.Close
   End If
End Sub
