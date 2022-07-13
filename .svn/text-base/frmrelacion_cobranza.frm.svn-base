VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmrelacion_cobranza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relación de Cobranza"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm_lista 
      Height          =   2430
      Left            =   3270
      TabIndex        =   31
      Top             =   2055
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1920
         Left            =   30
         TabIndex        =   32
         Top             =   450
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3387
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
         TabIndex        =   33
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_cambios_relacion 
      Height          =   2415
      Left            =   3150
      TabIndex        =   16
      Top             =   2070
      Width           =   5925
      Begin VB.TextBox txt_serie 
         Height          =   360
         Left            =   3060
         TabIndex        =   29
         Top             =   1965
         Width           =   510
      End
      Begin VB.TextBox txt_numero 
         Height          =   345
         Left            =   1095
         TabIndex        =   28
         Top             =   1973
         Width           =   1335
      End
      Begin VB.TextBox txt_documento 
         Height          =   345
         Left            =   1095
         TabIndex        =   27
         Top             =   1590
         Width           =   510
      End
      Begin VB.TextBox txt_cheque 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1095
         TabIndex        =   26
         Top             =   1215
         Width           =   1335
      End
      Begin VB.TextBox txt_nombre_cliente 
         Enabled         =   0   'False
         Height          =   345
         Left            =   2445
         TabIndex        =   25
         Top             =   840
         Width           =   3345
      End
      Begin VB.TextBox txt_clave_cliente 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1095
         TabIndex        =   24
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmd_cancelar_cambios 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   420
         Picture         =   "frmrelacion_cobranza.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Cancelar Cambios"
         Top             =   420
         Width           =   330
      End
      Begin VB.CommandButton cmd_aceptar_cambios 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Picture         =   "frmrelacion_cobranza.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Aplicar Pagos Alt + A"
         Top             =   420
         Width           =   330
      End
      Begin VB.Frame Frame4 
         Height          =   30
         Left            =   0
         TabIndex        =   19
         Top             =   750
         Width           =   5925
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   2610
         TabIndex        =   34
         Top             =   2048
         Width           =   405
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   150
         TabIndex        =   30
         Top             =   2048
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   1665
         Width           =   870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cheque:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   1290
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   915
         Width           =   525
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "Correcciones de relación de cobranza"
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   17
         Top             =   120
         Width           =   5850
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   450
      Picture         =   "frmrelacion_cobranza.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Nota de Crédito"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_aplicar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmrelacion_cobranza.frx":0396
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Aplicar Pagos Alt + A"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11190
      Picture         =   "frmrelacion_cobranza.frx":04E0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin MSComctlLib.ListView lv_detalle 
      Height          =   5175
      Left            =   210
      TabIndex        =   1
      Top             =   1920
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   9128
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
      NumItems        =   36
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cliente"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipo"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Número"
         Object.Width           =   1640
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fecha"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Importe     "
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Moneda"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Saldo      "
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Desc."
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Cheque"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Fecha"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Importe     "
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Desc."
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Aplicado"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Aplicar"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Grupo Actual"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Grupo Real"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Titular"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Establecimiento"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Moneda"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "IVA"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Impuesto 2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Impuesto 3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Descuento Aplicar"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Nota Crédito"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Serie"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Banco"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "Fecha cheque"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "partida"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "Fecha relacion"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "cheque deposito"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "Banco Cheque"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "Deposito"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "Banco Deposito"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "Fecha Deposito"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "Num. Dep."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "Fecha Insercion"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   90
      Left            =   90
      TabIndex        =   15
      Top             =   300
      Width           =   11490
   End
   Begin VB.Frame Frame2 
      Caption         =   " Relación de Cobranza "
      Height          =   5505
      Left            =   135
      TabIndex        =   14
      Top             =   1665
      Width           =   11430
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos Generales "
      Height          =   1140
      Left            =   135
      TabIndex        =   0
      Top             =   465
      Width           =   11430
      Begin VB.TextBox txt_fecha 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1365
         TabIndex        =   8
         Top             =   720
         Width           =   1620
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   4785
         TabIndex        =   9
         Top             =   720
         Width           =   1845
      End
      Begin VB.TextBox txt_nombre_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6120
         TabIndex        =   7
         Top             =   375
         Width           =   4530
      End
      Begin VB.TextBox txt_clave_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4785
         TabIndex        =   6
         Top             =   375
         Width           =   1320
      End
      Begin VB.TextBox txt_folio 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1365
         TabIndex        =   5
         Top             =   255
         Width           =   1620
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Relación:"
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   780
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe Relación:"
         Height          =   195
         Left            =   3480
         TabIndex        =   12
         Top             =   780
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   3495
         TabIndex        =   11
         Top             =   435
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   375
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmrelacion_cobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tabla As ADODB.Connection
Dim var_ruta As String
Dim var_serie As String


Private Sub cmd_aceptar_cambios_Click()
   Me.frm_cambios_relacion.Visible = False
   If Trim(txt_clave_cliente) <> "" Then
      If Trim(txt_documento) <> "" Then
         If Trim(txt_numero) <> "" Then
            rs.Open "SELECT * FROM TB_ENCABEZADO_CARTERA WHERE VCHA_CLI_CLAVE_ID = '" + txt_clave_cliente + "' AND INTE_CAR_NUMERO = " + txt_numero + " AND vcha_CAR_DOCUMENTO  = '" + Trim(txt_documento) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_si = MsgBox("¿Desea hacer los cambios indicados?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  var_si = MsgBox("Confirmar los cambios indicados?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     cnn.BeginTrans
                     '3
                     rsaux.Open "DELETE FROM TB_RELACION_COBRANZA WHERE VCHA_RCO_FOLIO = '" + Me.txt_folio + "' AND VCHA_CLI_CLAVE_ID = '" + lv_detalle.selectedItem + "' AND vcha_car_documento = '" + lv_detalle.selectedItem.SubItems(1) + "' and inte_Car_numero = " + lv_detalle.selectedItem.SubItems(2) + " and vcha_rco_cheque = '" + lv_detalle.selectedItem.SubItems(8) + "' and vcha_ban_banco_id = '" + lv_detalle.selectedItem.SubItems(25) + "' and inte_rco_partida = " + lv_detalle.selectedItem.SubItems(27), cnn, adOpenDynamic, adLockOptimistic
                     'rsaux.Open "DELETE FROM TB_cheques WHERE VCHA_RCO_FOLIO = '" + Me.txt_folio + "' AND VCHA_CLI_CLAVE_ID = '" + lv_detalle.selectedItem + "' and vcha_che_cheque = '" + lv_detalle.selectedItem.SubItems(8) + "' and vcha_ban_banco_id = '" + lv_detalle.selectedItem.SubItems(25) + "'", cnn, adOpenDynamic, adLockOptimistic
                     'rsaux.Open "update tb_cheques set floa_che_importe = floa_che_importe - " + CStr(CDbl(lv_detalle.selectedItem.SubItems(10))) + " WHERE VCHA_RCO_FOLIO = '" + Me.txt_folio + "' AND VCHA_CLI_CLAVE_ID = '" + lv_detalle.selectedItem + "' and vcha_che_cheque = '" + lv_detalle.selectedItem.SubItems(8) + "' and vcha_ban_banco_id = '" + lv_detalle.selectedItem.SubItems(25) + "'", cnn, adOpenDynamic, adLockOptimistic
                     '          "EXECUTE RELACION_COBRANZA_I    ?var_empresa, ?var_unidad,  ?var_folio,               ?var_fecha,                                     ?var_agent2,                  ?var_cliente,               ?var_cheque,              ?var_fecha_cheque,                             ?var_importe,                                  ?var_descuento,                           ?var_factura,?var_Cero, ?var_cero, ?var_part,                      ?var_cero,    ?serie,                                     ?var_tipo,                        ?var_banco")
                     Cadena = "EXECUTE RELACION_COBRANZA_I '" + var_empresa + "', '', '" + Me.txt_folio + "', '" + lv_detalle.selectedItem.SubItems(28) + "', '" + Me.txt_clave_agente + "', '" + txt_clave_cliente + "', '" + Me.txt_cheque + "', '" + lv_detalle.selectedItem.SubItems(26) + "', " + CStr(CDbl(lv_detalle.selectedItem.SubItems(10))) + ", " + lv_detalle.selectedItem.SubItems(11) + ", " + txt_numero + ", 0, 0, " + lv_detalle.selectedItem.SubItems(27) + ", 0, '" + txt_serie + "', '" + txt_documento + "', '" + Trim(lv_detalle.selectedItem.SubItems(25)) + "'"
                     rsaux.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     cnn.CommitTrans
                     Dim var_fecha_cheques As String
                     var_dia = CStr(Day(Date))
                     var_mes = CStr(Month(Date))
                     var_año = CStr(Year(Date))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     var_fecha_cheques = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     If rs.State = 1 Then
                        rs.Close
                     End If
                     rs.Open "select * from tb_relacion_cobranza with (nolock) where vcha_rco_folio = '" + txt_folio + "' and dtim_rco_fecha_cheque <= " + var_fecha_cheques, cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_importe = 0
                        txt_clave_agente = rs!VCHA_AGE_AGENTE_ID
                        rsaux2.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                        txt_nombre_agente = rsaux2!VCHA_AGE_NOMBRE
                        rsaux2.Close
                        txt_fecha = rs!dtim_rco_fecha_relacion
                        lv_detalle.ListItems.Clear
                        While Not rs.EOF
                           Set list_item = lv_detalle.ListItems.Add(, , rs!vcha_Cli_clave_id)
                           list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
                           list_item.SubItems(2) = IIf(IsNull(rs!inte_car_numero), "", rs!inte_car_numero)
                           list_item.SubItems(3) = IIf(IsNull(rs!dtim_car_fecha), "", Format(rs!dtim_car_fecha, "Short Date"))
                           list_item.SubItems(4) = IIf(IsNull(rs!floa_Car_importe), 0, Round(rs!floa_Car_importe, 2))
                           rsaux2.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux2.EOF Then
                              list_item.SubItems(5) = IIf(IsNull(rsaux2!vcha_mon_nombre_plural), "", Trim(rsaux2!vcha_mon_nombre_plural))
                           End If
                           rsaux2.Close
                           rsaux2.Open "select floa_sal_importe from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_car_numero = " + Str(rs!inte_car_numero) + " and vcha_ser_Serie_id = '" + rs!vcha_ser_Serie_id + "' and vcha_cli_clave_id = '" + rs!vcha_Cli_clave_id + "' AND VCHA_cAR_DOCUMENTO = '" + rs!vcha_Car_documento + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux2.EOF Then
                              list_item.SubItems(6) = Format(IIf(IsNull(rsaux2!FLOA_sAL_IMPORTE), 0, rsaux2!FLOA_sAL_IMPORTE), "###,##0.00")
                           Else
                              list_item.SubItems(6) = 0
                           End If
                           rsaux2.Close
                           list_item.SubItems(7) = IIf(IsNull(rs!floa_Car_descuento_aplicado), 0, rs!floa_Car_descuento_aplicado)
                           list_item.SubItems(8) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
                           list_item.SubItems(9) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
                           list_item.SubItems(10) = Format(IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe), "###,##0.00")
                           list_item.SubItems(11) = IIf(IsNull(rs!FLOA_RCO_DESCUENTO_OTORGADO), 0, rs!FLOA_RCO_DESCUENTO_OTORGADO)
                           list_item.SubItems(12) = IIf(IsNull(rs!char_rco_aplicada), "", rs!char_rco_aplicada)
                           list_item.SubItems(14) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
                           list_item.SubItems(15) = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
                           list_item.SubItems(16) = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
                           list_item.SubItems(17) = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
                           list_item.SubItems(18) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                           list_item.SubItems(19) = IIf(IsNull(rs!floa_rco_iva), 0, rs!floa_rco_iva)
                           list_item.SubItems(20) = IIf(IsNull(rs!floa_rco_impuesto_2), 0, rs!floa_rco_impuesto_2)
                           list_item.SubItems(21) = IIf(IsNull(rs!floa_rco_impuesto_3), 0, rs!floa_rco_impuesto_3)
                           list_item.SubItems(22) = IIf(IsNull(rs!floa_Rco_descuento_aplicar), 0, rs!floa_Rco_descuento_aplicar)
                           list_item.SubItems(23) = IIf(IsNull(rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO), 0, rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO)
                           list_item.SubItems(24) = IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id)
                           list_item.SubItems(25) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
                           list_item.SubItems(26) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
                           list_item.SubItems(27) = IIf(IsNull(rs!inte_rco_partida), "", rs!inte_rco_partida)
                           list_item.SubItems(28) = IIf(IsNull(rs!dtim_rco_fecha_relacion), "", rs!dtim_rco_fecha_relacion)
                           list_item.SubItems(29) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
                           list_item.SubItems(30) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
                           list_item.SubItems(31) = IIf(IsNull(rs!vcha_rco_deposito), "", rs!vcha_rco_deposito)
                           list_item.SubItems(32) = IIf(IsNull(rs!vcha_rco_banco_deposito), "", rs!vcha_rco_banco_deposito)
                           list_item.SubItems(33) = IIf(IsNull(rs!dtim_rco_fecha_deposito), "", rs!dtim_rco_fecha_deposito)
                           list_item.SubItems(34) = IIf(IsNull(rs!inte_rco_numero_deposito), 0, rs!inte_rco_numero_deposito)
                           list_item.SubItems(35) = IIf(IsNull(rs!dtim_rco_fecha_insercion), "", rs!dtim_rco_fecha_insercion)
                           var_importe = var_importe + IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe)
                           rs.MoveNext:
                           numero_items_rutas = numero_items_rutas + 1
                      Wend
                      txt_importe = Format(var_importe, "###,###.##")
                      rs.Close
                      n = lv_detalle.ListItems.Count
                      For i = 1 To n
                          lv_detalle.ListItems.Item(i).Selected = True
                          If lv_detalle.selectedItem.SubItems(12) = "*" Then
                              lv_detalle.selectedItem.ForeColor = &HC000&
                              lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC000&
                              lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC000&
                              lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC000&
                              lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC000&
                              lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC000&
                              lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC000&
                              lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC000&
                              lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC000&
                              lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC000&
                              lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC000&
                              lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC000&
                           Else
                              If (lv_detalle.selectedItem.SubItems(7) * 1) < (lv_detalle.selectedItem.SubItems(11) * 1) Then
                                 lv_detalle.selectedItem.ForeColor = &HC0C0&
                                 lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC0C0&
                                 lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC0C0&
                                 lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC0C0&
                                 lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC0C0&
                                 lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC0C0&
                                 lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC0C0&
                                 lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC0C0&
                                 lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC0C0&
                                 lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC0C0&
                                 lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC0C0&
                                 lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC0C0&
                              Else
                                 lv_detalle.selectedItem.ForeColor = &HFF0000
                                 lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF0000
                                 lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF0000
                                 lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF0000
                                 lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF0000
                                 lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF0000
                                 lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF0000
                                 lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF0000
                              End If
                              If (lv_detalle.selectedItem.SubItems(6) * 1) < (lv_detalle.selectedItem.SubItems(10) * 1) Then
                                 lv_detalle.selectedItem.ForeColor = &HFF&
                                 lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF&
                                 lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF&
                                 lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF&
                                 lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF&
                                 lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF&
                                 lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF&
                                 lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF&
                                 lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HFF&
                                 lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HFF&
                                 lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HFF&
                                 lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HFF&
                              End If
                           End If
                        Next i
                     End If
                  End If
               End If
            Else
               MsgBox "Documento invalido para el cliente seleccionado", vbOKOnly, "ATENCION"
            End If
            If rs.State = 1 Then
               rs.Close
            End If
         Else
            MsgBox "No se a indicado un número de documento", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a indicado un tipo de documento", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un cliente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_aplicar_Click()
   Me.frm_cambios_relacion.Visible = False
   Dim si As Integer
   Dim var_descuento_aplicar As Double
   Dim var_descuento_agente As Double
   Dim var_descuento_sistema As Double
   Dim var_porcentaje_iva As Double
   Dim var_porcentaje_impuesto_2 As Double
   Dim var_porcentaje_impuesto_3 As Double
   Dim var_subimporte As Double
   Dim var_importe_iva As Double
   Dim var_importe_impuesto_2 As Double
   Dim var_importe_impuesto_3 As Double
   Dim var_importe_sin_impuesto As Double
   Dim var_importe_descuento_aplicar As Double
   Dim var_importe_descuento_1 As Double
   Dim var_almacen As String
   Dim var_grupo_actual As String
   Dim var_grupo_real As String
   Dim var_cliente As String
   Dim var_titular As String
   Dim var_establecimiento As String
   Dim var_clave_moneda As String
   Dim var_agente As String
   Dim var_posible_tipo_cambio As Boolean
   Dim var_moneda_local As Integer
   Dim var_tipo_Cambio As Double
   Dim var_importe_total As Double
   Dim var_importe_total_cobranza As Double
   Dim var_importe As Double
   Dim var_numero_nota As Double
   Dim var_cheque As String
   Dim var_importe_saldo As Double
   Dim var_importe_cobranza As Double
   Dim var_descuento_saldo As Double
   Dim var_tipo_documento As String
   Dim var_banco As String
   Dim var_fecha_factura As Date
   Dim i, j As Integer
   Dim var_numero_folio As Double
   var_n = lv_detalle.ListItems.Count
   Dim var_posible_pagos As Boolean
   var_posible_pagos = False
             
   cnn.BeginTrans
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_RELACION_COBRANZA", cnn, adOpenDynamic, adLockOptimistic
   If rs.EOF Then
      var_consecutivo = 1
   Else
      var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
   End If
   rsaux.Open "insert into TB_TEMP_RELACION_COBRANZA (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
   rs.Close
   cnn.CommitTrans
   For var_i = 1 To var_n
       lv_detalle.ListItems.Item(var_i).Selected = True
       If Trim(lv_detalle.selectedItem.SubItems(12)) <> "*" Then
          If Trim(lv_detalle.selectedItem.SubItems(13)) = "*" Then
             rs.Open "INSERT INTO TB_TEMP_RELACION_COBRANZA (INTE_TEM_CONSECUTIVO,VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO) VALUES (" + CStr(var_consecutivo) + ", '" + var_empresa + "','" + Trim(lv_detalle.selectedItem.SubItems(1)) + "','" + Trim(lv_detalle.selectedItem.SubItems(24)) + "'," + lv_detalle.selectedItem.SubItems(2) + ")", cnn, adOpenDynamic, adLockOptimistic
          End If
          var_posible_pagos = True
       End If
   Next var_i
   
   If var_posible_pagos = True Then
      var_cadena = ""
   rs.Open "select * from vw_temp_relacion_cobranza_veces where veces > 1 and inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            If Trim(var_cadena) = "" Then
               var_cadena = var_cadena + CStr(rs!inte_car_numero)
            Else
               var_cadena = var_cadena + "," + CStr(rs!inte_car_numero)
            End If
            rs.MoveNext
      Wend
   End If
   rs.Close
   If Trim(var_cadena) = "" Then
   si = MsgBox("¿Deseas aplicar los pagos seleccionados?", vbYesNo, "ATENCION")
   If si = 6 Then
      si = MsgBox("Confirmar la aplicación de los pagos seleccionados", vbYesNo, "ATENCION")
      If si = 6 Then
         If lv_detalle.ListItems.Count > 0 Then
            n = lv_detalle.ListItems.Count
            var_posible_tipo_cambio = True
            For i = 1 To n
               lv_detalle.ListItems.Item(i).Selected = True
               var_clave_moneda = lv_detalle.selectedItem.SubItems(18)
                              
               rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
               var_moneda_local = 1
               If Not rs.EOF Then
                  var_moneda_local = IIf(IsNull(rs!inte_mon_moneda_local), 0, rs!inte_mon_moneda_local)
               End If
               If var_moneda_local = 0 Then
                  rsaux2.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                  If rsaux2.EOF Then
                     var_posible_tipo_cambio = False
                  End If
                  rsaux2.Close
               End If
               rs.Close
            Next i
            If var_posible_tipo_cambio = True Then
               Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
               Set TB_DEVOLUCIONES_ESTATUS = New TB_DEVOLUCIONES_ESTATUS
               Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
               For i = 1 To n
                  lv_detalle.ListItems.Item(i).Selected = True
                  If Trim(lv_detalle.selectedItem.SubItems(32)) <> "" Then
                     If lv_detalle.selectedItem.SubItems(13) = "*" Then
                        If Trim(lv_detalle.selectedItem) <> "" Then
                           If (lv_detalle.selectedItem.SubItems(6) * 1) < (lv_detalle.selectedItem.SubItems(10) * 1) Then
                              lv_detalle.selectedItem.SubItems(12) = "*"
                              cnn.BeginTrans
                              var_movimiento = 0
                              var_numero_movimiento = 0
                              var_almacen = ""
                              var_tipo_documento = lv_detalle.selectedItem.SubItems(1)
                              var_grupo_actual = lv_detalle.selectedItem.SubItems(14)
                              var_grupo_real = lv_detalle.selectedItem.SubItems(15)
                              var_titular = lv_detalle.selectedItem.SubItems(16)
                              var_agente = txt_clave_agente
                              var_cliente = lv_detalle.selectedItem
                              var_establecimiento = lv_detalle.selectedItem.SubItems(17)
                              var_clave_moneda = lv_detalle.selectedItem.SubItems(18)
                             
                              var_importe_total_cobranza = lv_detalle.selectedItem.SubItems(10)
                              var_cheque = lv_detalle.selectedItem.SubItems(8)
                              var_descuento_sistema = (lv_detalle.selectedItem.SubItems(7) * 1)
                              var_descuento_agente = (lv_detalle.selectedItem.SubItems(11) * 1)
                              var_descuento_aplicar = 0
                              var_porcentaje_iva = (lv_detalle.selectedItem.SubItems(19) * 1)
                              var_porcentaje_impuesto_2 = (lv_detalle.selectedItem.SubItems(20) * 1)
                              var_porcentaje_impuesto_3 = (lv_detalle.selectedItem.SubItems(21) * 1)
                              var_serie = lv_detalle.selectedItem.SubItems(24)
                              var_banco = lv_detalle.selectedItem.SubItems(25)
                              If var_descuento_agente < var_descuento_sistema Then
                                 var_descuento_aplicar = var_descuento_agente
                              End If
                              If var_descuento_sistema < var_descuento_agente Then
                                 var_descuento_aplicar = var_descuento_sistema
                              End If
                              If var_descuento_sistema = var_descuento_agente Then
                                 var_descuento_aplicar = var_descuento_sistema
                              End If
                              var_descuento_aplicar = 0
                              var_tipo_Cambio = 1
                              rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                              var_moneda_local = 1
                              If Not rs.EOF Then
                                 var_moneda_local = IIf(IsNull(rs!inte_mon_moneda_local), 0, rs!inte_mon_moneda_local)
                              End If
                              If var_moneda_local = 0 Then
                                 rsaux2.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux2.EOF Then
                                    var_tipo_Cambio = rsaux2!mone_tca_importe
                                 End If
                                 rsaux2.Close
                              End If
                              rs.Close
                              var_insertar = False
                              var_importe_total_cobranza = var_importe_total_cobranza * var_tipo_Cambio
                              var_importe_total = (var_importe_total_cobranza) * (1 + (var_descuento_aplicar / 100))
                              var_subimporte = var_importe_total / (1 + (var_descuento_aplicar / 100))
                              var_importe_descuento_1 = var_importe_total - var_subimporte
                              var_importe_descuento_2 = 0
                              var_importe_descuento_3 = 0
                              var_importe_iva = var_importe_total_cobranza - (var_importe_total_cobranza) / (1 + (var_porcentaje_iva / 100))
                              If var_porcentaje_impuesto_2 > 0 Then
                                 var_importe_impuesto_2 = (var_importe_total_cobranza - var_importe_iva) / (var_importe_total_cobranza - var_importe_iva) / (1 + (var_porcentaje_impuesto_2 / 100))
                              Else
                                 var_importe_impuesto_2 = 0
                              End If
                              If var_porcentaje_impuesto_3 > 0 Then
                                 var_importe_impuesto_3 = (var_importe_total_cobranza - var_importe - iva_var_impuesto_2) / (var_importe_total_cobranza - var_importe_iva - var_impuesto_2) / (1 + (var_porcentaje_impuesto_3 / 100))
                              Else
                                 var_importe_impuesto_3 = 0
                              End If
                              
                              'rs.Open "select maximo_numero from vw_maximo_numero_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and  vcha_car_tipo_documento = 'PA' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                              'If rs.EOF Then
                              '   var_numero_folio = 0
                              'Else
                              '   var_numero_folio = IIf(IsNull(rs!maximo_numero), 0, rs!maximo_numero)
                              'End If
                              'rs.Close
                              
                              
                              rs.Open "select * from TB_MAXIMO_PAGO", cnn_sid_quezada, adOpenDynamic, adLockOptimistic
                              If rs.EOF Then
                                 var_numero_folio = 0
                              Else
                                 var_numero_folio = IIf(IsNull(rs!inte_max_maximo_pago), 0, rs!inte_max_maximo_pago)
                              End If
                              rs.Close
                              
                              
                              var_numero_folio = var_numero_folio + 1
                              var_importe_sin_impuesto = var_importe_total_cobranza - (var_importe_iva + var_importe_descuento_2 + var_importe_descuento_3)
                              'var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "PA", "PA", "PA", var_numero_folio, "-", "", "", 0, CStr(Date), var_agente, var_grupo_actual, var_grupo_real, var_titular, var_cliente, var_establecimiento, 0, var_porcentaje_iva, var_porcentaje_impuesto_2, var_porcentaje_impuesto_3, var_descuento_aplicar, 0, 0, var_importe_total, var_importe_iva, var_importe_impuesto_2, var_importe_impuesto_3, var_importe_descuento_1, 0, 0, var_importe_sin_impuesto, var_importe_total_cobranza, "", var_clave_usuario_global, "", Date, 0, Date, Date, var_clave_moneda, var_tipo_Cambio, var_serie, "")
                        
                              Cadena = "INSERT INTO TB_ENCABEZADO_CARTERA (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, "
                              Cadena = Cadena + "FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3, FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, VCHA_SER_SERIE_ID, CHAR_CAR_ESTATUS, VCHA_CAR_CHEQUE_DEPOSITO, VCHA_CAR_BANCO_CHEQUE, VCHA_CAR_DEPOSITO, VCHA_CAR_BANCO_DEPOSITO, DTIM_CAR_FECHA_DEPOSITO) values ("
                              Cadena = Cadena + "'" + var_empresa + "', '" + var_unidad_organizacional + "', 'PA', 'PA', 'PA', " + CStr(var_numero_folio) + ", '-', '', '', 0, getdate(), '" + var_agente + "', '" + var_grupo_actual + "', '" + var_grupo_real + "', '" + var_titular + "', '" + var_cliente + "', '" + var_establecimiento + "', 0, " + CStr(var_porcentaje_iva) + ", " + CStr(var_porcentaje_impuesto_2) + ", " + CStr(var_porcentaje_impuesto_3) + ", " + CStr(var_descuento_aplicar) + ", 0, 0, " + CStr(var_importe_total) + ", " + CStr(var_importe_iva) + ", " + CStr(var_importe_impuesto_2) + ", " + CStr(var_importe_impuesto_3) + ", " + CStr(var_importe_descuento_1) + ", 0, 0, " + CStr(var_importe_sin_impuesto) + ", " + CStr(var_importe_total_cobranza) + ", '', '"
                              Cadena = Cadena + CStr(var_clave_usuario_global) + ", '', getdate(), 0, getdate(), getdate(), '" + var_clave_moneda + "', " + CStr(var_tipo_Cambio) + ", '" + var_serie + "', '','" + Me.lv_detalle.selectedItem.SubItems(29) + "', '" + lv_detalle.selectedItem.SubItems(30) + "', '" + lv_detalle.selectedItem.SubItems(31) + "','" + lv_detalle.selectedItem.SubItems(32) + "','" + lv_detalle.selectedItem.SubItems(33) + "')"
                              rsaux5.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                              rsaux5.Open "update TB_MAXIMO_PAGO set inte_max_maximo_pago = inte_max_maximo_pago + 1", cnn_sid_quezada, adOpenDynamic, adLockOptimistic
                              
                              var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, var_serie, "NA", var_numero_folio, var_serie, "PA", var_numero_folio, 0, var_importe_total_cobranza)
                              
                              rs.Open "update tb_relacion_cobranza set char_rco_aplicada = '*', FLOA_RCO_TIPO_CAMBIO = " + Str(var_tipo_Cambio) + ", INTE_RCO_PAGO = " + Str(var_numero_folio) + ",FLOA_RCO_DESCUENTO_APLICAR = 0, dtim_rco_fecha_aplicacion = '" + Str(Day(Date)) + "/" + Str(Month(Date)) + "/" + Str(Year(Date)) + "' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_cheque = '" + var_cheque + "' and vcha_cli_clave_id = '" + var_cliente + "' and vcha_ban_banco_id = '" + var_banco + "' and vcha_rco_folio = '" + txt_folio + "' and vcha_car_documento = '" + var_tipo_documento + "' and inte_Car_numero = " + lv_detalle.selectedItem.SubItems(2) + " and inte_rco_partida = " + Me.lv_detalle.selectedItem.SubItems(27), cnn, adOpenDynamic, adLockOptimistic
                              
                              cnn.CommitTrans
                           Else
                              lv_detalle.selectedItem.SubItems(12) = "*"
                              cnn.BeginTrans
                              var_serie = lv_detalle.selectedItem.SubItems(24)
                              var_banco = lv_detalle.selectedItem.SubItems(25)
                              
                              rsaux2.Open "select floa_sal_importe from tb_saldos where  VCHA_CAR_DOCUMENTO = '" + lv_detalle.selectedItem.SubItems(1) + "' AND vcha_emp_empresa_id = '" + var_empresa + "' and inte_car_numero = " + Str(lv_detalle.selectedItem.SubItems(2)) + " and vcha_Ser_Serie_id ='" + var_serie + "' AND FLOA_SAL_IMPORTE IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
                              var_importe_saldo = Format(IIf(IsNull(rsaux2!FLOA_sAL_IMPORTE), 0, rsaux2!FLOA_sAL_IMPORTE), "###,##0.00")
                              rsaux2.Close
                             
                              var_movimiento = 0
                              var_numero_movimiento = 0
                              var_almacen = ""
                              var_tipo_documento = lv_detalle.selectedItem.SubItems(1)
                              var_grupo_actual = lv_detalle.selectedItem.SubItems(14)
                              var_grupo_real = lv_detalle.selectedItem.SubItems(15)
                              var_titular = lv_detalle.selectedItem.SubItems(16)
                              var_agente = txt_clave_agente
                              var_cliente = lv_detalle.selectedItem
                              var_establecimiento = lv_detalle.selectedItem.SubItems(17)
                              var_clave_moneda = lv_detalle.selectedItem.SubItems(18)
                              var_importe_total_cobranza = lv_detalle.selectedItem.SubItems(10)
                              var_cheque = lv_detalle.selectedItem.SubItems(8)
                              var_descuento_sistema = (lv_detalle.selectedItem.SubItems(7) * 1)
                              var_descuento_agente = (lv_detalle.selectedItem.SubItems(11) * 1)
                              var_descuento_aplicar = 0
                              var_porcentaje_iva = (lv_detalle.selectedItem.SubItems(19) * 1)
                              var_porcentaje_impuesto_2 = (lv_detalle.selectedItem.SubItems(20) * 1)
                              var_porcentaje_impuesto_3 = (lv_detalle.selectedItem.SubItems(21) * 1)
                              var_tipo_Cambio = 1
                              rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                              var_moneda_local = 1
                              If Not rs.EOF Then
                                 var_moneda_local = IIf(IsNull(rs!inte_mon_moneda_local), 0, rs!inte_mon_moneda_local)
                              End If
                              If var_moneda_local = 0 Then
                                 rsaux2.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux2.EOF Then
                                    var_tipo_Cambio = rsaux2!mone_tca_importe
                                 End If
                                 rsaux2.Close
                              End If
                              rs.Close
                              If var_descuento_agente < var_descuento_sistema Then
                                 var_descuento_aplicar = var_descuento_agente
                              End If
                              If var_descuento_sistema < var_descuento_agente Then
                                 var_descuento_aplicar = var_descuento_sistema
                              End If
                              If var_descuento_sistema = var_descuento_agente Then
                                 var_descuento_aplicar = var_descuento_sistema
                              End If
                              var_insertar = False
                              var_importe_cobranza = (var_importe_total_cobranza * 100) / (100 - var_descuento_aplicar)
                              var_importe_total_cobranza = var_importe_total_cobranza * var_tipo_Cambio
                              var_importe_total = (var_importe_total_cobranza) * (1 + (var_descuento_aplicar / 100))
                              var_subimporte = var_importe_total / (1 + (var_descuento_aplicar / 100))
                              var_importe_descuento_1 = var_importe_total - var_subimporte
                              var_importe_descuento_2 = 0
                              var_importe_descuento_3 = 0
                              var_importe_iva = var_importe_total_cobranza - (var_importe_total_cobranza) / (1 + (var_porcentaje_iva / 100))
                              If var_porcentaje_impuesto_2 > 0 Then
                                 var_importe_impuesto_2 = (var_importe_total_cobranza - var_importe_iva) / (var_importe_total_cobranza - var_importe_iva) / (1 + (var_porcentaje_impuesto_2 / 100))
                              Else
                                 var_importe_impuesto_2 = 0
                              End If
                              If var_porcentaje_impuesto_3 > 0 Then
                                 var_importe_impuesto_3 = (var_importe_total_cobranza - var_importe - iva_var_impuesto_2) / (var_importe_total_cobranza - var_importe_iva - var_impuesto_2) / (1 + (var_porcentaje_impuesto_3 / 100))
                              Else
                                 var_importe_impuesto_3 = 0
                              End If
                              
                              'rs.Open "select maximo_numero from vw_maximo_numero_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and  vcha_car_tipo_documento = 'PA' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                              'If rs.EOF Then
                              '   var_numero_folio = 0
                              'Else
                              '   var_numero_folio = IIf(IsNull(rs!maximo_numero), 0, rs!maximo_numero)
                              'End If
                              'rs.Close
                              
                              
                              rs.Open "select * from TB_MAXIMO_PAGO", cnn_sid_quezada, adOpenDynamic, adLockOptimistic
                              If rs.EOF Then
                                 var_numero_folio = 0
                              Else
                                 var_numero_folio = IIf(IsNull(rs!inte_max_maximo_pago), 0, rs!inte_max_maximo_pago)
                              End If
                              rs.Close
                              
                              
                              var_numero_folio = var_numero_folio + 1
                              var_importe_sin_impuesto = var_importe_total_cobranza - (var_importe_iva + var_importe_descuento_2 + var_importe_descuento_3)
                        
                              ' 2.- se elimina por la elaboracion del procedimiento almacenado de sp_relacion_cobranza
                              'Cadena = "INSERT INTO TB_ENCABEZADO_CARTERA (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, "
                              'Cadena = Cadena + "FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3, FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, VCHA_SER_SERIE_ID, CHAR_CAR_ESTATUS, VCHA_CAR_CHEQUE_DEPOSITO, VCHA_CAR_BANCO_CHEQUE, VCHA_CAR_DEPOSITO, VCHA_CAR_BANCO_DEPOSITO, DTIM_CAR_FECHA_DEPOSITO) values ("
                              'Cadena = Cadena + "'" + var_empresa + "', '" + var_unidad_organizacional + "', 'PA', 'PA', 'PA', " + CStr(var_numero_folio) + ", '-', '', '', 0, '" + CStr(Date) + "', '" + var_agente + "', '" + var_grupo_actual + "', '" + var_grupo_real + "', '" + var_titular + "', '" + var_cliente + "', '" + var_establecimiento + "', 0, " + CStr(var_porcentaje_iva) + ", " + CStr(var_porcentaje_impuesto_2) + ", " + CStr(var_porcentaje_impuesto_3) + ", " + CStr(var_descuento_aplicar) + ", 0, 0, " + CStr(var_importe_total) + ", " + CStr(var_importe_iva) + ", " + CStr(var_importe_impuesto_2) + ", " + CStr(var_importe_impuesto_3) + ", " + CStr(var_importe_descuento_1) + ", 0, 0, " + CStr(var_importe_sin_impuesto) + ", " + CStr(var_importe_total_cobranza) + ", '', '"
                              'Cadena = Cadena + var_clave_usuario_global + "', '', getDate(), 0, getDate(), getDate(), '" + var_clave_moneda + "', " + CStr(var_tipo_Cambio) + ", '" + var_serie + "', '','" + Me.lv_detalle.selectedItem.SubItems(29) + "', '" + lv_detalle.selectedItem.SubItems(30) + "', '" + lv_detalle.selectedItem.SubItems(31) + "','" + lv_detalle.selectedItem.SubItems(32) + "','" + lv_detalle.selectedItem.SubItems(33) + "')"
                              'rsaux5.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                        
                              'var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, var_serie, var_tipo_documento, lv_detalle.selectedItem.SubItems(2), var_serie, "PA", var_numero_folio, 0, var_importe_total_cobranza)
                              
                              'If Round(var_importe_saldo, 2) <= Round(var_importe_cobranza, 2) Then
                              '   var_descuento_aplicar = 100 - ((var_importe_total_cobranza * 100) / var_importe_saldo)
                              '   rsaux2.Open "update tb_relacion_cobranza set floa_rco_descuento_aplicar = " + Str(var_descuento_aplicar) + ", dtim_rco_fecha_aplicacion =  '" + Str(Day(Date)) + "/" + Str(Month(Date)) + "/" + Str(Year(Date)) + "'  where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_cheque = '" + var_cheque + "' and vcha_cli_clave_id = '" + var_cliente + "' and vcha_ban_banco_id = '" + var_banco + "' and inte_car_numero = " + lv_detalle.selectedItem.SubItems(2) + " and vcha_rco_folio = '" + txt_folio + "' and vcha_car_documento = '" + var_tipo_documento + "' and inte_rco_partida = " + Me.lv_detalle.selectedItem.SubItems(27), cnn, adOpenDynamic, adLockOptimistic
                              'End If
                              'rs.Open "update tb_relacion_cobranza set char_rco_aplicada = '*', FLOA_RCO_TIPO_CAMBIO = " + Str(var_tipo_Cambio) + ", INTE_RCO_PAGO = " + Str(var_numero_folio) + ", dtim_rco_fecha_aplicacion =  '" + Str(Day(Date)) + "/" + Str(Month(Date)) + "/" + Str(Year(Date)) + "', dtim_rco_fecha_asignacion = '" + Str(Day(Date)) + "/" + Str(Month(Date)) + "/" + Str(Year(Date)) + "'   where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_cheque = '" + var_cheque + "' and vcha_cli_clave_id = '" + var_cliente + "' and vcha_ban_banco_id = '" + var_banco + "' and vcha_rco_folio = '" + txt_folio + "' and inte_Car_numero = " + lv_detalle.selectedItem.SubItems(2) + " and vcha_car_documento = '" + var_tipo_documento + "' and inte_rco_partida = " + Me.lv_detalle.selectedItem.SubItems(27), cnn, adOpenDynamic, adLockOptimistic
                              
                              '2.- fin de la eliminacion del procedimiento almacenado
                              
                              var_cadena = "exec sp_relacion_cobranza '" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_agente + "', '" + var_grupo_actual + "', '" + var_grupo_real + "','" + var_titular + "','" + var_cliente + "', '" + var_establecimiento + "', " + CStr(var_porcentaje_iva) + ", " + CStr(var_descuento_aplicar) + ", " + CStr(var_importe_total) + ", " + CStr(var_importe_iva) + ", " + CStr(var_importe_descuento_1) + ", " + CStr(var_importe_sin_impuesto) + ", " + CStr(var_importe_total_cobranza) + ",'" + var_clave_usuario_global + "', '" + var_clave_moneda + "', " + CStr(var_tipo_Cambio) + ", '" + var_serie + "', '" + Me.lv_detalle.selectedItem.SubItems(29) + "', '" + lv_detalle.selectedItem.SubItems(30) + "', '" + lv_detalle.selectedItem.SubItems(31) + "','" + lv_detalle.selectedItem.SubItems(32) + "','" + lv_detalle.selectedItem.SubItems(33) + "','" + Me.txt_folio + "'," + lv_detalle.selectedItem.SubItems(2) + ",'" + var_cheque + "','" + var_tipo_documento + "',"
                              var_cadena = var_cadena + "'" + var_banco + "'," + Me.lv_detalle.selectedItem.SubItems(27) + ", " + CStr(var_importe_saldo) + "," + CStr(var_importe_cobranza)
                              cnn.CommandTimeout = 360
                              rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              
                              rs.Open "update tb_maximo_pago set inte_max_maximo_pago = inte_max_maximo_pago + 1", cnn_sid_quezada, adOpenDynamic, adLockOptimistic
                              '1.- esto al parecer no sirve de nada 02/10/2006
                              'If var_tipo_documento = "FA" Then
                              '   rs.Open "select * from VW_DETALLE_FACTURACION_LINEAS WHERE VCHA_EMP_EMPRESA_ID =  '" + var_empresa + "' AND VCHA_SER_SERIE_ID = '" + var_serie + "' AND VCHA_CAR_TIPO_DOCUMENTO = 'FA' and inte_Car_numero = " + lv_detalle.selectedItem.SubItems(2), cnn, adOpenDynamic, adLockOptimistic
                              '   While Not rs.EOF
                              '         var_fecha_factura = CDate(Format(CStr(rs!dtim_Car_fecha), "short date"))
                              '         var_cadena = "INSERT INTO TB_COMISIONES_APLICADAS ([VCHA_EMP_EMPRESA_ID], [VCHA_AGE_AGENTE_ID], [VCHA_CAR_TIPO_DOCUMENTO], [VCHA_SER_SERIE_ID], [INTE_CAR_NUMERO], [DTIM_CAR_FECHA], [FLOA_CAP_IMPORTE_FACTURA], [VCHA_RCO_FOLIO], [DTIM_CAP_FECHA_PAGO], [VCHA_LIN_LINEA_ID], [FLOA_CAP_IMPORTE_PARTICIPACION], [FLOA_CAP_PORCENTAJE_PARTICIPACION], [FLOA_COM_PORCENTAJE], [FLOA_CAP_IMPORTE_COMISION], [VCHA_BAN_BANCO_ID] , [VCHA_RCO_CHEQUE], [FLOA_CAP_IMPORTE_PAGO], [VCHA_CLI_CLAVE_ID])"
                              '         var_cadena = var_cadena + "Values ('" + var_empresa + "', '" + var_agente + "', 'FA', '" + var_serie + "', " + lv_detalle.selectedItem.SubItems(2) + ", '" + Str(Day(var_fecha_factura)) + "/" + Str(Month(var_fecha_factura)) + "/" + Str(Year(var_fecha_factura)) + "', " + CStr(rs!floa_car_importe_neto / rs!FLOA_cAR_TIPO_cAMBIO) + ", '" + txt_folio + "', '" + Str(Day(Date)) + "/" + Str(Month(Date)) + "/" + Str(Year(Date)) + "', '" + rs!VCHA_LIN_LINEA_ID + "', " + CStr(rs!importe / rs!FLOA_cAR_TIPO_cAMBIO) + ", 0, 0, 0,'" + var_banco + "' ,'" + var_cheque + "', " + CStr(var_importe_sin_impuesto) + ", '" + var_cliente + "')"
                              '         rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              '         rs.MoveNext
                              '   Wend
                              '   rs.Close
                              'End If
                              '1.- fin de eliminacion de lo que no sirve para nada 02/10/2006
                              cnn.CommitTrans
                           End If
                        End If
                      End If
                   Else
                      MsgBox "No se a indicado el banco del deposito", vbOKOnly, "ATENCION"
                   End If
                Next i
                Call Command1_Click
''''''''''''
                Dim var_fecha_cheques As String
                var_dia = CStr(Day(Date))
                var_mes = CStr(Month(Date))
                var_año = CStr(Year(Date))
                If Len(Trim(var_dia)) = 1 Then
                   var_dia = "0" + var_dia
                End If
                If Len(Trim(var_mes)) = 1 Then
                   var_mes = "0" + var_mes
                End If
               var_fecha_cheques = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
               rs.Open "select * from tb_relacion_cobranza with (nolock) where vcha_rco_folio = '" + txt_folio + "' and dtim_rco_fecha_cheque <= " + var_fecha_cheques, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_importe = 0
                  txt_clave_agente = rs!VCHA_AGE_AGENTE_ID
                  rsaux2.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_nombre_agente = rsaux2!VCHA_AGE_NOMBRE
                  rsaux2.Close
                  txt_fecha = rs!dtim_rco_fecha_relacion
                  lv_detalle.ListItems.Clear
                  While Not rs.EOF
                     Set list_item = lv_detalle.ListItems.Add(, , IIf(IsNull(rs!vcha_Cli_clave_id), "", rs!vcha_Cli_clave_id))
                     list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
                     list_item.SubItems(2) = IIf(IsNull(rs!inte_car_numero), "", rs!inte_car_numero)
                     list_item.SubItems(3) = IIf(IsNull(rs!dtim_car_fecha), "", Format(rs!dtim_car_fecha, "Short Date"))
                     list_item.SubItems(4) = IIf(IsNull(rs!floa_Car_importe), 0, Round(rs!floa_Car_importe, 2))
                     rsaux2.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        list_item.SubItems(5) = IIf(IsNull(rsaux2!vcha_mon_nombre_plural), "", Trim(rsaux2!vcha_mon_nombre_plural))
                     End If
                     rsaux2.Close
                     rsaux2.Open "select floa_sal_importe from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_car_numero = " + Str(rs!inte_car_numero) + " and vcha_ser_Serie_id = '" + rs!vcha_ser_Serie_id + "' and vcha_cli_clave_id = '" + IIf(IsNull(rs!vcha_Cli_clave_id), "", rs!vcha_Cli_clave_id) + "' AND VCHA_CAR_DOCUMENTO = '" + rs!vcha_Car_documento + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        list_item.SubItems(6) = Format(IIf(IsNull(rsaux2!FLOA_sAL_IMPORTE), 0, rsaux2!FLOA_sAL_IMPORTE), "###,##0.00")
                     Else
                        list_item.SubItems(6) = 0
                     End If
                     rsaux2.Close
                     list_item.SubItems(7) = IIf(IsNull(rs!floa_Car_descuento_aplicado), 0, rs!floa_Car_descuento_aplicado)
                     list_item.SubItems(8) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
                     list_item.SubItems(9) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
                     list_item.SubItems(10) = Format(IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe), "###,##0.00")
                     list_item.SubItems(11) = IIf(IsNull(rs!FLOA_RCO_DESCUENTO_OTORGADO), 0, rs!FLOA_RCO_DESCUENTO_OTORGADO)
                     list_item.SubItems(12) = IIf(IsNull(rs!char_rco_aplicada), "", rs!char_rco_aplicada)
                     list_item.SubItems(14) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
                     list_item.SubItems(15) = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
                     list_item.SubItems(16) = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
                     list_item.SubItems(17) = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
                     list_item.SubItems(18) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                     list_item.SubItems(19) = IIf(IsNull(rs!floa_rco_iva), 0, rs!floa_rco_iva)
                     list_item.SubItems(20) = IIf(IsNull(rs!floa_rco_impuesto_2), 0, rs!floa_rco_impuesto_2)
                     list_item.SubItems(21) = IIf(IsNull(rs!floa_rco_impuesto_3), 0, rs!floa_rco_impuesto_3)
                     list_item.SubItems(22) = IIf(IsNull(rs!floa_Rco_descuento_aplicar), 0, rs!floa_Rco_descuento_aplicar)
                     list_item.SubItems(23) = IIf(IsNull(rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO), 0, rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO)
                     list_item.SubItems(24) = IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id)
                     list_item.SubItems(25) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
                     list_item.SubItems(26) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
                     list_item.SubItems(27) = IIf(IsNull(rs!inte_rco_partida), "", rs!inte_rco_partida)
                     list_item.SubItems(28) = IIf(IsNull(rs!dtim_rco_fecha_relacion), "", rs!dtim_rco_fecha_relacion)
                     list_item.SubItems(29) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
                     list_item.SubItems(30) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
                     list_item.SubItems(31) = IIf(IsNull(rs!vcha_rco_deposito), "", rs!vcha_rco_deposito)
                     list_item.SubItems(32) = IIf(IsNull(rs!vcha_rco_banco_deposito), "", rs!vcha_rco_banco_deposito)
                     list_item.SubItems(33) = IIf(IsNull(rs!dtim_rco_fecha_deposito), "", rs!dtim_rco_fecha_deposito)
                     list_item.SubItems(34) = IIf(IsNull(rs!inte_rco_numero_deposito), 0, rs!inte_rco_numero_deposito)
                     list_item.SubItems(35) = IIf(IsNull(rs!dtim_rco_fecha_insercion), "", rs!dtim_rco_fecha_insercion)
                     var_importe = var_importe + IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe)
                     rs.MoveNext:
                     numero_items_rutas = numero_items_rutas + 1
                  Wend
                  txt_importe = Format(var_importe, "###,###.##")
                  rs.Close
                  n = lv_detalle.ListItems.Count
                  For i = 1 To n
                     lv_detalle.ListItems.Item(i).Selected = True
                     If lv_detalle.selectedItem.SubItems(12) = "*" Then
                        lv_detalle.selectedItem.ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC000&
                     Else
                        If (lv_detalle.selectedItem.SubItems(7) * 1) < (lv_detalle.selectedItem.SubItems(11) * 1) Then
                           lv_detalle.selectedItem.ForeColor = &HC0C0&
                           lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC0C0&
                           lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC0C0&
                           lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC0C0&
                           lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC0C0&
                           lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC0C0&
                           lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC0C0&
                           lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC0C0&
                           lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC0C0&
                           lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC0C0&
                           lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC0C0&
                           lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC0C0&
                        Else
                           lv_detalle.selectedItem.ForeColor = &HFF0000
                           lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF0000
                           lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF0000
                           lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF0000
                           lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF0000
                           lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF0000
                           lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF0000
                           lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF0000
                        End If
                        If (lv_detalle.selectedItem.SubItems(6) * 1) < (lv_detalle.selectedItem.SubItems(10) * 1) Then
                           lv_detalle.selectedItem.ForeColor = &HFF&
                           lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF&
                           lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF&
                           lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF&
                           lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF&
                           lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF&
                           lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF&
                           lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF&
                           lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HFF&
                           lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HFF&
                           lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HFF&
                           lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HFF&
                        End If
                     End If
                  Next i
                End If
            
''''''''''
                MsgBox "Se a terminado aplicar los pagos seleccionados", vbOKOnly, "ATENCION"
             Else
                MsgBox "No se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
             End If
          Else
             MsgBox "No existen pagos a aplicar", vbOKOnly, "ATENCION"
          End If
       End If
    End If
    Else
        MsgBox "Los documentos numero " + var_cadena + " se aplican dos veces", vbOKOnly, "ATENCION"
        var_dia = CStr(Day(Date))
        var_mes = CStr(Month(Date))
        var_año = CStr(Year(Date))
        If Len(Trim(var_dia)) = 1 Then
           var_dia = "0" + var_dia
        End If
        If Len(Trim(var_mes)) = 1 Then
           var_mes = "0" + var_mes
        End If
        var_fecha_cheques = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
        If rs.State = 1 Then
           rs.Close
        End If
        rs.Open "select * from tb_relacion_cobranza with (nolock) where vcha_rco_folio = '" + txt_folio + "' and dtim_rco_fecha_cheque <= " + var_fecha_cheques, cnn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
           var_importe = 0
           txt_clave_agente = rs!VCHA_AGE_AGENTE_ID
           rsaux2.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
           txt_nombre_agente = rsaux2!VCHA_AGE_NOMBRE
           rsaux2.Close
           txt_fecha = rs!dtim_rco_fecha_relacion
           lv_detalle.ListItems.Clear
           While Not rs.EOF
                 Set list_item = lv_detalle.ListItems.Add(, , IIf(IsNull(rs!vcha_Cli_clave_id), "", rs!vcha_Cli_clave_id))
                 list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
                 list_item.SubItems(2) = IIf(IsNull(rs!inte_car_numero), "", rs!inte_car_numero)
                 list_item.SubItems(3) = IIf(IsNull(rs!dtim_car_fecha), "", Format(rs!dtim_car_fecha, "Short Date"))
                 list_item.SubItems(4) = IIf(IsNull(rs!floa_Car_importe), 0, Round(rs!floa_Car_importe, 2))
                 rsaux2.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                 If Not rsaux2.EOF Then
                    list_item.SubItems(5) = IIf(IsNull(rsaux2!vcha_mon_nombre_plural), "", Trim(rsaux2!vcha_mon_nombre_plural))
                 End If
                 rsaux2.Close
                 rsaux2.Open "select floa_sal_importe from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_car_numero = " + Str(rs!inte_car_numero) + " and vcha_ser_Serie_id = '" + rs!vcha_ser_Serie_id + "' and vcha_cli_clave_id = '" + IIf(IsNull(rs!vcha_Cli_clave_id), "", rs!vcha_Cli_clave_id) + "' AND VCHA_CAR_DOCUMENTO = '" + rs!vcha_Car_documento + "'", cnn, adOpenDynamic, adLockOptimistic
                 If Not rsaux2.EOF Then
                    list_item.SubItems(6) = Format(IIf(IsNull(rsaux2!FLOA_sAL_IMPORTE), 0, rsaux2!FLOA_sAL_IMPORTE), "###,##0.00")
                 Else
                    list_item.SubItems(6) = 0
                 End If
                 rsaux2.Close
                 list_item.SubItems(7) = IIf(IsNull(rs!floa_Car_descuento_aplicado), 0, rs!floa_Car_descuento_aplicado)
                 list_item.SubItems(8) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
                 list_item.SubItems(9) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
                 list_item.SubItems(10) = Format(IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe), "###,##0.00")
                 list_item.SubItems(11) = IIf(IsNull(rs!FLOA_RCO_DESCUENTO_OTORGADO), 0, rs!FLOA_RCO_DESCUENTO_OTORGADO)
                 list_item.SubItems(12) = IIf(IsNull(rs!char_rco_aplicada), "", rs!char_rco_aplicada)
                 list_item.SubItems(14) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
                 list_item.SubItems(15) = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
                 list_item.SubItems(16) = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
                 list_item.SubItems(17) = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
                 list_item.SubItems(18) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                 list_item.SubItems(19) = IIf(IsNull(rs!floa_rco_iva), 0, rs!floa_rco_iva)
                 list_item.SubItems(20) = IIf(IsNull(rs!floa_rco_impuesto_2), 0, rs!floa_rco_impuesto_2)
                 list_item.SubItems(21) = IIf(IsNull(rs!floa_rco_impuesto_3), 0, rs!floa_rco_impuesto_3)
                 list_item.SubItems(22) = IIf(IsNull(rs!floa_Rco_descuento_aplicar), 0, rs!floa_Rco_descuento_aplicar)
                 list_item.SubItems(23) = IIf(IsNull(rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO), 0, rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO)
                 list_item.SubItems(24) = IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id)
                 list_item.SubItems(25) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
                 list_item.SubItems(26) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
                 list_item.SubItems(27) = IIf(IsNull(rs!inte_rco_partida), "", rs!inte_rco_partida)
                 list_item.SubItems(28) = IIf(IsNull(rs!dtim_rco_fecha_relacion), "", rs!dtim_rco_fecha_relacion)
                 list_item.SubItems(29) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
                 list_item.SubItems(30) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
                 list_item.SubItems(31) = IIf(IsNull(rs!vcha_rco_deposito), "", rs!vcha_rco_deposito)
                 list_item.SubItems(32) = IIf(IsNull(rs!vcha_rco_banco_deposito), "", rs!vcha_rco_banco_deposito)
                 list_item.SubItems(33) = IIf(IsNull(rs!dtim_rco_fecha_deposito), "", rs!dtim_rco_fecha_deposito)
                 list_item.SubItems(34) = IIf(IsNull(rs!inte_rco_numero_deposito), 0, rs!inte_rco_numero_deposito)
                 list_item.SubItems(35) = IIf(IsNull(rs!dtim_rco_fecha_insercion), "", rs!dtim_rco_fecha_insercion)
                 var_importe = var_importe + IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe)
                 rs.MoveNext:
                 numero_items_rutas = numero_items_rutas + 1
           Wend
           txt_importe = Format(var_importe, "###,###.##")
           rs.Close
           n = lv_detalle.ListItems.Count
           For i = 1 To n
               lv_detalle.ListItems.Item(i).Selected = True
               If lv_detalle.selectedItem.SubItems(12) = "*" Then
                  lv_detalle.selectedItem.ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC000&
               Else
                  If (lv_detalle.selectedItem.SubItems(7) * 1) < (lv_detalle.selectedItem.SubItems(11) * 1) Then
                     lv_detalle.selectedItem.ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC0C0&
                  Else
                     lv_detalle.selectedItem.ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF0000
                  End If
                  If (lv_detalle.selectedItem.SubItems(6) * 1) < (lv_detalle.selectedItem.SubItems(10) * 1) Then
                     lv_detalle.selectedItem.ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HFF&
                  End If
               End If
           Next i
        End If
    End If
    Else
       MsgBox "No existen pagos a aplicar", vbOKOnly, "ATENCION"
    End If
    If rs.State = 1 Then
       rs.Close
    End If
    rs.Open "delete from tb_temp_relacion_cobranza where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
End Sub


Private Sub cmd_imprimir_Click()

End Sub

Private Sub cmd_cancelar_cambios_Click()
   Me.frm_cambios_relacion.Visible = False
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   Me.frm_cambios_relacion.Visible = False
   Dim n, i As Integer
   Dim var_posible As Boolean
   n = lv_detalle.ListItems.Count
   var_posible = False
   For i = 1 To n
      lv_detalle.ListItems.Item(i).Selected = True
      If lv_detalle.selectedItem.ListSubItems(12) = "*" And lv_detalle.selectedItem.ListSubItems(23) = 0 Then
         var_posible = True
      End If
   Next i
   rs.Open "select distinct vcha_cli_nombre from VW_NOTA_CREDITO_RELACION_COBRANZA with (nolock) where vcha_rco_folio = '" + txt_folio + "' and vcha_emp_empresa_id = '" + var_empresa + "'order by vcha_cli_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
   If rs.EOF Then
      var_posible = False
   End If
   rs.Close
   If var_posible = True Then
      frmnota_credito_descuento_financiero.txt_relacion_cobranza = txt_folio
      frmnota_credito_descuento_financiero.txt_clave_agente = txt_clave_agente
      frmnota_credito_descuento_financiero.txt_nombre_agente = txt_nombre_agente
      frmnota_credito_descuento_financiero.Show 1
      'frmrelacion_cobranza.Enabled = False
                
                
                
      rs.Open "select * from tb_relacion_cobranza with (nolock) where vcha_rco_folio = '" + txt_folio + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_importe = 0
         txt_clave_agente = rs!VCHA_AGE_AGENTE_ID
         If rsaux2.State = 1 Then
            rsaux2.Close
         End If
         rsaux2.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
         txt_nombre_agente = rsaux2!VCHA_AGE_NOMBRE
         rsaux2.Close
         txt_fecha = rs!dtim_rco_fecha_relacion
         lv_detalle.ListItems.Clear
         While Not rs.EOF
            Set list_item = lv_detalle.ListItems.Add(, , IIf(IsNull(rs!vcha_Cli_clave_id), "", rs!vcha_Cli_clave_id))
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
            list_item.SubItems(2) = IIf(IsNull(rs!inte_car_numero), "", rs!inte_car_numero)
            list_item.SubItems(3) = IIf(IsNull(rs!dtim_car_fecha), "", Format(rs!dtim_car_fecha, "Short Date"))
            list_item.SubItems(4) = IIf(IsNull(rs!floa_Car_importe), 0, Round(rs!floa_Car_importe, 2))
            rsaux2.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               list_item.SubItems(5) = IIf(IsNull(rsaux2!vcha_mon_nombre_plural), "", Trim(rsaux2!vcha_mon_nombre_plural))
            Else
               list_item.SubItems(5) = ""
            End If
            rsaux2.Close
            rsaux2.Open "select floa_sal_importe from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_car_numero = " + Str(rs!inte_car_numero) + " AND VCHA_CAR_DOCUMENTO = '" + rs!vcha_Car_documento + "' AND VCHA_SER_SERIE_ID = '" + rs!vcha_ser_Serie_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               list_item.SubItems(6) = Format(IIf(IsNull(rsaux2!FLOA_sAL_IMPORTE), 0, rsaux2!FLOA_sAL_IMPORTE), "###,##0.00")
            Else
               list_item.SubItems(6) = 0
            End If
            rsaux2.Close
            list_item.SubItems(7) = IIf(IsNull(rs!floa_Car_descuento_aplicado), 0, rs!floa_Car_descuento_aplicado)
            list_item.SubItems(8) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
            list_item.SubItems(9) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
            list_item.SubItems(10) = Format(IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe), "###,##0.00")
            list_item.SubItems(11) = IIf(IsNull(rs!FLOA_RCO_DESCUENTO_OTORGADO), 0, rs!FLOA_RCO_DESCUENTO_OTORGADO)
            list_item.SubItems(12) = IIf(IsNull(rs!char_rco_aplicada), "", rs!char_rco_aplicada)
            list_item.SubItems(14) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
            list_item.SubItems(15) = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
            list_item.SubItems(16) = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
            list_item.SubItems(17) = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
            list_item.SubItems(18) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
            list_item.SubItems(19) = IIf(IsNull(rs!floa_rco_iva), 0, rs!floa_rco_iva)
            list_item.SubItems(20) = IIf(IsNull(rs!floa_rco_impuesto_2), 0, rs!floa_rco_impuesto_2)
            list_item.SubItems(21) = IIf(IsNull(rs!floa_rco_impuesto_3), 0, rs!floa_rco_impuesto_3)
            list_item.SubItems(22) = IIf(IsNull(rs!floa_Rco_descuento_aplicar), 0, rs!floa_Rco_descuento_aplicar)
            list_item.SubItems(23) = IIf(IsNull(rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO), 0, rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO)
            list_item.SubItems(24) = IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id)
            var_importe = var_importe + IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe)
            rs.MoveNext:
            numero_items_rutas = numero_items_rutas + 1
         Wend
         txt_importe = Format(var_importe, "###,###.##")
         rs.Close
         n = lv_detalle.ListItems.Count
            For i = 1 To n
               lv_detalle.ListItems.Item(i).Selected = True
               If lv_detalle.selectedItem.SubItems(12) = "*" Then
                  lv_detalle.selectedItem.ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC000&
               Else
                  If (lv_detalle.selectedItem.SubItems(7) * 1) < (lv_detalle.selectedItem.SubItems(11) * 1) Then
                     lv_detalle.selectedItem.ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC0C0&
                  Else
                     lv_detalle.selectedItem.ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF0000
                  End If
                  If (lv_detalle.selectedItem.SubItems(6) * 1) < (lv_detalle.selectedItem.SubItems(10) * 1) Then
                     lv_detalle.selectedItem.ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HFF&
                  End If
               End If
            Next i
       End If
                
                
   
   Else
      MsgBox "No existen descuentos financieros que aplicar", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command2_Click()
  
End Sub

Private Sub Form_GotFocus()
   Me.frm_cambios_relacion.Visible = False
End Sub

Private Sub Form_Load()
   frm_lista.Visible = False
   var_cadena_seguridad = ""
   Top = 0
   Left = 0
   Set var_tabla = CreateObject("ADODB.connection")
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open "select vcha_pri_ruta_cobranza from tb_principal", cnn, adOpenDynamic, adLockOptimistic
   var_ruta = IIf(IsNull(rs!VCHA_PRI_RUTA_COBRANZA), "", rs!VCHA_PRI_RUTA_COBRANZA)
   rs.Close
   frm_cambios_relacion.Visible = False
Dim var_posible_clientes As Boolean
Dim var_posible_facturas As Boolean
Dim var_moneda_local As Integer
Dim var_clave_agente As String
Dim var_importe As Double
Dim var_partida As Double
'On Error GoTo salir:
txt_folio = frmrelacion_cobranza_listado.txt_folio
   If Trim(txt_folio) <> "" Then
      If Trim(var_ruta) <> "" Then
         Dim var_fecha_cheques As String
         var_dia = CStr(Day(Date))
         var_mes = CStr(Month(Date))
         var_año = CStr(Year(Date))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         var_fecha_cheques = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
         
         cnn.CommandTimeout = 360
         rs.Open "select * from tb_relacion_cobranza with (nolock) where vcha_rco_folio = '" + txt_folio + "' and vcha_emp_empresa_id = '" + var_empresa + "' and dtim_rco_fecha_cheque <= " + var_fecha_cheques + "+1-.00001", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_importe = 0
            txt_clave_agente = rs!VCHA_AGE_AGENTE_ID
            rsaux2.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               txt_nombre_agente = rsaux2!VCHA_AGE_NOMBRE
            Else
               GoTo salir_agente:
            End If
            rsaux2.Close
            txt_fecha = rs!dtim_rco_fecha_relacion
            lv_detalle.ListItems.Clear
            While Not rs.EOF
               Set list_item = lv_detalle.ListItems.Add(, , IIf(IsNull(rs!vcha_Cli_clave_id), "", rs!vcha_Cli_clave_id))
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
               list_item.SubItems(2) = IIf(IsNull(rs!inte_car_numero), "", rs!inte_car_numero)
               list_item.SubItems(3) = IIf(IsNull(rs!dtim_car_fecha), "", Format(rs!dtim_car_fecha, "Short Date"))
               list_item.SubItems(4) = IIf(IsNull(rs!floa_Car_importe), 0, Round(rs!floa_Car_importe, 2))
               rsaux2.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id) + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  list_item.SubItems(5) = IIf(IsNull(rsaux2!vcha_mon_nombre_plural), "", Trim(rsaux2!vcha_mon_nombre_plural))
               End If
               rsaux2.Close
               rsaux2.Open "select floa_sal_importe from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_car_numero = " + Str(rs!inte_car_numero) + " and vcha_ser_Serie_id = '" + IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id) + "' and vcha_cli_clave_id = '" + IIf(IsNull(rs!vcha_Cli_clave_id), "", rs!vcha_Cli_clave_id) + "' AND VCHA_CAR_DOCUMENTO = '" + rs!vcha_Car_documento + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  list_item.SubItems(6) = Format(IIf(IsNull(rsaux2!FLOA_sAL_IMPORTE), 0, rsaux2!FLOA_sAL_IMPORTE), "###,##0.00")
               Else
                  list_item.SubItems(6) = 0
               End If
               rsaux2.Close
               list_item.SubItems(7) = IIf(IsNull(rs!floa_Car_descuento_aplicado), 0, rs!floa_Car_descuento_aplicado)
               list_item.SubItems(8) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
               list_item.SubItems(9) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
               list_item.SubItems(10) = Format(IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe), "###,##0.00")
               list_item.SubItems(11) = IIf(IsNull(rs!FLOA_RCO_DESCUENTO_OTORGADO), 0, rs!FLOA_RCO_DESCUENTO_OTORGADO)
               list_item.SubItems(12) = IIf(IsNull(rs!char_rco_aplicada), "", rs!char_rco_aplicada)
               list_item.SubItems(14) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
               list_item.SubItems(15) = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
               list_item.SubItems(16) = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
               list_item.SubItems(17) = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
               list_item.SubItems(18) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
               list_item.SubItems(19) = IIf(IsNull(rs!floa_rco_iva), 0, rs!floa_rco_iva)
               list_item.SubItems(20) = IIf(IsNull(rs!floa_rco_impuesto_2), 0, rs!floa_rco_impuesto_2)
               list_item.SubItems(21) = IIf(IsNull(rs!floa_rco_impuesto_3), 0, rs!floa_rco_impuesto_3)
               list_item.SubItems(22) = IIf(IsNull(rs!floa_Rco_descuento_aplicar), 0, rs!floa_Rco_descuento_aplicar)
               list_item.SubItems(23) = IIf(IsNull(rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO), 0, rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO)
               list_item.SubItems(24) = IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id)
               list_item.SubItems(25) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
               list_item.SubItems(26) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
               list_item.SubItems(27) = IIf(IsNull(rs!inte_rco_partida), "", rs!inte_rco_partida)
               list_item.SubItems(28) = IIf(IsNull(rs!dtim_rco_fecha_relacion), "", rs!dtim_rco_fecha_relacion)
               list_item.SubItems(29) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
               list_item.SubItems(30) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
               list_item.SubItems(31) = IIf(IsNull(rs!vcha_rco_deposito), "", rs!vcha_rco_deposito)
               list_item.SubItems(32) = IIf(IsNull(rs!vcha_rco_banco_deposito), "", rs!vcha_rco_banco_deposito)
               list_item.SubItems(33) = IIf(IsNull(rs!dtim_rco_fecha_deposito), "", rs!dtim_rco_fecha_deposito)
               list_item.SubItems(34) = IIf(IsNull(rs!inte_rco_numero_deposito), 0, rs!inte_rco_numero_deposito)
               list_item.SubItems(35) = IIf(IsNull(rs!dtim_rco_fecha_insercion), "", rs!dtim_rco_fecha_insercion)
               var_importe = var_importe + IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe)
               rs.MoveNext:
               numero_items_rutas = numero_items_rutas + 1
            Wend
            txt_importe = Format(var_importe, "###,###.##")
            rs.Close
            n = lv_detalle.ListItems.Count
            For i = 1 To n
               lv_detalle.ListItems.Item(i).Selected = True
               If lv_detalle.selectedItem.SubItems(12) = "*" Then
                  lv_detalle.selectedItem.ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC000&
               Else
                  If (lv_detalle.selectedItem.SubItems(7) * 1) < (lv_detalle.selectedItem.SubItems(11) * 1) Then
                     lv_detalle.selectedItem.ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC0C0&
                  Else
                     lv_detalle.selectedItem.ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF0000
                  End If
                  If (lv_detalle.selectedItem.SubItems(6) * 1) < (lv_detalle.selectedItem.SubItems(10) * 1) Then
                     lv_detalle.selectedItem.ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HFF&
                  End If
               End If
            Next i
         Else
            z = 0
            If z = 1 Then
            rs.Close
            If Trim(var_ruta) <> "" Then
               If var_tabla.State = 1 Then
                  var_tabla.Close
               End If
               var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
            End If
            rsaux2.Open "select distinct cve_client,factura from " + var_ruta + "\" + Trim(txt_folio), var_tabla, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               var_posible_clientes = True
               var_posible_facturas = True
               var_posible_tipo_cambio = True
               While Not rsaux2.EOF
                  rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + Trim(rsaux2!cve_client) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If rs.EOF Then
                     var_posible_clientes = False
                  End If
                  rs.Close
                  rs.Open "select * from tb_encabezado_Cartera where vcha_cli_clave_id = '" + Trim(rsaux2!cve_client) + "' and inte_car_numero = " + Trim(CStr(rsaux2!FACTURA)), cnn, adOpenDynamic, adLockOptimistic
                  If rs.EOF Then
                     var_posible_facturas = False
                  End If
                  rs.Close
                  rsaux2.MoveNext
               Wend
            End If
            rsaux2.Close
            If var_posible_clientes = True Then
               If var_posible_facturas = True Then
                  rsaux2.Open "select * from " + var_ruta + "\" + Trim(txt_folio), var_tabla, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     cnn.BeginTrans
                     rs.Open "select a.vcha_age_agente_id from tb_agentes a, tb_rutas b where b.vcha_rut_ruta_id = '" + Trim(CStr(rsaux2!cve_agente)) + "' and a.vcha_age_agente_id = b.vcha_age_agente_id", cnn, adOpenDynamic, adLockOptimistic
                     var_clave_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                     rs.Close
                     rs.Open "select vcha_mon_moneda_id from tb_clientes where vcha_cli_clave_id = '" + Trim(rsaux2!cve_client) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                     End If
                     rs.Close
                     While Not rsaux2.EOF
                        rsaux3.Open "INSERT INTO TB_RELACION_COBRANZA([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_RCO_FOLIO], [DTIM_RCO_FECHA_RELACION], [VCHA_AGE_AGENTE_ID], [VCHA_CLI_CLAVE_ID], [VCHA_RCO_CHEQUE],[DTIM_RCO_FECHA_CHEQUE], [FLOA_RCO_IMPORTE],[FLOA_RCO_DESCUENTO_OTORGADO], [INTE_CAR_NUMERO], [FLOA_CAR_IMPORTE], [FLOA_CAR_DESCUENTO_APLICADO],[INTE_RCO_PARTIDA], [INTE_RCO_DESCUENTO_APLICADO],[VCHA_SER_SERIE_ID],[VCHA_CAR_DOCUMENTO],[VCHA_BAN_BANCO_ID]) values ( '" + var_empresa + "', '" + var_unidad_organizacional + "', '" + txt_folio + "', '" + CStr(Date) + "', '" + var_clave_agente + "', '" + Trim(CStr(rsaux2!cve_client)) + "', '" + CStr(rsaux2!cheque) + "', '" + CStr(rsaux2!fecha_cheq) + "', " + CStr(rsaux2!importe2) + ", " + CStr(rsaux2!descuento) + ", " + Trim(CStr(rsaux2!FACTURA)) + ", 0, 0, " + Str(var_partida) + ", 0, '" + Trim(rsaux2!Serie) + "', '" + rsaux2!tipo + "', '" + rsaux2!cvecuenta + "')", cnn, adOpenDynamic, adLockOptimistic
                        rsaux2.MoveNext
                     Wend
                     rsaux2.Close
                     cnn.CommitTrans
                     rs.Open "select * from tb_relacion_cobranza with (nolock) where vcha_rco_folio = '" + txt_folio + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_importe = 0
                        txt_clave_agente = rs!VCHA_AGE_AGENTE_ID
                        rsaux2.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                        txt_nombre_agente = rsaux2!VCHA_AGE_NOMBRE
                        rsaux2.Close
                        txt_fecha = rs!dtim_rco_fecha_relacion
                        lv_detalle.ListItems.Clear
                        While Not rs.EOF
                           Set list_item = lv_detalle.ListItems.Add(, , rs!vcha_Cli_clave_id)
                           list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
                           list_item.SubItems(2) = IIf(IsNull(rs!inte_car_numero), "", rs!inte_car_numero)
                           list_item.SubItems(3) = IIf(IsNull(rs!dtim_car_fecha), "", Format(rs!dtim_car_fecha, "Short Date"))
                           list_item.SubItems(4) = IIf(IsNull(rs!floa_Car_importe), 0, Round(rs!floa_Car_importe, 2))
                           rsaux2.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux2.EOF Then
                              list_item.SubItems(5) = IIf(IsNull(rsaux2!vcha_mon_nombre_plural), "", Trim(rsaux2!vcha_mon_nombre_plural))
                           End If
                           rsaux2.Close
                           rsaux2.Open "select floa_sal_importe from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_car_numero = " + Str(rs!inte_car_numero) + " and vcha_cli_clave_id = '" + rs!vcha_Cli_clave_id + "' AND VCHA_CAR_DOCUMENTO = '" + rs!vcha_Car_documento + "'", cnn, adOpenDynamic, adLockOptimistic
                           list_item.SubItems(6) = Format(IIf(IsNull(rsaux2!FLOA_sAL_IMPORTE), 0, rsaux2!FLOA_sAL_IMPORTE), "###,##0.00")
                           rsaux2.Close
                           list_item.SubItems(7) = IIf(IsNull(rs!floa_Car_descuento_aplicado), 0, rs!floa_Car_descuento_aplicado)
                           list_item.SubItems(8) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
                           list_item.SubItems(9) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
                           list_item.SubItems(10) = IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe)
                           list_item.SubItems(11) = IIf(IsNull(rs!FLOA_RCO_DESCUENTO_OTORGADO), 0, rs!FLOA_RCO_DESCUENTO_OTORGADO)
                           list_item.SubItems(12) = IIf(IsNull(rs!char_rco_aplicada), "", rs!char_rco_aplicada)
                           list_item.SubItems(14) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
                           list_item.SubItems(15) = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
                           list_item.SubItems(16) = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
                           list_item.SubItems(17) = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
                           list_item.SubItems(18) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                           list_item.SubItems(19) = IIf(IsNull(rs!floa_rco_iva), 0, rs!floa_rco_iva)
                           list_item.SubItems(20) = IIf(IsNull(rs!floa_rco_impuesto_2), 0, rs!floa_rco_impuesto_2)
                           list_item.SubItems(21) = IIf(IsNull(rs!floa_rco_impuesto_3), 0, rs!floa_rco_impuesto_3)
                           list_item.SubItems(22) = IIf(IsNull(rs!floa_Rco_descuento_aplicar), 0, rs!floa_Rco_descuento_aplicar)
                           list_item.SubItems(23) = IIf(IsNull(rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO), 0, rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO)
                           list_item.SubItems(24) = IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id)
                           list_item.SubItems(25) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
                           list_item.SubItems(26) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
                           list_item.SubItems(27) = IIf(IsNull(rs!inte_rco_partida), "", rs!inte_rco_partida)
                           list_item.SubItems(28) = IIf(IsNull(rs!dtim_rco_fecha_relacion), "", rs!dtim_rco_fecha_relacion)
                           list_item.SubItems(29) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
                           list_item.SubItems(30) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
                           list_item.SubItems(31) = IIf(IsNull(rs!vcha_rco_deposito), "", rs!vcha_rco_deposito)
                           list_item.SubItems(32) = IIf(IsNull(rs!vcha_rco_banco_deposito), "", rs!vcha_rco_banco_deposito)
                           list_item.SubItems(33) = IIf(IsNull(rs!dtim_rco_fecha_deposito), "", rs!dtim_rco_fecha_deposito)
                           list_item.SubItems(34) = IIf(IsNull(rs!inte_rco_numero_deposito), 0, rs!inte_rco_numero_deposito)
                           list_item.SubItems(35) = IIf(IsNull(rs!dtim_rco_fecha_insercion), "", rs!dtim_rco_fecha_insercion)
                           var_importe = var_importe + IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe)
                           rs.MoveNext:
                           numero_items_rutas = numero_items_rutas + 1
                        Wend
                        txt_importe = Format(var_importe, "###,###.##")
                        n = lv_detalle.ListItems.Count
                        For i = 1 To n
                           lv_detalle.ListItems.Item(i).Selected = True
                           If (lv_detalle.selectedItem.SubItems(7) * 1) < (lv_detalle.selectedItem.SubItems(11) * 1) Then
                              lv_detalle.selectedItem.ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC0C0&
                           Else
                              lv_detalle.selectedItem.ForeColor = &HFF0000
                              lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF0000
                              lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF0000
                              lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF0000
                              lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF0000
                              lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF0000
                              lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF0000
                              lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF0000
                           End If
                           If (lv_detalle.selectedItem.SubItems(6) * 1) < (lv_detalle.selectedItem.SubItems(10) * 1) Then
                              lv_detalle.selectedItem.ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HFF&
                           End If
                        Next i
                     End If
                     rs.Close
                  Else
                     rsaux2.Close
                     MsgBox "La relación de cobranza número " + txt_folio + " no existe", vbOKOnly, "ATENCION"
                  End If
                  var_tabla.Close
               Else
                  MsgBox "Existen inconcistencias en las factura, favor de revisar el archivo " + txt_folio, vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Existen inconcistencias con los clientes, favor de revisar el archivo " + txt_folio, vbOKOnly, "ATENCION"
            End If
            End If
         End If
      Else
         MsgBox "No se a indicado una ruta en donde se encuentren los archivos de cobranza", vbOKOnly, "ATENCION"
      End If
   End If
Exit Sub
On Error GoTo salir:
salir:
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   MsgBox "La relación de cobranza número " + txt_folio + " no existe", vbOKOnly, "ATENCION"
   Exit Sub
salir_agente:
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   MsgBox "La clave del agente es incorrecta", vbOKOnly, "ATENCION"
   Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_relacion_cobranza)
End Sub

Private Sub lv_detalle_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Marque los pagos a aplicar presionando enter y presione F6 para cambiar la información del pago"
   Me.frm_cambios_relacion.Visible = False
End Sub

Private Sub lv_detalle_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      x = 0
      If x = 1 Then
      txt_clave_cliente = lv_detalle.selectedItem
      txt_cheque = lv_detalle.selectedItem.SubItems(8)
      txt_documento = lv_detalle.selectedItem.SubItems(1)
      txt_numero = lv_detalle.selectedItem.SubItems(2)
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE from tb_clientes order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      var_tipo_lista = 8
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
   End If
   If KeyCode = 117 Then
      x = 0
      If x = 1 Then
      If lv_detalle.ListItems.Count > 0 Then
         If (lv_detalle.selectedItem.SubItems(6) * 1) < (lv_detalle.selectedItem.SubItems(10) * 1) Then
            If Trim(lv_detalle.selectedItem.SubItems(13)) <> "*" Then
               rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + lv_detalle.selectedItem + "' AND VCHA_AGE_AGENTE_ID = '" + Me.txt_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  txt_nombre_cliente = rs!vcha_cli_nombre
               End If
               rs.Close
               txt_clave_cliente = lv_detalle.selectedItem
               txt_cheque = lv_detalle.selectedItem.SubItems(8)
               txt_documento = lv_detalle.selectedItem.SubItems(1)
               txt_numero = lv_detalle.selectedItem.SubItems(2)
               Me.frm_cambios_relacion.Visible = True
            Else
               MsgBox "El pago ya fue aplicado y es imposible cambiar la información", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se puede cambiar la información de la factura seleccionada", vbOKOnly, "ATENCION"
         End If
      End If
      End If
   End If
   If KeyCode = 119 Then
      frmrelacion_cobranza_correccion.txt_relacion = Me.txt_folio
      frmrelacion_cobranza_correccion.txt_agente = Me.txt_clave_agente
      frmrelacion_cobranza_correccion.txt_nombre_agente = Me.txt_nombre_agente
      frmrelacion_cobranza_correccion.txt_fecha = Format(CDate(Me.txt_fecha), "Short Date")
      frmrelacion_cobranza_correccion.txt_fecha_insercion = lv_detalle.selectedItem.SubItems(35)
      
      frmrelacion_cobranza_correccion.txt_cheque = lv_detalle.selectedItem.SubItems(29)
      frmrelacion_cobranza_correccion.txt_consecutivo = lv_detalle.selectedItem.SubItems(27)
      frmrelacion_cobranza_correccion.txt_banco_cheque = lv_detalle.selectedItem.SubItems(30)
      
      rsaux5.Open "select *  from tb_bancos where vcha_ban_banco_id = '" + lv_detalle.selectedItem.SubItems(30) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux5.EOF Then
         frmrelacion_cobranza_correccion.txt_nombre_banco_cheque = IIf(IsNull(rsaux5!VCHA_BAN_NOMBRE), "", rsaux5!VCHA_BAN_NOMBRE)
      End If
      rsaux5.Close
      
      frmrelacion_cobranza_correccion.txt_fecha_cheque = lv_detalle.selectedItem.SubItems(9)
      frmrelacion_cobranza_correccion.txt_deposito = lv_detalle.selectedItem.SubItems(31)
      frmrelacion_cobranza_correccion.txt_banco = lv_detalle.selectedItem.SubItems(32)
      frmrelacion_cobranza_correccion.txt_fecha_deposito = lv_detalle.selectedItem.SubItems(33)
      frmrelacion_cobranza_correccion.txt_numero_deposito = lv_detalle.selectedItem.SubItems(34)
      frmrelacion_cobranza_correccion.txt_consecutivo = lv_detalle.selectedItem.SubItems(27)
      
      rsaux5.Open "select *  from tb_bancos where vcha_ban_banco_id = '" + lv_detalle.selectedItem.SubItems(32) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux5.EOF Then
         frmrelacion_cobranza_correccion.txt_nombre_banco = IIf(IsNull(rsaux5!VCHA_BAN_NOMBRE), "", rsaux5!VCHA_BAN_NOMBRE)
      End If
      rsaux5.Close
      
      'frmrelacion_cobranza_correccion.txt_fecha = lv_detalle.selectedItem.SubItems(33)
      frmrelacion_cobranza_correccion.Show 1
      
      Dim var_fecha_cheques As String
      var_dia = CStr(Day(Date))
      var_mes = CStr(Month(Date))
      var_año = CStr(Year(Date))
      If Len(Trim(var_dia)) = 1 Then
         var_dia = "0" + var_dia
      End If
      If Len(Trim(var_mes)) = 1 Then
         var_mes = "0" + var_mes
      End If
      var_fecha_cheques = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
      
      rs.Open "select * from tb_relacion_cobranza with (nolock) where vcha_rco_folio = '" + txt_folio + "' and dtim_rco_fecha_cheque <= " + var_fecha_cheques, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_importe = 0
         txt_clave_agente = rs!VCHA_AGE_AGENTE_ID
         rsaux2.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
         txt_nombre_agente = rsaux2!VCHA_AGE_NOMBRE
         rsaux2.Close
         txt_fecha = rs!dtim_rco_fecha_relacion
         lv_detalle.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_detalle.ListItems.Add(, , IIf(IsNull(rs!vcha_Cli_clave_id), "", rs!vcha_Cli_clave_id))
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
               list_item.SubItems(2) = IIf(IsNull(rs!inte_car_numero), "", rs!inte_car_numero)
               list_item.SubItems(3) = IIf(IsNull(rs!dtim_car_fecha), "", Format(rs!dtim_car_fecha, "Short Date"))
               list_item.SubItems(4) = IIf(IsNull(rs!floa_Car_importe), 0, Round(rs!floa_Car_importe, 2))
               rsaux2.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id) + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  list_item.SubItems(5) = IIf(IsNull(rsaux2!vcha_mon_nombre_plural), "", Trim(rsaux2!vcha_mon_nombre_plural))
               End If
               rsaux2.Close
               rsaux2.Open "select floa_sal_importe from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_car_numero = " + Str(rs!inte_car_numero) + " and vcha_ser_Serie_id = '" + rs!vcha_ser_Serie_id + "' and vcha_cli_clave_id = '" + IIf(IsNull(rs!vcha_Cli_clave_id), "", rs!vcha_Cli_clave_id) + "' AND VCHA_cAR_DOCUMENTO = '" + rs!vcha_Car_documento + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  list_item.SubItems(6) = Format(IIf(IsNull(rsaux2!FLOA_sAL_IMPORTE), 0, rsaux2!FLOA_sAL_IMPORTE), "###,##0.00")
               Else
                  list_item.SubItems(6) = 0
               End If
               rsaux2.Close
               list_item.SubItems(7) = IIf(IsNull(rs!floa_Car_descuento_aplicado), 0, rs!floa_Car_descuento_aplicado)
               list_item.SubItems(8) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
               list_item.SubItems(9) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
               list_item.SubItems(10) = Format(IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe), "###,##0.00")
               list_item.SubItems(11) = IIf(IsNull(rs!FLOA_RCO_DESCUENTO_OTORGADO), 0, rs!FLOA_RCO_DESCUENTO_OTORGADO)
               list_item.SubItems(12) = IIf(IsNull(rs!char_rco_aplicada), "", rs!char_rco_aplicada)
               list_item.SubItems(14) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
               list_item.SubItems(15) = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
               list_item.SubItems(16) = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
               list_item.SubItems(17) = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
               list_item.SubItems(18) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
               list_item.SubItems(19) = IIf(IsNull(rs!floa_rco_iva), 0, rs!floa_rco_iva)
               list_item.SubItems(20) = IIf(IsNull(rs!floa_rco_impuesto_2), 0, rs!floa_rco_impuesto_2)
               list_item.SubItems(21) = IIf(IsNull(rs!floa_rco_impuesto_3), 0, rs!floa_rco_impuesto_3)
               list_item.SubItems(22) = IIf(IsNull(rs!floa_Rco_descuento_aplicar), 0, rs!floa_Rco_descuento_aplicar)
               list_item.SubItems(23) = IIf(IsNull(rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO), 0, rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO)
               list_item.SubItems(24) = IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id)
               list_item.SubItems(25) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
               list_item.SubItems(26) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
               list_item.SubItems(27) = IIf(IsNull(rs!inte_rco_partida), "", rs!inte_rco_partida)
               list_item.SubItems(28) = IIf(IsNull(rs!dtim_rco_fecha_relacion), "", rs!dtim_rco_fecha_relacion)
               list_item.SubItems(29) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
               list_item.SubItems(30) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
               list_item.SubItems(31) = IIf(IsNull(rs!vcha_rco_deposito), "", rs!vcha_rco_deposito)
               list_item.SubItems(32) = IIf(IsNull(rs!vcha_rco_banco_deposito), "", rs!vcha_rco_banco_deposito)
               list_item.SubItems(33) = IIf(IsNull(rs!dtim_rco_fecha_deposito), "", rs!dtim_rco_fecha_deposito)
               list_item.SubItems(34) = IIf(IsNull(rs!inte_rco_numero_deposito), 0, rs!inte_rco_numero_deposito)
               list_item.SubItems(35) = IIf(IsNull(rs!dtim_rco_fecha_insercion), "", rs!dtim_rco_fecha_insercion)
               var_importe = var_importe + IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe)
               
               rs.MoveNext:
               numero_items_rutas = numero_items_rutas + 1
         Wend
         txt_importe = Format(var_importe, "###,###.##")
         rs.Close
         n = lv_detalle.ListItems.Count
         For i = 1 To n
             lv_detalle.ListItems.Item(i).Selected = True
             If lv_detalle.selectedItem.SubItems(12) = "*" Then
                lv_detalle.selectedItem.ForeColor = &HC000&
                lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC000&
                lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC000&
                lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC000&
                lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC000&
                lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC000&
                lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC000&
                lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC000&
                lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC000&
                lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC000&
                lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC000&
                lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC000&
             Else
                If (lv_detalle.selectedItem.SubItems(7) * 1) < (lv_detalle.selectedItem.SubItems(11) * 1) Then
                   lv_detalle.selectedItem.ForeColor = &HC0C0&
                   lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC0C0&
                   lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC0C0&
                   lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC0C0&
                   lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC0C0&
                   lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC0C0&
                   lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC0C0&
                   lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC0C0&
                   lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC0C0&
                   lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC0C0&
                   lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC0C0&
                   lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC0C0&
                Else
                   lv_detalle.selectedItem.ForeColor = &HFF0000
                   lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF0000
                   lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF0000
                   lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF0000
                   lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF0000
                   lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF0000
                   lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF0000
                   lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF0000
                End If
                If (lv_detalle.selectedItem.SubItems(6) * 1) < (lv_detalle.selectedItem.SubItems(10) * 1) Then
                   lv_detalle.selectedItem.ForeColor = &HFF&
                   lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF&
                   lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF&
                   lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF&
                   lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF&
                   lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF&
                   lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF&
                   lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF&
                   lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HFF&
                   lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HFF&
                   lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HFF&
                   lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HFF&
                End If
             End If
         Next i
      End If
   End If
   If KeyCode = 120 Then
      If lv_detalle.selectedItem.SubItems(12) = "*" Then
         MsgBox "El pago ya fue aplicado y no puede ser utilizado", vbOKOnly, "ATENCION"
      Else
         var_si = MsgBox("¿Desea eliminar el pago?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar la eliminación del pago", vbYesNo, "ATENCION")
            If var_si = 6 Then
               '1
               rsaux.Open "DELETE FROM TB_RELACION_COBRANZA WHERE VCHA_RCO_FOLIO = '" + Me.txt_folio + "' AND vcha_car_documento = '" + lv_detalle.selectedItem.SubItems(1) + "' and inte_Car_numero = " + lv_detalle.selectedItem.SubItems(2) + " and vcha_rco_cheque = '" + lv_detalle.selectedItem.SubItems(8) + "' and vcha_ban_banco_id = '" + lv_detalle.selectedItem.SubItems(25) + "' and inte_rco_partida = " + lv_detalle.selectedItem.SubItems(27), cnn, adOpenDynamic, adLockOptimistic
               var_dia = CStr(Day(Date))
               var_mes = CStr(Month(Date))
               var_año = CStr(Year(Date))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_cheques = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
      
               rs.Open "select * from tb_relacion_cobranza with (nolock) where vcha_rco_folio = '" + txt_folio + "' and dtim_rco_fecha_cheque <= " + var_fecha_cheques, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_importe = 0
                  txt_clave_agente = rs!VCHA_AGE_AGENTE_ID
                  rsaux2.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_nombre_agente = rsaux2!VCHA_AGE_NOMBRE
                  rsaux2.Close
                  txt_fecha = rs!dtim_rco_fecha_relacion
                  lv_detalle.ListItems.Clear
                  While Not rs.EOF
                        Set list_item = lv_detalle.ListItems.Add(, , IIf(IsNull(rs!vcha_Cli_clave_id), "", rs!vcha_Cli_clave_id))
                        list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
                        list_item.SubItems(2) = IIf(IsNull(rs!inte_car_numero), "", rs!inte_car_numero)
                        list_item.SubItems(3) = IIf(IsNull(rs!dtim_car_fecha), "", Format(rs!dtim_car_fecha, "Short Date"))
                        list_item.SubItems(4) = IIf(IsNull(rs!floa_Car_importe), 0, Round(rs!floa_Car_importe, 2))
                        rsaux2.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           list_item.SubItems(5) = IIf(IsNull(rsaux2!vcha_mon_nombre_plural), "", Trim(rsaux2!vcha_mon_nombre_plural))
                        End If
                        rsaux2.Close
                        rsaux2.Open "select floa_sal_importe from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_car_numero = " + Str(rs!inte_car_numero) + " and vcha_ser_Serie_id = '" + rs!vcha_ser_Serie_id + "' and vcha_cli_clave_id = '" + IIf(IsNull(rs!vcha_Cli_clave_id), "", rs!vcha_Cli_clave_id) + "' AND VCHA_CAR_DOCUMENTO = '" + rs!vcha_Car_documento + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           list_item.SubItems(6) = Format(IIf(IsNull(rsaux2!FLOA_sAL_IMPORTE), 0, rsaux2!FLOA_sAL_IMPORTE), "###,##0.00")
                        Else
                           list_item.SubItems(6) = 0
                        End If
                        rsaux2.Close
                        list_item.SubItems(7) = IIf(IsNull(rs!floa_Car_descuento_aplicado), 0, rs!floa_Car_descuento_aplicado)
                        list_item.SubItems(8) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
                        list_item.SubItems(9) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
                        list_item.SubItems(10) = Format(IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe), "###,##0.00")
                        list_item.SubItems(11) = IIf(IsNull(rs!FLOA_RCO_DESCUENTO_OTORGADO), 0, rs!FLOA_RCO_DESCUENTO_OTORGADO)
                        list_item.SubItems(12) = IIf(IsNull(rs!char_rco_aplicada), "", rs!char_rco_aplicada)
                        list_item.SubItems(14) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
                        list_item.SubItems(15) = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
                        list_item.SubItems(16) = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
                        list_item.SubItems(17) = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
                        list_item.SubItems(18) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                        list_item.SubItems(19) = IIf(IsNull(rs!floa_rco_iva), 0, rs!floa_rco_iva)
                        list_item.SubItems(20) = IIf(IsNull(rs!floa_rco_impuesto_2), 0, rs!floa_rco_impuesto_2)
                        list_item.SubItems(21) = IIf(IsNull(rs!floa_rco_impuesto_3), 0, rs!floa_rco_impuesto_3)
                        list_item.SubItems(22) = IIf(IsNull(rs!floa_Rco_descuento_aplicar), 0, rs!floa_Rco_descuento_aplicar)
                        list_item.SubItems(23) = IIf(IsNull(rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO), 0, rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO)
                        list_item.SubItems(24) = IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id)
                        list_item.SubItems(25) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
                        list_item.SubItems(26) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
                        list_item.SubItems(27) = IIf(IsNull(rs!inte_rco_partida), "", rs!inte_rco_partida)
                        list_item.SubItems(28) = IIf(IsNull(rs!dtim_rco_fecha_relacion), "", rs!dtim_rco_fecha_relacion)
                        list_item.SubItems(29) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
                        list_item.SubItems(30) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
                        list_item.SubItems(31) = IIf(IsNull(rs!vcha_rco_deposito), "", rs!vcha_rco_deposito)
                        list_item.SubItems(32) = IIf(IsNull(rs!vcha_rco_banco_deposito), "", rs!vcha_rco_banco_deposito)
                        list_item.SubItems(33) = IIf(IsNull(rs!dtim_rco_fecha_deposito), "", rs!dtim_rco_fecha_deposito)
                        list_item.SubItems(34) = IIf(IsNull(rs!inte_rco_numero_deposito), 0, rs!inte_rco_numero_deposito)
                        list_item.SubItems(35) = IIf(IsNull(rs!dtim_rco_fecha_insercion), "", rs!dtim_rco_fecha_insercion)
                        var_importe = var_importe + IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe)
                        rs.MoveNext
                        numero_items_rutas = numero_items_rutas + 1
                   Wend
                   txt_importe = Format(var_importe, "###,###.##")
                   rs.Close
                   n = lv_detalle.ListItems.Count
                   For i = 1 To n
                       lv_detalle.ListItems.Item(i).Selected = True
                       If lv_detalle.selectedItem.SubItems(12) = "*" Then
                          lv_detalle.selectedItem.ForeColor = &HC000&
                          lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC000&
                          lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC000&
                          lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC000&
                          lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC000&
                          lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC000&
                          lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC000&
                          lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC000&
                          lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC000&
                          lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC000&
                          lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC000&
                          lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC000&
                       Else
                          If (lv_detalle.selectedItem.SubItems(7) * 1) < (lv_detalle.selectedItem.SubItems(11) * 1) Then
                             lv_detalle.selectedItem.ForeColor = &HC0C0&
                             lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC0C0&
                             lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC0C0&
                             lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC0C0&
                             lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC0C0&
                             lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC0C0&
                             lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC0C0&
                             lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC0C0&
                             lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC0C0&
                             lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC0C0&
                             lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC0C0&
                             lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC0C0&
                          Else
                             lv_detalle.selectedItem.ForeColor = &HFF0000
                             lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF0000
                             lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF0000
                             lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF0000
                             lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF0000
                             lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF0000
                             lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF0000
                             lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF0000
                          End If
                          If (lv_detalle.selectedItem.SubItems(6) * 1) < (lv_detalle.selectedItem.SubItems(10) * 1) Then
                             lv_detalle.selectedItem.ForeColor = &HFF&
                             lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF&
                             lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF&
                             lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF&
                             lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF&
                             lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF&
                             lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF&
                             lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF&
                             lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HFF&
                             lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HFF&
                             lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HFF&
                             lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HFF&
                          End If
                       End If
                   Next i
                End If
            End If
         End If
      End If
   End If
End Sub

Private Sub lv_detalle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim si As Integer
      If Trim(lv_detalle.selectedItem) = "" Then
         MsgBox "Se debe de indicar un cliente", vbOKOnly, "ATENCION"
      Else
         If lv_detalle.selectedItem.SubItems(12) = "*" Then
            MsgBox "El pago ya fue aplicado con anterioridad", vbOKOnly, "ATENCION"
         Else
            If lv_detalle.selectedItem.SubItems(13) = "" Then
               If ((lv_detalle.selectedItem.SubItems(7) * 1) < (lv_detalle.selectedItem.SubItems(11) * 1)) Or ((lv_detalle.selectedItem.SubItems(6) * 1) < (lv_detalle.selectedItem.SubItems(10) * 1)) Then
                  If (lv_detalle.selectedItem.SubItems(7) * 1) < (lv_detalle.selectedItem.SubItems(11) * 1) Then
                     si = MsgBox("El descuento otorgado por el agente no corresponde con el descuento que aplica por politicas ¿Desea aplicar el pago?", vbYesNo, "ATENCION")
                     If si = 6 Then
                        lv_detalle.selectedItem.SubItems(13) = "*"
                        lv_detalle.selectedItem.ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC000&
                        lv_detalle.Refresh
                        lv_detalle.SetFocus
                     End If
                  End If
                  If (lv_detalle.selectedItem.SubItems(6) * 1) < (lv_detalle.selectedItem.SubItems(10) * 1) Then
                     MsgBox "No es posible aplicar este pago", vbOKOnly, "ATENCION"
                     'si = MsgBox("El importe aplicado en la relación de cobranza es mayor al saldo de la factura por lo que no es aplicable a esta ¿Desea aplicar el pago?", vbYesNo, "ATENCION")
                     si = 8
                     If si = 6 Then
                        lv_detalle.selectedItem.SubItems(13) = "*"
                        lv_detalle.selectedItem.ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC000&
                        lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC000&
                        lv_detalle.Refresh
                        lv_detalle.SetFocus
                     End If
                  End If
               Else
                  lv_detalle.selectedItem.SubItems(13) = "*"
                  lv_detalle.selectedItem.ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC000&
               End If
            Else
               lv_detalle.selectedItem.SubItems(13) = ""
               lv_detalle.selectedItem.ForeColor = &H80000007
               lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &H80000007
               lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &H80000007
               lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &H80000007
               lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &H80000007
               lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &H80000007
               lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &H80000007
               lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &H80000007
               lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &H80000007
               lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &H80000007
               lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &H80000007
               lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &H80000007
               If (lv_detalle.selectedItem.SubItems(7) * 1) < (lv_detalle.selectedItem.SubItems(11) * 1) Then
                  lv_detalle.selectedItem.SubItems(13) = ""
                  lv_detalle.selectedItem.ForeColor = &HC0C0&
                  lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC0C0&
                  lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC0C0&
                  lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC0C0&
                  lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC0C0&
                  lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC0C0&
                  lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC0C0&
                  lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC0C0&
                  lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC0C0&
                  lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC0C0&
                  lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC0C0&
                  lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC0C0&
               End If
               If (lv_detalle.selectedItem.SubItems(6) * 1) < (lv_detalle.selectedItem.SubItems(10) * 1) Then
                  lv_detalle.selectedItem.SubItems(13) = ""
                  lv_detalle.selectedItem.ForeColor = 255
                  lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = 255
                  lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = 255
                  lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = 255
                  lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = 255
                  lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = 255
                  lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = 255
                  lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = 255
                  lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = 255
                  lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = 255
                  lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = 255
                  lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = 255
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub lv_detalle_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         var_si = MsgBox("¿Desea cambiar el cliente en la relación de cobranza?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar el cambio de cliente", vbYesNo, "ATENCION")
            If var_si = 6 Then
               cnn.BeginTrans
               '2
               rsaux.Open "DELETE FROM TB_RELACION_COBRANZA WHERE VCHA_RCO_FOLIO = '" + Me.txt_folio + "' AND vcha_car_documento = '" + lv_detalle.selectedItem.SubItems(1) + "' and inte_Car_numero = " + lv_detalle.selectedItem.SubItems(2) + " and vcha_rco_cheque = '" + lv_detalle.selectedItem.SubItems(8) + "' and vcha_ban_banco_id = '" + lv_detalle.selectedItem.SubItems(25) + "' and inte_rco_partida = " + lv_detalle.selectedItem.SubItems(27), cnn, adOpenDynamic, adLockOptimistic
               'Cadena = "EXECUTE RELACION_COBRANZA_I '" + var_empresa + "', '', '" + Me.txt_folio + "', '" + lv_detalle.selectedItem.SubItems(28) + "', '" + Me.txt_clave_agente + "', '" + lv_lista.selectedItem + "', '" + Me.txt_cheque + "', '" + lv_detalle.selectedItem.SubItems(26) + "', " + CStr(CDbl(lv_detalle.selectedItem.SubItems(10))) + ", " + lv_detalle.selectedItem.SubItems(11) + ", " + txt_numero + ", 0, 0, " + lv_detalle.selectedItem.SubItems(27) + ", 0, '" + lv_detalle.selectedItem.SubItems(24) + "', '" + txt_documento + "', '" + Trim(lv_detalle.selectedItem.SubItems(25)) + "'"
               Cadena = "INSERT INTO TB_RELACION_COBRANZA (VCHA_EMP_EMPRESA_ID, VCHA_RCO_FOLIO, DTIM_RCO_FECHA_RELACION, VCHA_AGE_AGENTE_ID, VCHA_CLI_CLAVE_ID, VCHA_RCO_CHEQUE, DTIM_RCO_FECHA_CHEQUE, FLOA_RCO_IMPORTE, FLOA_RCO_DESCUENTO_OTORGADO, INTE_CAR_NUMERO, FLOA_CAR_IMPORTE, FLOA_CAR_DESCUENTO_APLICADO, INTE_RCO_PARTIDA, INTE_RCO_DESCUENTO_APLICADO, VCHA_SER_SERIE_ID, VCHA_CAR_DOCUMENTO, VCHA_BAN_BANCO_ID, CHAR_RCO_APLICADA, VCHA_RCO_DEPOSITO, VCHA_RCO_BANCO_DEPOSITO, DTIM_RCO_FECHA_DEPOSITO) VALUES ("
               Cadena = Cadena + "'" + var_empresa + "', '" + Me.txt_folio + "', '" + lv_detalle.selectedItem.SubItems(28) + "', '" + Me.txt_clave_agente + "', '" + lv_lista.selectedItem + "', '" + Me.txt_cheque + "', '" + lv_detalle.selectedItem.SubItems(26) + "', " + CStr(CDbl(lv_detalle.selectedItem.SubItems(10))) + ", " + lv_detalle.selectedItem.SubItems(11) + ", " + txt_numero + ", 0, 0, " + lv_detalle.selectedItem.SubItems(27) + ", 0, '" + lv_detalle.selectedItem.SubItems(24) + "', '" + txt_documento + "', '" + Trim(lv_detalle.selectedItem.SubItems(25)) + "','','" + Trim(lv_detalle.selectedItem.SubItems(31)) + "','" + Trim(lv_detalle.selectedItem.SubItems(32)) + "','" + Trim(lv_detalle.selectedItem.SubItems(33)) + "')"
               rsaux.Open Cadena, cnn, adOpenDynamic, adLockOptimistic

               cnn.CommitTrans
               Dim var_fecha_cheques As String
               var_dia = CStr(Day(Date))
               var_mes = CStr(Month(Date))
               var_año = CStr(Year(Date))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_cheques = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
               If rs.State = 1 Then
                  rs.Close
               End If
               rs.Open "select * from tb_relacion_cobranza  with (nolock) where vcha_rco_folio = '" + txt_folio + "' and dtim_rco_fecha_cheque <= " + var_fecha_cheques, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_importe = 0
                  txt_clave_agente = rs!VCHA_AGE_AGENTE_ID
                  rsaux2.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_nombre_agente = rsaux2!VCHA_AGE_NOMBRE
                  rsaux2.Close
                  txt_fecha = rs!dtim_rco_fecha_relacion
                  lv_detalle.ListItems.Clear
                  While Not rs.EOF
                        Set list_item = lv_detalle.ListItems.Add(, , IIf(IsNull(rs!vcha_Cli_clave_id), "", rs!vcha_Cli_clave_id))
                        list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
                        list_item.SubItems(2) = IIf(IsNull(rs!inte_car_numero), "", rs!inte_car_numero)
                        list_item.SubItems(3) = IIf(IsNull(rs!dtim_car_fecha), "", Format(rs!dtim_car_fecha, "Short Date"))
                        list_item.SubItems(4) = IIf(IsNull(rs!floa_Car_importe), 0, Round(rs!floa_Car_importe, 2))
                        rsaux2.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           list_item.SubItems(5) = IIf(IsNull(rsaux2!vcha_mon_nombre_plural), "", Trim(rsaux2!vcha_mon_nombre_plural))
                        End If
                        rsaux2.Close
                        rsaux2.Open "select floa_sal_importe from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_car_numero = " + Str(rs!inte_car_numero) + " and vcha_ser_Serie_id = '" + rs!vcha_ser_Serie_id + "' and vcha_cli_clave_id = '" + IIf(IsNull(rs!vcha_Cli_clave_id), "", rs!vcha_Cli_clave_id) + "' AND VCHA_CAR_DOCUMENTO = '" + rs!vcha_Car_documento + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           list_item.SubItems(6) = Format(IIf(IsNull(rsaux2!FLOA_sAL_IMPORTE), 0, rsaux2!FLOA_sAL_IMPORTE), "###,##0.00")
                        Else
                           list_item.SubItems(6) = 0
                        End If
                        rsaux2.Close
                        list_item.SubItems(7) = IIf(IsNull(rs!floa_Car_descuento_aplicado), 0, rs!floa_Car_descuento_aplicado)
                        list_item.SubItems(8) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
                        list_item.SubItems(9) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
                        list_item.SubItems(10) = Format(IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe), "###,##0.00")
                        list_item.SubItems(11) = IIf(IsNull(rs!FLOA_RCO_DESCUENTO_OTORGADO), 0, rs!FLOA_RCO_DESCUENTO_OTORGADO)
                        list_item.SubItems(12) = IIf(IsNull(rs!char_rco_aplicada), "", rs!char_rco_aplicada)
                        list_item.SubItems(14) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
                        list_item.SubItems(15) = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
                        list_item.SubItems(16) = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
                        list_item.SubItems(17) = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
                        list_item.SubItems(18) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                        list_item.SubItems(19) = IIf(IsNull(rs!floa_rco_iva), 0, rs!floa_rco_iva)
                        list_item.SubItems(20) = IIf(IsNull(rs!floa_rco_impuesto_2), 0, rs!floa_rco_impuesto_2)
                        list_item.SubItems(21) = IIf(IsNull(rs!floa_rco_impuesto_3), 0, rs!floa_rco_impuesto_3)
                        list_item.SubItems(22) = IIf(IsNull(rs!floa_Rco_descuento_aplicar), 0, rs!floa_Rco_descuento_aplicar)
                        list_item.SubItems(23) = IIf(IsNull(rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO), 0, rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO)
                        list_item.SubItems(24) = IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id)
                        list_item.SubItems(25) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
                        list_item.SubItems(26) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
                        list_item.SubItems(27) = IIf(IsNull(rs!inte_rco_partida), "", rs!inte_rco_partida)
                        list_item.SubItems(28) = IIf(IsNull(rs!dtim_rco_fecha_relacion), "", rs!dtim_rco_fecha_relacion)
                        list_item.SubItems(29) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
                        list_item.SubItems(30) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
                        list_item.SubItems(31) = IIf(IsNull(rs!vcha_rco_deposito), "", rs!vcha_rco_deposito)
                        list_item.SubItems(32) = IIf(IsNull(rs!vcha_rco_banco_deposito), "", rs!vcha_rco_banco_deposito)
                        list_item.SubItems(33) = IIf(IsNull(rs!dtim_rco_fecha_deposito), "", rs!dtim_rco_fecha_deposito)
                        list_item.SubItems(34) = IIf(IsNull(rs!inte_rco_numero_deposito), 0, rs!inte_rco_numero_deposito)
                        list_item.SubItems(35) = IIf(IsNull(rs!dtim_rco_fecha_insercion), "", rs!dtim_rco_fecha_insercion)
                        var_importe = var_importe + IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe)
                        rs.MoveNext:
                        numero_items_rutas = numero_items_rutas + 1
                  Wend
                  txt_importe = Format(var_importe, "###,###.##")
                  rs.Close
                  n = lv_detalle.ListItems.Count
                  For i = 1 To n
                      lv_detalle.ListItems.Item(i).Selected = True
                      If lv_detalle.selectedItem.SubItems(12) = "*" Then
                         lv_detalle.selectedItem.ForeColor = &HC000&
                         lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC000&
                         lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC000&
                         lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC000&
                         lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC000&
                         lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC000&
                         lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC000&
                         lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC000&
                         lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC000&
                         lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC000&
                         lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC000&
                         lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC000&
                      Else
                         If (lv_detalle.selectedItem.SubItems(7) * 1) < (lv_detalle.selectedItem.SubItems(11) * 1) Then
                            lv_detalle.selectedItem.ForeColor = &HC0C0&
                            lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC0C0&
                            lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC0C0&
                            lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC0C0&
                            lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC0C0&
                            lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC0C0&
                            lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC0C0&
                            lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC0C0&
                            lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC0C0&
                            lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC0C0&
                            lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC0C0&
                            lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC0C0&
                         Else
                            lv_detalle.selectedItem.ForeColor = &HFF0000
                            lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF0000
                            lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF0000
                            lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF0000
                            lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF0000
                            lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF0000
                            lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF0000
                            lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF0000
                         End If
                         If (lv_detalle.selectedItem.SubItems(6) * 1) < (lv_detalle.selectedItem.SubItems(10) * 1) Then
                            lv_detalle.selectedItem.ForeColor = &HFF&
                            lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF&
                            lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF&
                            lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF&
                            lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF&
                            lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF&
                            lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF&
                            lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF&
                            lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HFF&
                            lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HFF&
                            lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HFF&
                            lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HFF&
                         End If
                      End If
                  Next i
               End If
            End If
          End If
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

Private Sub txt_cheque_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
   If KeyAscii = 27 Then
      Me.frm_cambios_relacion.Visible = False
   End If
End Sub

Private Sub txt_clave_cliente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clave_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE from tb_clientes where vcha_age_agente_id = '" + Me.txt_clave_agente + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      var_tipo_lista = 8
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

Private Sub txt_clave_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
   If KeyAscii = 27 Then
      Me.frm_cambios_relacion.Visible = False
   End If
End Sub

Private Sub txt_clave_cliente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_clave_cliente) <> "" Then
      rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + txt_clave_cliente + "' and vcha_age_agente_id = '" + Me.txt_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_cliente = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
      Else
         txt_clave_cliente = ""
         txt_nombre_cliente = ""
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_cliente = ""
   End If
End Sub

Private Sub txt_documento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
   If KeyAscii = 27 Then
      Me.frm_cambios_relacion.Visible = False
   End If
End Sub

Private Sub txt_documento_LostFocus()
   If Trim(txt_documento) <> "FA" And Trim(txt_documento) <> "NC" And Trim(txt_documento) <> "CH" Then
      txt_documento = ""
      MsgBox "Clave de documento inexistente, solo puede ser FA para facturas, CH para cheques o NC para notas de cargo", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub txt_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_clave_agente.SetFocus
   End If
End Sub

Private Sub txt_folio_LostFocus()
Dim var_posible_clientes As Boolean
Dim var_posible_facturas As Boolean
Dim var_moneda_local As Integer
Dim var_clave_agente As String
Dim var_importe As Double
Dim var_partida As Double
'On Error GoTo salir:
   If Trim(txt_folio) <> "" Then
      If Trim(var_ruta) <> "" Then
         rs.Open "select * from tb_relacion_cobranza with (nolock) where vcha_rco_folio = '" + txt_folio + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_importe = 0
            txt_clave_agente = rs!VCHA_AGE_AGENTE_ID
            rsaux2.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               txt_nombre_agente = rsaux2!VCHA_AGE_NOMBRE
            Else
               GoTo salir_agente:
            End If
            rsaux2.Close
            txt_fecha = rs!dtim_rco_fecha_relacion
            lv_detalle.ListItems.Clear
            While Not rs.EOF
               Set list_item = lv_detalle.ListItems.Add(, , rs!vcha_Cli_clave_id)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
               list_item.SubItems(2) = IIf(IsNull(rs!inte_car_numero), "", rs!inte_car_numero)
               list_item.SubItems(3) = IIf(IsNull(rs!dtim_car_fecha), "", Format(rs!dtim_car_fecha, "Short Date"))
               list_item.SubItems(4) = IIf(IsNull(rs!floa_Car_importe), 0, Round(rs!floa_Car_importe, 2))
               rsaux2.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id) + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  list_item.SubItems(5) = IIf(IsNull(rsaux2!vcha_mon_nombre_plural), "", Trim(rsaux2!vcha_mon_nombre_plural))
               End If
               rsaux2.Close
               rsaux2.Open "select floa_sal_importe from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_car_numero = " + Str(rs!inte_car_numero) + " and vcha_ser_Serie_id = '" + rs!vcha_ser_Serie_id + "' AND VCHA_CAR_DOCUMENTO = '" + rs!vcha_Car_documento + "'", cnn, adOpenDynamic, adLockOptimistic
               list_item.SubItems(6) = Format(IIf(IsNull(rsaux2!FLOA_sAL_IMPORTE), 0, rsaux2!FLOA_sAL_IMPORTE), "###,##0.00")
               rsaux2.Close
               list_item.SubItems(7) = IIf(IsNull(rs!floa_Car_descuento_aplicado), 0, rs!floa_Car_descuento_aplicado)
               list_item.SubItems(8) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
               list_item.SubItems(9) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
               list_item.SubItems(10) = Format(IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe), "###,##0.00")
               list_item.SubItems(11) = IIf(IsNull(rs!FLOA_RCO_DESCUENTO_OTORGADO), 0, rs!FLOA_RCO_DESCUENTO_OTORGADO)
               list_item.SubItems(12) = IIf(IsNull(rs!char_rco_aplicada), "", rs!char_rco_aplicada)
               list_item.SubItems(14) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
               list_item.SubItems(15) = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
               list_item.SubItems(16) = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
               list_item.SubItems(17) = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
               list_item.SubItems(18) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
               list_item.SubItems(19) = IIf(IsNull(rs!floa_rco_iva), 0, rs!floa_rco_iva)
               list_item.SubItems(20) = IIf(IsNull(rs!floa_rco_impuesto_2), 0, rs!floa_rco_impuesto_2)
               list_item.SubItems(21) = IIf(IsNull(rs!floa_rco_impuesto_3), 0, rs!floa_rco_impuesto_3)
               list_item.SubItems(22) = IIf(IsNull(rs!floa_Rco_descuento_aplicar), 0, rs!floa_Rco_descuento_aplicar)
               list_item.SubItems(23) = IIf(IsNull(rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO), 0, rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO)
               list_item.SubItems(24) = IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id)
               var_importe = var_importe + IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe)
               rs.MoveNext:
               numero_items_rutas = numero_items_rutas + 1
            Wend
            txt_importe = Format(var_importe, "###,###.##")
            rs.Close
            n = lv_detalle.ListItems.Count
            For i = 1 To n
               lv_detalle.ListItems.Item(i).Selected = True
               If lv_detalle.selectedItem.SubItems(12) = "*" Then
                  lv_detalle.selectedItem.ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC000&
                  lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC000&
               Else
                  If (lv_detalle.selectedItem.SubItems(7) * 1) < (lv_detalle.selectedItem.SubItems(11) * 1) Then
                     lv_detalle.selectedItem.ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC0C0&
                     lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC0C0&
                  Else
                     lv_detalle.selectedItem.ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF0000
                     lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF0000
                  End If
                  If (lv_detalle.selectedItem.SubItems(6) * 1) < (lv_detalle.selectedItem.SubItems(10) * 1) Then
                     lv_detalle.selectedItem.ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HFF&
                     lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HFF&
                  End If
               End If
            Next i
         Else
            If Trim(var_ruta) <> "" Then
               If var_tabla.State = 1 Then
                  var_tabla.Close
               End If
               var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
            End If
            rs.Close
            rsaux2.Open "select distinct cve_client,factura from " + var_ruta + "\" + Trim(txt_folio), var_tabla, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               var_posible_clientes = True
               var_posible_facturas = True
               var_posible_tipo_cambio = True
               While Not rsaux2.EOF
                  rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + Trim(rsaux2!cve_client) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If rs.EOF Then
                     var_posible_clientes = False
                  End If
                  rs.Close
                  rs.Open "select * from tb_encabezado_Cartera where vcha_cli_clave_id = '" + Trim(rsaux2!cve_client) + "' and inte_car_numero = " + Trim(CStr(rsaux2!FACTURA)), cnn, adOpenDynamic, adLockOptimistic
                  If rs.EOF Then
                     var_posible_facturas = False
                  End If
                  rs.Close
                  rsaux2.MoveNext
               Wend
            End If
            rsaux2.Close
            If var_posible_clientes = True Then
               If var_posible_facturas = True Then
                  rsaux2.Open "select * from " + var_ruta + "\" + Trim(txt_folio), var_tabla, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     cnn.BeginTrans
                     rs.Open "select a.vcha_age_agente_id from tb_agentes a, tb_rutas b where b.vcha_rut_ruta_id = '" + Trim(CStr(rsaux2!cve_agente)) + "' and a.vcha_age_agente_id = b.vcha_age_agente_id", cnn, adOpenDynamic, adLockOptimistic
                     var_clave_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                     rs.Close
                     rs.Open "select vcha_mon_moneda_id from tb_clientes where vcha_cli_clave_id = '" + Trim(rsaux2!cve_client) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                     End If
                     rs.Close
                     While Not rsaux2.EOF
                        rsaux3.Open "INSERT INTO TB_RELACION_COBRANZA([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_RCO_FOLIO], [DTIM_RCO_FECHA_RELACION], [VCHA_AGE_AGENTE_ID], [VCHA_CLI_CLAVE_ID], [VCHA_RCO_CHEQUE],[DTIM_RCO_FECHA_CHEQUE], [FLOA_RCO_IMPORTE],[FLOA_RCO_DESCUENTO_OTORGADO], [INTE_CAR_NUMERO], [FLOA_CAR_IMPORTE], [FLOA_CAR_DESCUENTO_APLICADO],[INTE_RCO_PARTIDA], [INTE_RCO_DESCUENTO_APLICADO],[VCHA_SER_SERIE_ID],[VCHA_CAR_DOCUMENTO],[VCHA_BAN_BANCO_ID]) values ( '" + var_empresa + "', '" + var_unidad_organizacional + "', '" + txt_folio + "', '" + CStr(Date) + "', '" + var_clave_agente + "', '" + Trim(CStr(rsaux2!cve_client)) + "', '" + CStr(rsaux2!cheque) + "', '" + CStr(rsaux2!fecha_cheq) + "', " + CStr(rsaux2!importe2) + ", " + CStr(rsaux2!descuento) + ", " + Trim(CStr(rsaux2!FACTURA)) + ", 0, 0, " + Str(var_partida) + ", 0, '" + Trim(rsaux2!Serie) + "', '" + rsaux2!tipo + "', '" + rsaux2!cvecuenta + "')", cnn, adOpenDynamic, adLockOptimistic
                        rsaux2.MoveNext
                     Wend
                     rsaux2.Close
                     cnn.CommitTrans
                     rs.Open "select * from tb_relacion_cobranza with (nolock) where vcha_rco_folio = '" + txt_folio + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_importe = 0
                        txt_clave_agente = rs!VCHA_AGE_AGENTE_ID
                        rsaux2.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                        txt_nombre_agente = rsaux2!VCHA_AGE_NOMBRE
                        rsaux2.Close
                        txt_fecha = rs!dtim_rco_fecha_relacion
                        lv_detalle.ListItems.Clear
                        While Not rs.EOF
                           Set list_item = lv_detalle.ListItems.Add(, , rs!vcha_Cli_clave_id)
                           list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
                           list_item.SubItems(2) = IIf(IsNull(rs!inte_car_numero), "", rs!inte_car_numero)
                           list_item.SubItems(3) = IIf(IsNull(rs!dtim_car_fecha), "", Format(rs!dtim_car_fecha, "Short Date"))
                           list_item.SubItems(4) = IIf(IsNull(rs!floa_Car_importe), 0, Round(rs!floa_Car_importe, 2))
                           rsaux2.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux2.EOF Then
                              list_item.SubItems(5) = IIf(IsNull(rsaux2!vcha_mon_nombre_plural), "", Trim(rsaux2!vcha_mon_nombre_plural))
                           End If
                           rsaux2.Close
                           rsaux2.Open "select floa_sal_importe from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_car_numero = " + Str(rs!inte_car_numero) + " AND VCHA_SER_SERIE_ID = '" + rs!vcha_ser_Serie_id + "' AND VCHA_CAR_DOCUMENTO = '" + rs!vcha_Car_documento + "'", cnn, adOpenDynamic, adLockOptimistic
                           list_item.SubItems(6) = Format(IIf(IsNull(rsaux2!FLOA_sAL_IMPORTE), 0, rsaux2!FLOA_sAL_IMPORTE), "###,##0.00")
                           rsaux2.Close
                           list_item.SubItems(7) = IIf(IsNull(rs!floa_Car_descuento_aplicado), 0, rs!floa_Car_descuento_aplicado)
                           list_item.SubItems(8) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
                           list_item.SubItems(9) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
                           list_item.SubItems(10) = IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe)
                           list_item.SubItems(11) = IIf(IsNull(rs!FLOA_RCO_DESCUENTO_OTORGADO), 0, rs!FLOA_RCO_DESCUENTO_OTORGADO)
                           list_item.SubItems(12) = IIf(IsNull(rs!char_rco_aplicada), "", rs!char_rco_aplicada)
                           list_item.SubItems(14) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
                           list_item.SubItems(15) = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
                           list_item.SubItems(16) = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
                           list_item.SubItems(17) = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
                           list_item.SubItems(18) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                           list_item.SubItems(19) = IIf(IsNull(rs!floa_rco_iva), 0, rs!floa_rco_iva)
                           list_item.SubItems(20) = IIf(IsNull(rs!floa_rco_impuesto_2), 0, rs!floa_rco_impuesto_2)
                           list_item.SubItems(21) = IIf(IsNull(rs!floa_rco_impuesto_3), 0, rs!floa_rco_impuesto_3)
                           list_item.SubItems(22) = IIf(IsNull(rs!floa_Rco_descuento_aplicar), 0, rs!floa_Rco_descuento_aplicar)
                           list_item.SubItems(23) = IIf(IsNull(rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO), 0, rs!INTE_RCO_NUMERO_DESCUENTO_FINANCIERO)
                           list_item.SubItems(24) = IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id)
                           var_importe = var_importe + IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe)
                           rs.MoveNext:
                           numero_items_rutas = numero_items_rutas + 1
                        Wend
                        txt_importe = Format(var_importe, "###,###.##")
                        n = lv_detalle.ListItems.Count
                        For i = 1 To n
                           lv_detalle.ListItems.Item(i).Selected = True
                           If (lv_detalle.selectedItem.SubItems(7) * 1) < (lv_detalle.selectedItem.SubItems(11) * 1) Then
                              lv_detalle.selectedItem.ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HC0C0&
                              lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HC0C0&
                           Else
                              lv_detalle.selectedItem.ForeColor = &HFF0000
                              lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF0000
                              lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF0000
                              lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF0000
                              lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF0000
                              lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF0000
                              lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF0000
                              lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF0000
                           End If
                           If (lv_detalle.selectedItem.SubItems(6) * 1) < (lv_detalle.selectedItem.SubItems(10) * 1) Then
                              lv_detalle.selectedItem.ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(1).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(2).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(3).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(4).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(5).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(6).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(7).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(8).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(9).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(10).ForeColor = &HFF&
                              lv_detalle.selectedItem.ListSubItems.Item(11).ForeColor = &HFF&
                           End If
                        Next i
                     End If
                     rs.Close
                  Else
                     rsaux2.Close
                     MsgBox "La relación de cobranza número " + txt_folio + " no existe", vbOKOnly, "ATENCION"
                  End If
                  var_tabla.Close
               Else
                  MsgBox "Existen inconcistencias en las factura, favor de revisar el archivo " + txt_folio, vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Existen inconcistencias con los clientes, favor de revisar el archivo " + txt_folio, vbOKOnly, "ATENCION"
            End If
         End If
      Else
         MsgBox "No se a indicado una ruta en donde se encuentren los archivos de cobranza", vbOKOnly, "ATENCION"
      End If
   End If
Exit Sub
On Error GoTo salir:
salir:
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   MsgBox "La relación de cobranza número " + txt_folio + " no existe", vbOKOnly, "ATENCION"
   Exit Sub
salir_agente:
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   MsgBox "La clave del agente es incorrecta", vbOKOnly, "ATENCION"
   Exit Sub
End Sub

Private Sub txt_nombre_cliente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE from tb_clientes where vcha_age_agente_id = '" + Me.txt_clave_agente + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      var_tipo_lista = 8
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

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 And KeyAscii <> 27 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
   If KeyAscii = 27 Then
      Me.frm_cambios_relacion.Visible = False
   End If
End Sub

Private Sub txt_nombre_cliente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_aceptar_cambios.SetFocus
   End If
End Sub
