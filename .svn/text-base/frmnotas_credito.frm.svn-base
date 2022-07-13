VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmnotas_credito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas de Crédito"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11640
   Begin VB.CommandButton cmd_nota_credito_electronica 
      Appearance      =   0  'Flat
      Caption         =   "NC Electronica"
      Enabled         =   0   'False
      Height          =   315
      Left            =   735
      Picture         =   "frmnotas_credito.frx":0000
      TabIndex        =   46
      Top             =   30
      Width           =   1485
   End
   Begin VB.Frame frm_lista 
      Height          =   2895
      Left            =   735
      TabIndex        =   41
      Top             =   255
      Width           =   4050
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2430
         Left            =   30
         TabIndex        =   43
         Top             =   405
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   4286
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
         TabIndex        =   42
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   " Movimiento "
      Height          =   660
      Left            =   75
      TabIndex        =   38
      Top             =   435
      Width           =   4995
      Begin VB.TextBox txt_nombre_clase 
         Height          =   315
         Left            =   885
         TabIndex        =   40
         Top             =   240
         Width           =   4035
      End
      Begin VB.TextBox txt_clase 
         Height          =   315
         Left            =   75
         TabIndex        =   39
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmnotas_credito.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Actualizar Alt + A"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11160
      Picture         =   "frmnotas_credito.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmnotas_credito.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Datos Gemerales "
      Height          =   3270
      Left            =   5175
      TabIndex        =   5
      Top             =   435
      Width           =   6435
      Begin VB.TextBox txt_falta_aplicar 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   2820
         Width           =   1400
      End
      Begin VB.ComboBox cmb_series 
         Height          =   315
         Left            =   1290
         TabIndex        =   36
         Top             =   2820
         Width           =   885
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2490
         Width           =   2145
      End
      Begin VB.TextBox txt_nombre_movimiento 
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   180
         Width           =   3840
      End
      Begin VB.TextBox txt_nombre_establecimiento 
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1830
         Width           =   3840
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1500
         Width           =   3840
      End
      Begin VB.TextBox txt_nombre_titular 
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1170
         Width           =   3840
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   840
         Width           =   3840
      End
      Begin VB.TextBox txt_rfc 
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2160
         Width           =   2145
      End
      Begin VB.TextBox txt_establecimiento 
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1830
         Width           =   1155
      End
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1500
         Width           =   1155
      End
      Begin VB.TextBox txt_titular 
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1170
         Width           =   1155
      End
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   1155
      End
      Begin VB.TextBox txt_movimiento 
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   180
         Width           =   1155
      End
      Begin VB.TextBox txt_numero 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   510
         Width           =   1155
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Falta Por Aplicar:"
         Height          =   195
         Left            =   3675
         TabIndex        =   45
         Top             =   2880
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   37
         Top             =   2880
         Width           =   405
      End
      Begin VB.Label lbl_moneda 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3570
         TabIndex        =   32
         Top             =   2505
         Width           =   2760
      End
      Begin VB.Label lbl_importe 
         AutoSize        =   -1  'True
         Caption         =   "Importe Dev.:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   2550
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "RFC:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   2205
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1890
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1230
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   900
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   570
         Width           =   600
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Devoluciones Terminadas "
      Height          =   2595
      Left            =   75
      TabIndex        =   3
      Top             =   1110
      Width           =   4995
      Begin MSComctlLib.ListView lv_devoluciones 
         Height          =   2220
         Left            =   45
         TabIndex        =   4
         Top             =   285
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   3916
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
            Text            =   "Clave"
            Object.Width           =   1464
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5380
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Número"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "clave agente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "nombre agente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "clave titular"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "nombre titular "
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "clave cliente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "nombre cliente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "clave establecimiento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "nombre establecimiento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "rfc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "almacen"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Grupo Actual"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Grupo Real"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Referencia"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   90
      Left            =   15
      TabIndex        =   2
      Top             =   300
      Width           =   11520
   End
   Begin VB.Frame Frame1 
      Caption         =   " Cargos "
      Height          =   3570
      Left            =   75
      TabIndex        =   0
      Top             =   3720
      Width           =   11505
      Begin VB.Frame frm_importe_aplicar 
         Height          =   1065
         Left            =   3780
         TabIndex        =   33
         Top             =   105
         Width           =   2160
         Begin VB.TextBox txt_importe_aplicar 
            Height          =   360
            Left            =   90
            TabIndex        =   34
            Top             =   495
            Width           =   1950
         End
         Begin VB.Label lbl_importe_aplicar 
            BackColor       =   &H8000000D&
            Caption         =   " Importe a Aplicar"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   30
            TabIndex        =   35
            Top             =   120
            Width           =   2085
         End
      End
      Begin VB.TextBox txt_total_neto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10455
         TabIndex        =   28
         Top             =   4590
         Width           =   1200
      End
      Begin MSComctlLib.ListView lv_facturas 
         Height          =   3300
         Left            =   75
         TabIndex        =   1
         Top             =   210
         Width           =   11340
         _ExtentX        =   20003
         _ExtentY        =   5821
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Número"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "     Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   " Plazo    "
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Fecha Vencim.   "
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Importe       "
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Abonos Ant."
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Abono Actual     "
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Saldo       "
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Serie"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Cliente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Iva"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Total Neto:"
         Height          =   195
         Left            =   9555
         TabIndex        =   29
         Top             =   4650
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmnotas_credito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_cantidad As Double
Dim var_precio As Double
Dim var_imp_descuent
Dim var_imp_descuento_1 As Double
Dim var_imp_descuento_2 As Double
Dim var_imp_descuento_3 As Double
Dim var_descuento_1 As Double
Dim var_descuento_2 As Double
Dim var_descuento_3 As Double
Dim var_sub_importe As Double
Dim var_iva As Double
Dim var_imp_iva As Double
Dim var_total As Double
Dim var_text_descuento As String
Dim var_total_neto As Double
Dim var_importe_total As Double
Dim var_importe_total_iva As Double
Dim var_importe_total_descuento_1 As Double
Dim var_importe_total_descuento_2 As Double
Dim var_importe_total_descuento_3 As Double
Dim var_importe_total_subimporte As Double
Dim var_tipo_Cambio As Double
Dim var_contador_encabezado As Double
Dim var_clave_moneda As String
Dim var_serie As String
Dim var_clase_documento As String


Private Sub encabezado()
   Dim n As Integer
   n = lv_devoluciones.ListItems.Count
   If n > 0 Then
      txt_movimiento = lv_devoluciones.selectedItem
      txt_nombre_movimiento = lv_devoluciones.selectedItem.SubItems(1)
      txt_agente = lv_devoluciones.selectedItem.SubItems(3)
      txt_nombre_agente = lv_devoluciones.selectedItem.SubItems(4)
      txt_titular = lv_devoluciones.selectedItem.SubItems(5)
      txt_nombre_titular = lv_devoluciones.selectedItem.SubItems(6)
      txt_cliente = lv_devoluciones.selectedItem.SubItems(7)
      txt_nombre_cliente = lv_devoluciones.selectedItem.SubItems(8)
      txt_establecimiento = lv_devoluciones.selectedItem.SubItems(9)
      txt_nombre_establecimiento = lv_devoluciones.selectedItem.SubItems(10)
      txt_rfc = lv_devoluciones.selectedItem.SubItems(11)
      txt_numero = lv_devoluciones.selectedItem.SubItems(2)
   End If
End Sub


Private Sub cmb_series_Click()
   var_serie = cmb_serie
End Sub

Private Sub cmd_guardar_Click()
   txt_numero = ""
   lv_devoluciones.ListItems.Clear
   lv_facturas.ListItems.Clear
   var_contador_encabezado = 0
   rs.Open "select * from vw_devolucion_encabezado_nota_credito where vcha_emp_empresa_id = '" + var_empresa + "'  and vcha_mov_movimiento_id <>'CAVT'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Dim list_item As ListItem
      While Not rs.EOF
         var_contador_encabezado = var_contador_encabezado + 1
         Set list_item = lv_devoluciones.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
         list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre))
         list_item.SubItems(2) = IIf(IsNull(rs!INTE_EMO_NUMERO), 0, rs!INTE_EMO_NUMERO)
         list_item.SubItems(3) = Trim(IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID))
         list_item.SubItems(4) = Trim(IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE))
         list_item.SubItems(5) = Trim(IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id))
         list_item.SubItems(6) = Trim(IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE))
         list_item.SubItems(7) = Trim(IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id))
         list_item.SubItems(8) = Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
         list_item.SubItems(9) = Trim(IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id))
         list_item.SubItems(10) = Trim(IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE))
         list_item.SubItems(11) = Trim(IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC))
         list_item.SubItems(12) = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
         list_item.SubItems(13) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
         list_item.SubItems(14) = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
         list_item.SubItems(15) = IIf(IsNull(rs!vcha_cde_referencia), "", rs!vcha_cde_referencia)
         rs.MoveNext
      Wend
      rs.Close
      Call encabezado
      If var_contador_encabezado > 9 Then
         lv_devoluciones.ColumnHeaders.item(2).Width = 2800
      Else
         lv_devoluciones.ColumnHeaders.item(2).Width = 3050.07
      End If
   Else
      rs.Close
   End If
   lv_devoluciones.SetFocus
End Sub

Private Sub cmd_imprimir_Click()
   Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
   Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
   Dim var_numero_nota As Double
   Dim var_almacen As String
   Dim var_movimiento As String
   Dim var_numero_movimiento As Double
   Dim var_grupo_actual As String
   Dim var_grupo_real As String
   Dim var_agente As String
   Dim var_titular As String
   Dim var_cliente As String
   Dim var_establecimiento As String
   Dim var_suma_importe As Double
   Dim var_total_neto_1 As Double
   Dim var_diferencia As Double
   Dim n As Double
   Dim i As Double
   Dim var_moneda_local As Integer
   Dim var_tipo_cambio_2 As Double
   Dim var_contador_renglones As Integer
   Dim var_cantidad_total As Double
   Dim var_cantidad_total_str As String
   Dim var_costo As Double
   Dim var_importe_costo As Double
   n = lv_facturas.ListItems.Count
   var_suma_importe = 0
   Dim var_clave_cliente_nuevo As String
   Dim var_clave_cliente_anterior As String
   Dim var_posible_cliente As Boolean
   Dim var_iva_pasado As Double
   var_posible_iva = 1
   var_iva_pasado = 0
   For var_j = 1 To lv_facturas.ListItems.Count
       lv_facturas.ListItems.item(var_j).Selected = True
       If (lv_facturas.selectedItem.SubItems(7) * 1) > 0 Then
          If var_iva_pasado = 0 Then
             var_iva_pasado = CDbl(Me.lv_facturas.selectedItem.SubItems(12))
          Else
             If var_iva_pasado <> CDbl(Me.lv_facturas.selectedItem.SubItems(12)) Then
                var_posible_iva = 0
             End If
          End If
       End If
   Next var_j
   If var_posible_iva = 1 Then
      var_iva = var_iva_pasado
   
   cnn.BeginTrans
   rs.Open "select max(inte_tem_consecutivo) from tb_temp_clientes_devoluciones", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
   Else
      var_consecutivo = 1
   End If
   rs.Close
   rs.Open "insert into tb_temp_clientes_devoluciones (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
   cnn.CommitTrans
   For i = 1 To n
      lv_facturas.ListItems.item(i).Selected = True
      If (lv_facturas.selectedItem.SubItems(7) * 1) > 0 Then
         rs.Open "select * from tb_temp_clientes_devoluciones where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_Cli_clave_id = '" + lv_facturas.selectedItem.SubItems(11) + "'", cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            rsaux.Open "insert into tb_temp_clientes_devoluciones (inte_tem_consecutivo, vcha_cli_clave_id) values (" + CStr(var_consecutivo) + ",'" + lv_facturas.selectedItem.SubItems(11) + "')", cnn, adOpenDynamic, adLockOptimistic
         End If
         rs.Close
      End If
   Next i
   For i = 1 To n
      lv_facturas.ListItems.item(i).Selected = True
      If (lv_facturas.selectedItem.SubItems(7) * 1) > 0 Then
         var_suma_importe = var_suma_importe + (lv_facturas.selectedItem.SubItems(7) * 1)
      End If
   Next i
   
   rsaux5.Open "select vcha_cli_clave_id from tb_temp_clientes_devoluciones where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_cli_clave_id is not null", cnn, adOpenDynamic, adLockOptimistic
   var_contador_clientes = 0
   While Not rsaux5.EOF
         var_clave_cliente_nuevo = rsaux5!vcha_cli_clave_id
         var_contador_clientes = var_contador_clientes + 1
         rsaux5.MoveNext
   Wend
   rsaux5.Close

   If var_contador_clientes = 1 Then
   If lv_devoluciones.ListItems.Count > 0 Then
      If IsNumeric(txt_importe) Then
         var_total_neto_1 = txt_importe
         var_diferencia = Round(var_suma_importe, 2) - Round(var_total_neto_1, 2)
         If var_diferencia = 0 Then
            rs.Open "select inte_ser_nota_credito from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_Ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_numero_nota = IIf(IsNull(rs!inte_ser_nota_credito), 1, rs!inte_ser_nota_credito)
            Else
               var_numero_nota = 0
            End If
            rs.Close
            If var_numero_nota > 0 Then
               si = MsgBox("¿Deseas Imprimir la Nota de Crédito " + Str(var_numero_nota) + "?", vbYesNo, "ATENCION")
               If si = 6 Then
                  si = MsgBox("Confirmación de la impresión de la nota de crédito", vbYesNo, "ATENCION")
                  If si = 6 Then
                     cnn.BeginTrans
                     var_importe_costo = 0
                     var_movimiento = lv_devoluciones.selectedItem
                     var_numero_movimiento = lv_devoluciones.selectedItem.SubItems(2)
                     var_almacen = lv_devoluciones.selectedItem.SubItems(12)
                     var_grupo_actual = lv_devoluciones.selectedItem.SubItems(13)
                     var_grupo_real = lv_devoluciones.selectedItem.SubItems(14)
                     var_titular = lv_devoluciones.selectedItem.SubItems(5)
                     var_agente = lv_devoluciones.selectedItem.SubItems(3)
                     var_cliente = var_clave_cliente_nuevo
                     var_establecimiento = lv_devoluciones.selectedItem.SubItems(9)
                     var_insertar = False
                     var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "NC", "DV", txt_clase, var_numero_nota, "-", var_almacen, var_movimiento, var_numero_movimiento, Date, var_agente, var_grupo_actual, var_grupo_real, var_titular, var_cliente, var_establecimiento, 0, var_iva, 0, 0, 0, 0, 0, var_importe_total * var_tipo_Cambio, var_importe_total_iva * var_tipo_Cambio, 0, 0, 0, 0, 0, var_importe_total_subimporte * var_tipo_Cambio, var_total_neto * var_tipo_Cambio, "", var_clave_usuario_global, "", Date, 0, Date, Date, var_clave_moneda, var_tipo_Cambio, var_serie, "")
                     rsaux3.Open "insert into tb_secuencia_notas_credito (vcha_emp_empresa_id, vcha_Ser_serie_id, inte_snc_numero_anterior, inte_snc_numero_actual) values ('" + var_empresa + "', '" + var_serie + "', " + CStr(var_numero_nota) + ", " + CStr(var_numero_nota) + ")", cnn, adOpenDynamic, adLockOptimistic
                     Set TB_DEVOLUCIONES_ESTATUS = New TB_DEVOLUCIONES_ESTATUS
                     var_estatus = "I"
                     var_modifica = False
                     var_modifica = TB_DEVOLUCIONES_ESTATUS.Anadir(var_empresa, var_unidad_organizacional, var_almacen, var_movimiento, CInt(var_numero_movimiento), "N")
                     n = lv_facturas.ListItems.Count
                     For i = 1 To n
                        lv_facturas.ListItems.item(i).Selected = True
                        If (lv_facturas.selectedItem.SubItems(7) * 1) > 0 Then
                           var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, lv_facturas.selectedItem.SubItems(10), lv_facturas.selectedItem, lv_facturas.selectedItem.SubItems(1), var_serie, "DV", var_numero_nota, 0, (lv_facturas.selectedItem.SubItems(7) * 1) * var_tipo_Cambio)
                           rsaux9.Open "select * from tb_estado_cuenta where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ecu_serie_cargo = '" + Trim(lv_facturas.selectedItem.SubItems(10)) + "' and vcha_ecu_movimiento_cargo = '" + Trim(lv_facturas.selectedItem) + "' and inte_ecu_numero_cargo = " + Trim(CStr(CDbl(lv_facturas.selectedItem.SubItems(1)))) + " and inte_ecu_numero_abono = " + CStr(var_numero_nota) + " and vcha_ecu_movimiento_abono = 'DV' and vcha_ecu_serie_abono = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                           If rsaux9.EOF Then
                              var_cadena = "insert into tb_estado_cuenta (vcha_emp_empresa_id,vcha_uor_unidad_id, vcha_ecu_movimiento_cargo, vcha_ecu_serie_cargo, inte_ecu_numero_cargo, floa_ecu_importe_cargo, vcha_ecu_movimiento_abono, vcha_ecu_serie_abono, inte_ecu_numero_abono, floa_ecu_importe_abono) "
                              var_cadena = var_cadena + "   values ('" + var_empresa + "','" + var_unidad_organizacional + "','" + Trim(lv_facturas.selectedItem) + "','" + Trim(lv_facturas.selectedItem.SubItems(10)) + "'," + Trim(CStr(CDbl(lv_facturas.selectedItem.SubItems(1)))) + ",0,'DV','" + var_serie + "'," + CStr(var_numero_nota) + "," + CStr((lv_facturas.selectedItem.SubItems(7) * 1) * var_tipo_Cambio) + ")"
                              rsaux8.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux9.Close
                           rs.Open "insert into TB_DETALLE_DEVOLUCION_IMPORTES_ASIGNADOS (vcha_emp_empresa_id, vcha_Car_documento, vcha_ser_serie_id, inte_Car_numero, vcha_dia_documento, vcha_dia_Serie, inte_dia_numero,floa_dia_importe) values ('" + var_empresa + "','DV','" + var_serie + "', " + Str(var_numero_nota) + ", '" + lv_facturas.selectedItem + "', '" + lv_facturas.selectedItem.SubItems(10) + "'," + lv_facturas.selectedItem.SubItems(1) + ", " + CStr(lv_facturas.selectedItem.SubItems(7) * 1) + ")", cnn, adOpenDynamic, adLockOptimistic
                        End If
                     Next i
                     rsaux3.Open "update tb_Series set inte_ser_nota_credito = inte_ser_nota_credito + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_Ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                     cnn.CommitTrans
                     rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DV' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_numero_nota), cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
''' ''''''''      ''' IMPRESION DE LA NOTA DE CARGO
                        Open (App.Path & "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".txt") For Output As #1
                        Print #1, Chr(27) + Chr(64)
                        Print #1, Spc(92); Str(rs!inte_Car_numero)
                        Print #1, ""
                        Print #1, Spc(92); "FECHA: "; Format(rs!dtim_Car_fecha, "Short Date")
                        var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                        For var_j = 1 + Len(Trim(var_cliente)) To 83
                            var_cliente = var_cliente + " "
                        Next var_j
                        var_cliente = var_cliente + "AGUASCALIENTES, AGS."
                        Print #1, ""
                        Print #1, Spc(12); var_cliente
                        var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " COL.: " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre) + "  C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
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
                        var_rfc = var_rfc + "ENTRADA: " + txt_movimiento + " " + txt_numero + " RELACION: " + lv_devoluciones.selectedItem.SubItems(15)
                        var_rfc = var_rfc
                        Print #1, Spc(6); var_rfc
                        Print #1, ""
                        Print #1, ""
                        Print #1, ""
                        rsaux3.Open "select * from vw_devolucion_nota_credito where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                        var_contador_renglones = 0
                        var_cantidad_total = 0
                        While Not rsaux3.EOF
                              var_factura = CStr(rsaux3!inte_fac_factura) + IIf(IsNull(rsaux3!vcha_Ser_Serie_id), "", rsaux3!vcha_Ser_Serie_id)
                              If Len(Trim(var_factura)) < 14 Then
                                 For var_j = Len(Trim(var_factura)) To 14
                                    var_factura = var_factura + " "
                                 Next var_j
                              End If
                              var_linea = var_factura + rsaux3!VCHA_ART_ARTICULO_ID + " " + Trim(rsaux3!vcha_Art_nombre_español)
                              If Len(Trim(var_linea)) < 88 Then
                                 For var_j = Len(Trim(var_linea)) To 88
                                 var_linea = var_linea + " "
                                 Next var_j
                              End If
                              var_cantidad_str = Format(IIf(IsNull(rsaux3!Cantidad_leida), 0, rsaux3!Cantidad_leida), "###,###,##0.00")
                              var_tipo_Cambio = IIf(IsNull(rsaux3!floa_dev_tipo_cambio), 1, rsaux3!floa_dev_tipo_cambio)
                              var_cantidad = Format(IIf(IsNull(rsaux3!Cantidad), 0, rsaux3!Cantidad), "###,###,##0.00")
                              var_cantidad_total = var_cantidad_total + IIf(IsNull(rsaux3!Cantidad), 0, rsaux3!Cantidad)
                              var_precio = IIf(IsNull(rsaux3!floa_cde_precio), 0, rsaux3!floa_cde_precio) / IIf(IsNull(rsaux3!Cantidad), 0, rsaux3!Cantidad)
                              var_importe_costo = var_importe_costo + IIf(IsNull(rsaux3!floa_cde_costo), 0, rsaux3!floa_cde_costo)
                              var_descuento_1 = IIf(IsNull(rsaux3!floa_cde_descuento_1), 0, rsaux3!floa_cde_descuento_1)
                              var_descuento_2 = IIf(IsNull(rsaux3!floa_cde_descuento_2), 0, rsaux3!floa_cde_descuento_2)
                              var_descuento_3 = IIf(IsNull(rsaux3!floa_cde_descuento_3), 0, rsaux3!floa_cde_descuento_3)
                              var_tipo_Cambio = IIf(IsNull(rsaux3!floa_dev_tipo_cambio), 1, rsaux3!floa_dev_tipo_cambio)
                              var_precio = var_precio * (1 - (var_descuento_1 / 100))
                              var_precio = var_precio * (1 - (var_descuento_2 / 100))
                              var_precio = var_precio * (1 - (var_descuento_3 / 100))
                              var_precio = var_precio / var_tipo_Cambio
                     
                              var_iva = IIf(IsNull(rsaux3!floa_cde_iva), 0, rsaux3!floa_cde_iva)
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              If Len(Trim(var_rfc)) = 0 Then
                                 var_precio = var_precio * (1 + var_iva / 100)
                              End If
                              var_precio_str = Format(var_precio, "###,###,##0.00")
                              var_importe_str = Format(var_precio * var_cantidad, "###,###,##0.00")
                              If Len(Trim(var_cantidad_str)) < 14 Then
                                 For var_j = Len(Trim(var_cantidad_str)) To 14
                                     var_cantidad_str = " " + var_cantidad_str
                                 Next var_j
                              End If
                              var_linea = var_linea + var_cantidad_str
                              If Len(Trim(var_linea)) < 104 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 104
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              If Len(Trim(var_precio_str)) < 14 Then
                                 For var_j = Len(Trim(var_precio_str)) To 14
                                     var_precio_str = " " + var_precio_str
                                 Next var_j
                              End If
                              If Len(Trim(var_importe_str)) < 14 Then
                                 For var_j = Len(Trim(var_importe_str)) To 14
                                     var_importe_str = " " + var_importe_str
                                 Next var_j
                              End If
                              var_linea = var_linea + var_precio_str + var_importe_str
                              Print #1, var_linea
                              rsaux3.MoveNext
                              var_contador_renglones = var_contador_renglones + 1
                        Wend
                        var_cantidad_total_str = Format(var_cantidad_total, "###,###,##0.00")
                        If Len(Trim(var_cantidad_total_str)) < 14 Then
                           For var_j = Len(Trim(var_cantidad_total_str)) To 14
                               var_cantidad_total_str = " " + var_cantidad_total_str
                           Next var_j
                        End If
                        If var_contador_renglones < 30 Then
                           For var_j = var_contador_renglones To 30
                               Print #1, ""
                           Next var_j
                        End If
                        rsaux3.Close
                        
                        Print #1, ""
                        Print #1, ""
                        Print #1, ""
                        var_contador_renglones = 0
                        rsaux3.Open "select * from TB_DETALLE_DEVOLUCION_IMPORTES_ASIGNADOS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DV' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_numero_nota), cnn, adOpenDynamic, adLockOptimistic
                        var_linea = ""
                        While Not rsaux3.EOF
                              If Len(Trim(var_linea + ", " + CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00"))) < 108 Then
                                 If Len(Trim(var_linea)) = 0 Then
                                    var_linea = var_linea + CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00")
                                 Else
                                    var_linea = var_linea + ", " + CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00")
                                 End If
                              Else
                                 Print #1, var_linea
                                 var_contador_renglones = var_contador_renglones + 1
                                 var_linea = ""
                                 var_linea = CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00")
                              End If
                              rsaux3.MoveNext
                              If rsaux3.EOF And Len(var_linea) < 118 Then
                                 Print #1, var_linea
                                 var_contador_renglones = var_contador_renglones + 1
                              End If
                        Wend
                        If var_contador_renglones < 4 Then
                           For var_j = var_contador_renglones To 4
                               Print #1, ""
                           Next var_j
                        End If
                        rsaux3.Close
                        var_cantidad_letra = rs!vcha_car_importe_letra
                         
                        var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                     
                        If Len(Trim(var_linea)) < 62 Then
                           For var_j = 1 + Len(Trim(var_linea)) To 62
                               var_linea = var_linea + " "
                           Next var_j
                        End If
                        var_linea = var_linea + var_cantidad_total_str
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
                        If Len(Trim(var_linea)) < 115 Then
                            For var_j = Len(Trim(var_linea)) To 115
                                var_linea = var_linea + " "
                            Next var_j
                        End If
                        var_linea = var_linea + var_subimporte_str
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
                        Print #1, ""
                        Close #1
                        var_cadena = "update tb_encabezado_cartera set floa_car_costo = " + CStr(var_importe_costo) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and VCHA_CAR_TIPO_DOCUMENTO = 'NC' and  VCHA_CAR_DOCUMENTO = 'DV' and VCHA_CAR_CLASE_ID = '" + txt_clase + "'  and INTE_CAR_NUMERO = " + CStr(var_numero_nota) + " and vcha_Ser_serie_id = '" + var_serie + "'"
                        rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        Open (App.Path & "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".bat") For Output As #2
                        var_Archivo = App.Path & "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".bat"
                        Print #2, "copy " + App.Path + "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".txt lpt1"
                        Close #2
                       x = Shell(var_Archivo, vbHide)
''''''''''''  '''
                     End If
                     rs.Close
                     MsgBox "Se a terminado de generar la nota de crédito", vbOKOnly, "ATENCION"
                     txt_movimiento = ""
                     txt_nombre_movimiento = ""
                     txt_numero = ""
                     txt_agente = ""
                     txt_nombre_agente = ""
                     txt_titular = ""
                     txt_nombre_titular = ""
                     txt_cliente = ""
                     txt_nombre_cliente = ""
                     txt_establecimiento = ""
                     txt_nombre_establecimiento = ""
                     txt_rfc = ""
                     Me.lv_facturas.ListItems.Clear
                     lv_devoluciones.ListItems.Remove (lv_devoluciones.selectedItem.Index)
                     lv_devoluciones.SetFocus
                  Else
                     MsgBox "La impresión de la nota de crédito a sido cancelada", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "La impresión de la nota de crédito a sido cancelada", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Es imposible imprimir la nota de credito ya que no existe un folio", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Es imposible imprimir la nota de crédito ya que no se a asignado el total del importe", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Importe incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No existen devoluciones a aplicar", vbOKOnly, "ATENCION"
   End If
   Else
       MsgBox "Las facturas a las que se a abonado el importe de la devolución corresponden a mas de un cliente", vbOKOnly, "ATENCION"
   End If
   rsaux5.Open "delete from tb_temp_clientes_devoluciones where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   Else
      MsgBox "No puede mezclar facturas con distintos tipos de IVA", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir:
   MsgBox "No se puede imprimir la Nota de Crédito ya que no se a indicado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
End Sub

Private Sub cmd_nota_credito_electronica_Click()
   Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
   Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
   Dim var_numero_nota As Double
   Dim var_almacen As String
   Dim var_movimiento As String
   Dim var_numero_movimiento As Double
   Dim var_grupo_actual As String
   Dim var_grupo_real As String
   Dim var_agente As String
   Dim var_titular As String
   Dim var_cliente As String
   Dim var_establecimiento As String
   Dim var_suma_importe As Double
   Dim var_total_neto_1 As Double
   Dim var_diferencia As Double
   Dim n As Double
   Dim i As Double
   Dim var_moneda_local As Integer
   Dim var_tipo_cambio_2 As Double
   Dim var_contador_renglones As Integer
   Dim var_cantidad_total As Double
   Dim var_cantidad_total_str As String
   Dim var_costo As Double
   Dim var_importe_costo As Double
   n = lv_facturas.ListItems.Count
   var_suma_importe = 0
   Dim var_clave_cliente_nuevo As String
   Dim var_clave_cliente_anterior As String
   Dim var_posible_cliente As Boolean
   Dim var_iva_pasado As Double
   var_posible_iva = 1
   var_iva_pasado = 0
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
   
   For var_j = 1 To lv_facturas.ListItems.Count
       lv_facturas.ListItems.item(var_j).Selected = True
       If (lv_facturas.selectedItem.SubItems(7) * 1) > 0 Then
          If var_iva_pasado = 0 Then
             var_iva_pasado = CDbl(Me.lv_facturas.selectedItem.SubItems(12))
          Else
             If var_iva_pasado <> CDbl(Me.lv_facturas.selectedItem.SubItems(12)) Then
                var_posible_iva = 0
             End If
          End If
       End If
   Next var_j
   If var_posible_iva = 1 Then
      var_iva = var_iva_pasado
   
   cnn.BeginTrans
   rs.Open "select max(inte_tem_consecutivo) from tb_temp_clientes_devoluciones", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
   Else
      var_consecutivo = 1
   End If
   rs.Close
   rs.Open "insert into tb_temp_clientes_devoluciones (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
   cnn.CommitTrans
   For i = 1 To n
      lv_facturas.ListItems.item(i).Selected = True
      If (lv_facturas.selectedItem.SubItems(7) * 1) > 0 Then
         rs.Open "select * from tb_temp_clientes_devoluciones where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_Cli_clave_id = '" + lv_facturas.selectedItem.SubItems(11) + "'", cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open "insert into tb_temp_clientes_devoluciones (inte_tem_consecutivo, vcha_cli_clave_id) values (" + CStr(var_consecutivo) + ",'" + lv_facturas.selectedItem.SubItems(11) + "')", cnn, adOpenDynamic, adLockOptimistic
         End If
         rs.Close
      End If
   Next i
   For i = 1 To n
      lv_facturas.ListItems.item(i).Selected = True
      If (lv_facturas.selectedItem.SubItems(7) * 1) > 0 Then
         var_suma_importe = var_suma_importe + (lv_facturas.selectedItem.SubItems(7) * 1)
      End If
   Next i
   
   rsaux5.Open "select vcha_cli_clave_id from tb_temp_clientes_devoluciones where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_cli_clave_id is not null", cnn, adOpenDynamic, adLockOptimistic
   var_contador_clientes = 0
   While Not rsaux5.EOF
         var_clave_cliente_nuevo = rsaux5!vcha_cli_clave_id
         var_contador_clientes = var_contador_clientes + 1
         rsaux5.MoveNext
   Wend
   rsaux5.Close



   If var_contador_clientes = 1 Then
   If lv_devoluciones.ListItems.Count > 0 Then
      If IsNumeric(txt_importe) Then
         var_total_neto_1 = txt_importe
         var_diferencia = Round(var_suma_importe, 2) - Round(var_total_neto_1, 2)
         If var_diferencia = 0 Then
            rs.Open "select inte_ser_nota_credito from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_numero_nota = IIf(IsNull(rs!inte_ser_nota_credito), 1, rs!inte_ser_nota_credito)
            Else
               var_numero_nota = 0
            End If
            rs.Close
            If var_numero_nota > 0 Then
               si = MsgBox("¿Deseas Imprimir la Nota de Crédito " + Str(var_numero_nota) + "?", vbYesNo, "ATENCION")
               If si = 6 Then
                  si = MsgBox("Confirmación de la impresión de la nota de crédito", vbYesNo, "ATENCION")
                  If si = 6 Then
                     cnn.BeginTrans
                     var_importe_costo = 0
                     var_movimiento = lv_devoluciones.selectedItem
                     var_numero_movimiento = lv_devoluciones.selectedItem.SubItems(2)
                     var_almacen = lv_devoluciones.selectedItem.SubItems(12)
                     var_grupo_actual = lv_devoluciones.selectedItem.SubItems(13)
                     var_grupo_real = lv_devoluciones.selectedItem.SubItems(14)
                     var_titular = lv_devoluciones.selectedItem.SubItems(5)
                     var_agente = lv_devoluciones.selectedItem.SubItems(3)
                     var_cliente = var_clave_cliente_nuevo
                     var_establecimiento = lv_devoluciones.selectedItem.SubItems(9)
                     var_insertar = False
                     var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "NC", "DV", txt_clase, var_numero_nota, "-", var_almacen, var_movimiento, var_numero_movimiento, Date, var_agente, var_grupo_actual, var_grupo_real, var_titular, var_cliente, var_establecimiento, 0, var_iva, 0, 0, 0, 0, 0, var_importe_total * var_tipo_Cambio, var_importe_total_iva * var_tipo_Cambio, 0, 0, 0, 0, 0, var_importe_total_subimporte * var_tipo_Cambio, var_total_neto * var_tipo_Cambio, "", var_clave_usuario_global, "", Date, 0, Date, Date, var_clave_moneda, var_tipo_Cambio, var_serie, "")
                     rsaux3.Open "insert into tb_secuencia_notas_credito (vcha_emp_empresa_id, vcha_Ser_serie_id, inte_snc_numero_anterior, inte_snc_numero_actual) values ('" + var_empresa + "', '" + var_serie + "', " + CStr(var_numero_nota) + ", " + CStr(var_numero_nota) + ")", cnn, adOpenDynamic, adLockOptimistic
                     Set TB_DEVOLUCIONES_ESTATUS = New TB_DEVOLUCIONES_ESTATUS
                     var_estatus = "I"
                     var_modifica = False
                     If var_aplicar_nota_credito = 1 Then
                        var_modifica = TB_DEVOLUCIONES_ESTATUS.Anadir(var_empresa, var_unidad_organizacional, var_almacen, var_movimiento, CInt(var_numero_movimiento), "")
                     Else
                        var_modifica = TB_DEVOLUCIONES_ESTATUS.Anadir(var_empresa, var_unidad_organizacional, var_almacen, var_movimiento, CInt(var_numero_movimiento), "N")
                     End If
                     n = lv_facturas.ListItems.Count
                     For i = 1 To n
                        lv_facturas.ListItems.item(i).Selected = True
                        If (lv_facturas.selectedItem.SubItems(7) * 1) > 0 Then
                           var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, lv_facturas.selectedItem.SubItems(10), lv_facturas.selectedItem, lv_facturas.selectedItem.SubItems(1), var_serie, "DV", var_numero_nota, 0, (lv_facturas.selectedItem.SubItems(7) * 1) * var_tipo_Cambio)
                           rsaux9.Open "select * from tb_estado_cuenta where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ecu_serie_cargo = '" + Trim(lv_facturas.selectedItem.SubItems(10)) + "' and vcha_ecu_movimiento_cargo = '" + Trim(lv_facturas.selectedItem) + "' and inte_ecu_numero_cargo = " + Trim(CStr(CDbl(lv_facturas.selectedItem.SubItems(1)))) + " and inte_ecu_numero_abono = " + CStr(var_numero_nota) + " and vcha_ecu_movimiento_abono = 'DV' and vcha_ecu_serie_abono = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                           If rsaux9.EOF Then
                              var_cadena = "insert into tb_estado_cuenta (vcha_emp_empresa_id,vcha_uor_unidad_id, vcha_ecu_movimiento_cargo, vcha_ecu_serie_cargo, inte_ecu_numero_cargo, floa_ecu_importe_cargo, vcha_ecu_movimiento_abono, vcha_ecu_serie_abono, inte_ecu_numero_abono, floa_ecu_importe_abono) "
                              var_cadena = var_cadena + "   values ('" + var_empresa + "','" + var_unidad_organizacional + "','" + Trim(lv_facturas.selectedItem) + "','" + Trim(lv_facturas.selectedItem.SubItems(10)) + "'," + Trim(CStr(CDbl(lv_facturas.selectedItem.SubItems(1)))) + ",0,'DV','" + var_serie + "'," + CStr(var_numero_nota) + "," + CStr((lv_facturas.selectedItem.SubItems(7) * 1) * var_tipo_Cambio) + ")"
                              rsaux8.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux9.Close
                           rs.Open "insert into TB_DETALLE_DEVOLUCION_IMPORTES_ASIGNADOS (vcha_emp_empresa_id, vcha_Car_documento, vcha_ser_serie_id, inte_Car_numero, vcha_dia_documento, vcha_dia_Serie, inte_dia_numero,floa_dia_importe) values ('" + var_empresa + "','DV','" + var_serie + "', " + Str(var_numero_nota) + ", '" + lv_facturas.selectedItem + "', '" + lv_facturas.selectedItem.SubItems(10) + "'," + lv_facturas.selectedItem.SubItems(1) + ", " + CStr(lv_facturas.selectedItem.SubItems(7) * 1) + ")", cnn, adOpenDynamic, adLockOptimistic
                        End If
                     Next i
                     rsaux3.Open "update tb_Series set inte_ser_nota_credito = inte_ser_nota_credito + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                     cnn.CommitTrans
                     rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DV' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_numero_nota), cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
''' ''''''''      ''' IMPRESION DE LA NOTA DE CARGO
                        var_k = var_numero_nota_inicio
                        'Close #1
                        Open (var_ruta_documentos_electronicos & "\" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".fI") For Output As #1
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
                        
                        
                        
                        var_contador_renglones = 0
                        If rsaux3.State = 1 Then
                           rsaux3.Close
                        End If
                        rsaux3.Open "select * from TB_DETALLE_DEVOLUCION_IMPORTES_ASIGNADOS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DV' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_numero_nota), cnn, adOpenDynamic, adLockOptimistic
                        var_linea = ""
                        While Not rsaux3.EOF
                              If Len(Trim(var_linea)) = 0 Then
                                 var_linea = var_linea + CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00")
                              Else
                                 var_linea = var_linea + ", " + CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00")
                              End If
                              rsaux3.MoveNext
                        Wend
                        rsaux3.Close
                                                
                        
                        
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
                        var_cadena = var_cadena + "formaDePago=CONTADO" + Chr(13)
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
                        rsaux2.Open "select * from tb_empresa_FACTURA_ELECTRONICA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_certificado = rsaux2!vcha_emp_certificado
                        var_expedido = rsaux2!vcha_emp_expedido
                        If var_rfc_cliente = "XAXX010101000" Then
                           var_cadena = var_cadena + "iva=" + Format(CStr(0), "###,###,##0.000000") + Chr(13)
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
                        'var_cadena = var_cadena + "email=" + "fserna@vianney.com.mx" + Chr(13)
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
                        
                        
                        
                        
                        
                        
                        If rsaux3.State = 1 Then
                           rsaux3.Close
                        End If
                        
                        rsaux3.Open "select * from vw_devolucion_nota_credito where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                        
                        var_i = 1
                        var_piezas_totales = 0
                        While Not rsaux3.EOF
                             var_referencia_agente = "NA " + IIf(IsNull(rsaux3!vcha_emo_referencia), "", rsaux3!vcha_emo_referencia)
                              pxx = CStr(var_i)
                              If Len(pxx) = 1 Then
                                 pxx = "0" + pxx
                              End If
                              var_cadena = var_cadena + "p" + pxx + "_cantidad=" + CStr(IIf(IsNull(rsaux3!Cantidad_leida), 0, rsaux3!Cantidad_leida)) + Chr(13)
                              var_piezas_totales = var_piezas_totales + IIf(IsNull(rsaux3!Cantidad_leida), 0, rsaux3!Cantidad_leida)
                              If rsaux4.State = 1 Then
                                 rsaux4.Close
                              End If
                              rsaux4.Open "SELECT dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_UNIDADES.VCHA_UNI_UNIDAD_ID, dbo.TB_UNIDADES.VCHA_UNI_NOMBRE, dbo.TB_Articulos.VCHA_ART_ARTICULO_ID FROM dbo.TB_ARTICULOS LEFT OUTER JOIN dbo.TB_UNIDADES ON dbo.TB_ARTICULOS.VCHA_UNI_UNIDAD_ID = dbo.TB_UNIDADES.VCHA_UNI_UNIDAD_ID WHERE (dbo.TB_ARTICULOS.vcha_art_Articulo_id = '" + Trim(IIf(IsNull(rsaux3!VCHA_ART_ARTICULO_ID), "", rsaux3!VCHA_ART_ARTICULO_ID)) + "')", cnn, adOpenDynamic, adLockOptimistic
                              var_factura = IIf(IsNull(rsaux3!vcha_Ser_Serie_id), "", rsaux3!vcha_Ser_Serie_id) + CStr(rsaux3!inte_fac_factura)
                              var_linea = var_factura + " " + rsaux3!VCHA_ART_ARTICULO_ID + " " + Trim(rsaux4!vcha_Art_nombre_español)
                              var_cadena = var_cadena + "p" + pxx + "_unidad=" + IIf(IsNull(rsaux4!VCHA_UNI_NOMBRE), "", rsaux4!VCHA_UNI_NOMBRE) + Chr(13)
                              var_cadena = var_cadena + "p" + pxx + "_noIdentificacion=" + Chr(13)
                              
                              rsaux4.Close
                              var_cadena = var_cadena + "p" + pxx + "_descripcion=" + var_linea + Chr(13)
                              'var_importe_str = var_importe_str = Format(((IIf(IsNull(rsaux3!FLOA_dbo_IMPORTE), 0, rsaux3!FLOA_dbo_IMPORTE)) / (1 + (rsaux3!floa_dbo_iva / 100)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                              
                              var_precio = IIf(IsNull(rsaux3!floa_cde_precio), 0, rsaux3!floa_cde_precio) / IIf(IsNull(rsaux3!Cantidad), 0, rsaux3!Cantidad)
                              var_descuento_1 = IIf(IsNull(rsaux3!floa_cde_descuento_1), 0, rsaux3!floa_cde_descuento_1)
                              var_descuento_2 = IIf(IsNull(rsaux3!floa_cde_descuento_2), 0, rsaux3!floa_cde_descuento_2)
                              var_descuento_3 = IIf(IsNull(rsaux3!floa_cde_descuento_3), 0, rsaux3!floa_cde_descuento_3)
                              var_tipo_Cambio = IIf(IsNull(rsaux3!floa_dev_tipo_cambio), 1, rsaux3!floa_dev_tipo_cambio)
                              var_precio = var_precio * (1 - (var_descuento_1 / 100))
                              var_precio = var_precio * (1 - (var_descuento_2 / 100))
                              var_precio = var_precio * (1 - (var_descuento_3 / 100))
                              var_precio = var_precio / var_tipo_Cambio
                              var_iva = IIf(IsNull(rsaux3!floa_cde_iva), 0, rsaux3!floa_cde_iva)
                              If var_rfc_cliente = "XAXX010101000" Then
                                 var_precio = var_precio * (1 + (var_iva / 100))
                              End If
                              var_cadena = var_cadena + "p" + pxx + "_valorUnitario=" + Format(CStr(var_precio), "###,###,##0.000000") + Chr(13)
                              var_cadena = var_cadena + "p" + pxx + "_importe=" + Format(CStr(var_precio * IIf(IsNull(rsaux3!Cantidad_leida), 0, rsaux3!Cantidad_leida)), "###,###,##0.000000") + Chr(13)
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
                        var_cadena = var_cadena + "pedido=" + var_referencia_agente + Chr(13)
                        var_cadena = var_cadena + "Embarque=CA " + Me.txt_numero + Chr(13)
                        var_referencia_Bancaria = ""
                        var_cadena = var_cadena + "referenciabancaria=" + Chr(13)
                        var_cadena = var_cadena + "fechaPedido=" + Chr(13)
                        var_cadena = var_cadena + "expedicion=" + Chr(13)
                        
                        If rsaux3.State = 1 Then
                           rsaux3.Close
                        End If
                         rsaux3.Open "select * from TB_DETALLE_DEVOLUCION_IMPORTES_ASIGNADOS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DV' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_numero_nota), cnn, adOpenDynamic, adLockOptimistic
                         var_linea = ""
                         While Not rsaux3.EOF
                               If Len(Trim(var_linea)) = 0 Then
                                  var_linea = var_linea + CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00")
                               Else
                                  var_linea = var_linea + ", " + CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00")
                               End If
                               rsaux3.MoveNext
                         Wend
                         rsaux3.Close
                         var_linea_importes_asignados = var_linea
                        
                        
                        
                        var_cadena = var_cadena + "observaciones=" + var_linea_importes_asignados + Chr(13)
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
                        
                        If var_empresa = "02" Or var_empresa = "03" Or var_empresa = "18" Or var_empresa = "17" Or var_empresa = "06" Then
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
                        If var_empresa = "16" Then
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
                        var_cadena = var_cadena + "piezas_totales=" + CStr(var_piezas_totales) + Chr(13)
                        var_cadena = var_cadena + "<addenda>" + Chr(13)
                        var_cadena = var_cadena + "</addenda>" + Chr(13) + Chr(13)
                        var_cadena = var_cadena + "</Factura>"
                        Print #1, var_cadena
                        Close #1
                        rsaux.Open "insert into tb_notas_credito_devoluciones (vcha_emp_empresa_id,  vcha_uor_unidad_id, vcha_mov_movimiento_id, inte_emo_numero ) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + Me.lv_devoluciones.selectedItem + "'," + CStr(lv_devoluciones.selectedItem.SubItems(2)) + ")", cnn, adOpenDynamic, adLockOptimistic
                        var_Archivo = App.Path & "\renombra" + var_serie + Trim(Str(rs!inte_Car_numero)) + ".bat"
                        Open (App.Path & "\renombra" + var_serie + Trim(Str(rs!inte_Car_numero)) + ".bat") For Output As #2
                        Print #2, "ren " + var_ruta_documentos_electronicos + "\" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".fi " + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".ff"
                        Close #2
            
                         x = Shell(var_Archivo, vbHide)
                        
                        
''''''''''''  '''
                     End If
                     rs.Close
                     MsgBox "Se a terminado de generar la nota de crédito", vbOKOnly, "ATENCION"
                     txt_movimiento = ""
                     txt_nombre_movimiento = ""
                     txt_numero = ""
                     txt_agente = ""
                     txt_nombre_agente = ""
                     txt_titular = ""
                     txt_nombre_titular = ""
                     txt_cliente = ""
                     txt_nombre_cliente = ""
                     txt_establecimiento = ""
                     txt_nombre_establecimiento = ""
                     txt_rfc = ""
                     Me.lv_facturas.ListItems.Clear
                     lv_devoluciones.ListItems.Remove (lv_devoluciones.selectedItem.Index)
                     lv_devoluciones.SetFocus
                  Else
                     MsgBox "La impresión de la nota de crédito a sido cancelada", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "La impresión de la nota de crédito a sido cancelada", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Es imposible imprimir la nota de credito ya que no existe un folio", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Es imposible imprimir la nota de crédito ya que no se a asignado el total del importe", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Importe incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No existen devoluciones a aplicar", vbOKOnly, "ATENCION"
   End If
   Else
       MsgBox "Las facturas a las que se a abonado el importe de la devolución corresponden a mas de un cliente", vbOKOnly, "ATENCION"
   End If
   rsaux5.Open "delete from tb_temp_clientes_devoluciones where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   Else
      MsgBox "No puede mezclar facturas con distintos tipos de IVA", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir:
   MsgBox "No se puede imprimir la Nota de Crédito ya que no se a indicado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 0
   frm_lista.Visible = False
   var_contador_encabezado = 0
   frm_importe_aplicar.Visible = False
   cnn.CommandTimeout = 360
   var_cadena = "SELECT DISTINCT VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, INTE_EMO_NUMERO, VCHA_CDE_REFERENCIA, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_TIT_TITULAR_ID, VCHA_TIT_NOMBRE, VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE, VCHA_GRE_GRUPO_REAL_ID, VCHA_GRE_NOMBRE, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GAC_NOMBRE, VCHA_MOV_MOVIMIENTO_ID, VCHA_MOV_NOMBRE, VCHA_CLI_RFC, VCHA_ESB_ESTABLECIMIENTO_ID , VCHA_ESB_NOMBRE, CHAR_CDE_ESTATUS From dbo.VW_DEVOLUCION_NOTA_CREDITO where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id <>'CAVT' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento_nc + "' AND INTE_EMO_NUMERO  = " + CStr(var_numero_nc)
   
   If var_aplicar_nota_credito = 1 Then
      'rs.Open "select * from vw_devolucion_encabezado_nota_credito where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id <>'CAVT' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento_nc + "' AND INTE_EMO_NUMERO  = " + CStr(var_numero_nc), cnn, adOpenDynamic, adLockOptimistic
      rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
   Else
      rs.Open "select * from vw_devolucion_encabezado_nota_credito where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id <>'CAVT' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   If Not rs.EOF Then
      Dim list_item As ListItem
      While Not rs.EOF
         var_contador_encabezado = var_contador_encabezado + 1
         Set list_item = lv_devoluciones.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
         list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre))
         list_item.SubItems(2) = IIf(IsNull(rs!INTE_EMO_NUMERO), 0, rs!INTE_EMO_NUMERO)
         list_item.SubItems(3) = Trim(IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID))
         list_item.SubItems(4) = Trim(IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE))
         list_item.SubItems(5) = Trim(IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id))
         list_item.SubItems(6) = Trim(IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE))
         list_item.SubItems(7) = Trim(IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id))
         list_item.SubItems(8) = Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
         list_item.SubItems(9) = Trim(IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id))
         list_item.SubItems(10) = Trim(IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE))
         list_item.SubItems(11) = Trim(IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC))
         list_item.SubItems(12) = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
         list_item.SubItems(13) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
         list_item.SubItems(14) = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
         list_item.SubItems(15) = IIf(IsNull(rs!vcha_cde_referencia), "", rs!vcha_cde_referencia)
         rs.MoveNext
      Wend
      rs.Close
      'Call encabezado
      If var_contador_encabezado > 9 Then
         lv_devoluciones.ColumnHeaders.item(2).Width = 2800
      Else
         lv_devoluciones.ColumnHeaders.item(2).Width = 3050.07
      End If
      txt_numero = ""
      Call encabezado
   Else
      rs.Close
   End If
   
   rs.Open "select vcha_ser_serie_id from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_contador_serie = 0
      While Not rs.EOF
         var_contador_serie = var_contador_serie + 1
         rs.MoveNext
      Wend
      rs.MoveFirst
      Call RecsetToCombo(cmb_series.hwnd, rs, 0)
      If var_contador_serie > 1 Then
         cmb_series.Enabled = True
      Else
         cmb_series.Enabled = False
      End If
      rs.MoveFirst
      cmb_series = rs!vcha_Ser_Serie_id
      var_serie = rs!vcha_Ser_Serie_id
   Else
      MsgBox "No se a indicado una serie para esta Unidad organizacional", vbOKOnly, "ATENCION"
   End If
   rs.Close
   rs.Open "select * from tb_clases_Cartera where vcha_car_documento = 'DV' order by vcha_car_nombre ", cnn, adOpenDynamic, adLockBatchOptimistic
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
      txt_nombre_clase = rs!vcha_Car_nombre
      txt_clase = rs!vcha_Car_clase_id
   Else
      MsgBox "No se a indicado una clase de Devolución", vbOKOnly, "ATENCION"
      txt_clase.Enabled = False
      cmb_clases.Enabled = False
   End If
   rs.Close
   Me.cmd_nota_credito_electronica.Enabled = True
   Me.cmd_imprimir.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_notas_credito)
End Sub

Private Sub lst_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_devoluciones_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_devoluciones, ColumnHeader)
End Sub

Private Sub lv_devoluciones_GotFocus()
   If lv_devoluciones.ListItems.Count > 0 Then
      Call encabezado
   End If
End Sub

Private Sub lv_devoluciones_ItemClick(ByVal item As MSComctlLib.ListItem)
   If lv_devoluciones.ListItems.Count > 0 Then
      Call encabezado
   End If
End Sub


Private Sub lv_facturas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_facturas, ColumnHeader)
End Sub

Private Sub lv_facturas_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F4 para indicar el importe a aplicar a la factura"
End Sub

Private Sub lv_facturas_ItemClick(ByVal item As MSComctlLib.ListItem)
   Frmmenu2.StatusBar1.Panels(1) = "Presione F4 para indicar el importe a aplicar a la factura"
End Sub

Private Sub lv_facturas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 115 Then
      If Me.lv_facturas.ListItems.Count > 0 Then
         txt_importe_aplicar = ""
         frm_importe_aplicar.Visible = True
         txt_importe_aplicar.SetFocus
      Else
         MsgBox "No existen facturas", vbOKOnly, "ATENCION"
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

Private Sub txt_clase_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clase_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      Dim list_item As ListItem
      Dim var_contador_lista As Integer
      rs.Open "select vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where vcha_car_documento= 'DV' order by vcha_Car_nombre", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            var_contador_lista = var_contador_lista + 1
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Car_clase_id)
            list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_Car_nombre), "", rs!vcha_Car_nombre))
            rs.MoveNext
         Wend
      End If
      rs.Close
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clase_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_clase_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_importe_aplicar_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Dim var_importe_aplicar As Double
      Dim var_importe_total As Double
      Dim var_item_seleccionado As Integer
      Dim var_importe_saldo As Double
      Dim var_importe_total_l2 As Double
      Dim n As Integer
      Dim i As Integer
      If IsNumeric(txt_importe) Then
         If IsNumeric(txt_importe_aplicar) Then
            var_importe_aplicar = txt_importe_aplicar
            var_item_seleccionado = lv_facturas.selectedItem.Index
            n = lv_facturas.ListItems.Count
            var_importe_total = 0
            For i = 1 To n
               lv_facturas.ListItems.item(i).Selected = True
               var_importe_total = var_importe_total + (lv_facturas.selectedItem.SubItems(7) * 1)
            Next i
            lv_facturas.ListItems.item(var_item_seleccionado).Selected = True
            var_importe_saldo = lv_facturas.selectedItem.SubItems(8) * 1
            var_suma_importe_totaL = (var_importe_total + var_importe_aplicar)
            var_importe_total_l2 = CDbl(Me.txt_importe) * 1
            If Round(var_suma_importe_totaL, 2) <= Round(var_importe_total_l2, 2) Then
              If Round(var_importe_saldo * 1, 2) >= Round(var_importe_aplicar * 1, 2) Then
                  lv_facturas.selectedItem.SubItems(7) = Format((lv_facturas.selectedItem.SubItems(7) * 1) + var_importe_aplicar, "###,###,##0.00")
                  lv_facturas.selectedItem.SubItems(8) = Format((lv_facturas.selectedItem.SubItems(8) * 1) - var_importe_aplicar, "###,###,##0.00")
                  txt_falta_aplicar = Format((CDbl(txt_falta_aplicar) * 1) - var_importe_aplicar, "###,###,##0.00")
                  frm_importe_aplicar.Visible = False
                  If lv_facturas.ListItems.Count > 0 Then
                     lv_facturas.SetFocus
                  End If
               Else
                  MsgBox "El importe exede al importe del saldo de la factura", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El importe exede al importe de la devolución", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Importe Incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El importe de la devolución es incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      frm_importe_aplicar.Visible = False
   End If
End Sub

Private Sub txt_movimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(txt_movimiento) > 0 Then
         rs.Open "select * from tb_movimientos where vcha_mov_movimiento_id ='" + txt_movimiento + "' and char_mov_afectacion <> 'T'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            cmb_movimientos = rs!vcha_mov_nombre
            rs.Close
            cmb_movimientos.Enabled = False
            txt_numero.Enabled = True
            txt_numero.SetFocus
            txt_movimiento.Enabled = False
         Else
            rs.Close
            MsgBox "Clave de movimiento incorrecta", vbOKOnly, "ATENCION"
            cmb_movimientos.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txt_nombre_clase_Change()
   If KeyCode = 116 Then
      frm_lista.Visible = True
      lst_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_clase_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
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
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Car_clase_id)
            list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_Car_nombre), "", rs!vcha_Car_nombre))
            rs.MoveNext
         Wend
      End If
      rs.Close
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_clase_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_numero_Change()
Dim var_contador_registro As Integer
Dim var_precio_sin_descuentos As Double
      If Trim(txt_numero) <> "" Then
         lv_facturas.ListItems.Clear
         lbl_moneda = ""
         Dim list_item As ListItem
         Dim var_plazo As Integer
         rs.Open "select * from vw_devolucion_nota_credito where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_falta_aplicar = 0
            var_total_neto = 0
            var_importe_total = 0
            var_importe_total_iva = 0
            var_importe_total_descuento_1 = 0
            var_importe_total_descuento_2 = 0
            var_importe_total_descuento_3 = 0
            var_importe_total_subimporte = 0
            While Not rs.EOF
               var_text_descuento = ""
               var_cantidad = 0
               var_precio = 0
               var_precio_sin_descuentos = 0
               var_subimporte = 0
               var_imp_descuento_1 = 0
               var_imp_descuento_2 = 0
               var_imp_descuento_3 = 0
               var_iva = 0
               var_cantidad = Format(IIf(IsNull(rs!Cantidad_leida), 0, rs!Cantidad_leida), "###,###,##0.00")
               If var_unidad_organizacional = "23" Then
                  var_precio = IIf(IsNull(rs!Precio), 0, rs!Precio) / IIf(IsNull(rs!Cantidad_leida), 0, rs!Cantidad_leida)
               Else
                  'cambio realizado por carlos aleman ya que el importe unitario se estaba dividiento entre la cantidad de piezas para ALM GRAL
                  'var_precio = IIf(IsNull(rs!floa_cde_precio), 0, rs!floa_cde_precio) / IIf(IsNull(rs!cantidad_leida), 0, rs!cantidad_leida)
                  var_precio = IIf(IsNull(rs!floa_cde_precio), 0, rs!floa_cde_precio) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
               End If
               'var_precio_sin_descuentos = IIf(IsNull(rs!floa_cde_precio), 0, rs!floa_cde_precio)
               var_precio_sin_descuentos = IIf(IsNull(rs!Precio), 0, rs!Precio)
               var_descuento_1 = IIf(IsNull(rs!floa_cde_descuento_1), 0, rs!floa_cde_descuento_1)
               var_descuento_2 = IIf(IsNull(rs!floa_cde_descuento_2), 0, rs!floa_cde_descuento_2)
               var_descuento_3 = IIf(IsNull(rs!floa_cde_descuento_3), 0, rs!floa_cde_descuento_3)
               var_tipo_Cambio = IIf(IsNull(rs!floa_dev_tipo_cambio), 1, rs!floa_dev_tipo_cambio)
               var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
               var_precio = var_precio * (1 - (var_descuento_1 / 100))
               var_imp_descuento_1 = ((var_precio_sin_descuentos - var_precio) * var_cantidad) / var_tipo_Cambio
               var_precio_sin_descuentos = var_precio
               var_precio = var_precio * (1 - (var_descuento_2 / 100))
               var_imp_descuento_2 = ((var_precio_sin_descuentos - var_precio) * var_cantidad) / var_tipo_Cambio
               var_precio_sin_descuentos = var_precio
               var_precio = var_precio * (1 - (var_descuento_3 / 100))
               var_imp_descuento_3 = ((var_precio_sin_descuentos - var_precio) * var_cantidad) / var_tipo_Cambio
               var_precio = var_precio / var_tipo_Cambio
               If var_empresa = "03" Or var_empresa = "28" Then
                  var_iva = 0
               Else
               ' var_iva = IIf(IsNull(rs!floa_cde_iva), 0, rs!floa_cde_iva)
                 var_iva = 16
               End If
               var_subimporte = var_precio * var_cantidad
               var_importe_total = var_importe_total + var_subimporte
               var_importe_total_descuento_1 = var_importe_total_descuento_1 + var_imp_descuento_1
               var_importe_total_descuento_2 = var_importe_total_descuento_2 + var_imp_descuento_2
               var_importe_total_descuento_3 = var_importe_total_descuento_3 + var_imp_descuento_3
               var_total = var_subimporte
               var_imp_iva = var_total * (var_iva / 100)
               var_importe_total_iva = var_importe_total_iva + var_imp_iva
               var_importe_total_subimporte = var_importe_total_subimporte + var_total
               var_total = var_total + var_imp_iva
               var_total_neto = var_total_neto + var_total
               lbl_moneda = rs!vcha_mon_nombre_plural
               rs.MoveNext
            Wend
            var_total_neto = var_total_neto
            txt_total_neto = Format(Round(var_total_neto, 2), "###,###,##0.00")
            txt_falta_aplicar = Format(Round(var_total_neto, 2), "###,###,##0.00")
            txt_importe = Format(var_total_neto, "###,###,##0.00")
            If rsaux2.State = 1 Then
               rsaux2.Close
            End If
            'rsaux2.Open "select * from vw_suma_importe_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_TIT_TITULAR_ID = '" + txt_titular + "' and IMPORTE_SALDO > 0", cnn, adOpenDynamic, adLockOptimistic
            var_cadena = "SELECT     dbo.TB_SALDOS.VCHA_CLI_CLAVE_ID, dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO, dbo.TB_SALDOS.INTE_CAR_NUMERO AS INTE_ECU_NUMERO_CARGO, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, dbo.TB_SALDOS.FLOA_SAL_IMPORTE AS IMPORTE_SALDO, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO AS IMPORTE_CARGO, 0 AS IMPORTE_ABONOS, dbo.TB_SALDOS.VCHA_SER_SERIE_ID AS VCHA_ECU_sERIE_CARGO FROM         dbo.TB_ENCABEZADO_CARTERA INNER JOIN  dbo.TB_SALDOS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = dbo.TB_SALDOS.VCHA_SER_SERIE_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO AND dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID = dbo.TB_SALDOS.VCHA_CLI_CLAVE_ID AND dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO = dbo.TB_SALDOS.INTE_CAR_NUMERO "
            var_cadena = var_cadena + " WHERE     (dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID = '" + txt_titular + "') AND (dbo.TB_SALDOS.FLOA_SAL_IMPORTE > 0)"
            rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            var_contador_registro = 0
            While Not rsaux2.EOF
               Set list_item = lv_facturas.ListItems.Add(, , rsaux2!vcha_car_documento)
               list_item.SubItems(1) = IIf(IsNull(rsaux2!inte_ecu_numero_cargo), "", rsaux2!inte_ecu_numero_cargo)
               list_item.SubItems(2) = Format(IIf(IsNull(rsaux2!dtim_Car_fecha), "", rsaux2!dtim_Car_fecha), "DD/MM/YY")
               list_item.SubItems(3) = IIf(IsNull(rsaux2!INTE_CAR_PLAZO), 0, rsaux2!INTE_CAR_PLAZO)
               var_plazo = IIf(IsNull(rsaux2!INTE_CAR_PLAZO), 0, rsaux2!INTE_CAR_PLAZO)
               list_item.SubItems(4) = Format(IIf(IsNull(rsaux2!dtim_Car_fecha), "", rsaux2!dtim_Car_fecha + var_plazo), "DD/MM/YY")
               'list_item.SubItems(3) = Format(IIf(IsNull(rsaux2!DTIM_CAR_FECHA_VENCIMIENTO), "", rsaux2!DTIM_CAR_FECHA_VENCIMIENTO), "DD/MM/YY")
               list_item.SubItems(5) = Format(IIf(IsNull(rsaux2!importe_Cargo), "", rsaux2!importe_Cargo), "###,###,##0.00")
               list_item.SubItems(6) = Format(IIf(IsNull(rsaux2!importe_abonos), "", rsaux2!importe_abonos), "###,###,##0.00")
               list_item.SubItems(7) = Format(0, "###,###,##0.00")
               list_item.SubItems(8) = Format(IIf(IsNull(rsaux2!importe_saldo), "", rsaux2!importe_saldo), "###,###,##0.00")
               list_item.SubItems(10) = IIf(IsNull(rsaux2!vcha_ecu_serie_cargo), "", rsaux2!vcha_ecu_serie_cargo)
               list_item.SubItems(11) = IIf(IsNull(rsaux2!vcha_cli_clave_id), "", rsaux2!vcha_cli_clave_id)
               'list_item.SubItems(12) = IIf(IsNull(rsaux2!floa_car_porcentaje_iva), 0, rsaux2!floa_car_porcentaje_iva)
               list_item.SubItems(12) = var_iva
               rsaux2.MoveNext
               var_contador_registro = var_contador_registro + 1
            Wend
            rsaux2.Close
            If var_contador_registro > 13 Then
               lv_facturas.ColumnHeaders(2).Width = 1160
            Else
               lv_facturas.ColumnHeaders(2).Width = 1399.74
            End If
         Else
            MsgBox "El Movimiento no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
End Sub

