VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmfact_merc_vistas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modulo de Facturación de Mercancia a Vistas"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11700
   Icon            =   "frmfact_merc_vistas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11700
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   75
      TabIndex        =   40
      Top             =   390
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   41
         Top             =   480
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3228
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
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   42
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_datos 
      Caption         =   " Cliente "
      Height          =   2295
      Left            =   5160
      TabIndex        =   0
      Top             =   570
      Width           =   6390
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2520
         TabIndex        =   45
         Top             =   1050
         Width           =   3810
      End
      Begin VB.TextBox txt_nombre_establecimiento 
         Height          =   315
         Left            =   2505
         TabIndex        =   44
         Top             =   675
         Width           =   3810
      End
      Begin VB.TextBox txt_nombre_titular 
         Height          =   315
         Left            =   2505
         TabIndex        =   43
         Top             =   315
         Width           =   3810
      End
      Begin VB.TextBox txt_titular 
         Height          =   315
         Left            =   1365
         TabIndex        =   8
         Top             =   315
         Width           =   1125
      End
      Begin VB.TextBox txt_establecimiento 
         Height          =   315
         Left            =   1365
         TabIndex        =   9
         Top             =   675
         Width           =   1125
      End
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   1365
         TabIndex        =   10
         Top             =   1035
         Width           =   1125
      End
      Begin VB.TextBox txt_rfc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1365
         TabIndex        =   11
         Top             =   1410
         Width           =   1770
      End
      Begin VB.TextBox txt_descuentos 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1365
         TabIndex        =   12
         Top             =   1770
         Width           =   2265
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   37
         Top             =   375
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   30
         Top             =   735
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   1095
         Width           =   525
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "RFC:"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   1470
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Descuentos:"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   1830
         Width           =   900
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmfact_merc_vistas.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Movimiento Alt + N"
      Top             =   75
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmfact_merc_vistas.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Buscar Movimiento Alt + B"
      Top             =   75
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   750
      Picture         =   "frmfact_merc_vistas.frx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Movimiento Alt + I"
      Top             =   75
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11265
      Picture         =   "frmfact_merc_vistas.frx":0BD0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   75
      Width           =   330
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   1200
      TabIndex        =   34
      Top             =   525
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         TabIndex        =   35
         Top             =   495
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   30
         TabIndex        =   36
         Top             =   120
         Width           =   3060
      End
   End
   Begin VB.Frame frm_eliminar 
      Height          =   840
      Left            =   4590
      TabIndex        =   31
      Top             =   4620
      Width           =   2910
      Begin VB.TextBox txt_cantidad_eliminar 
         Height          =   330
         Left            =   90
         TabIndex        =   32
         Top             =   390
         Width           =   2745
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad a eliminar"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   33
         Top             =   15
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Movimiento "
      Height          =   1290
      Left            =   90
      TabIndex        =   26
      Top             =   1575
      Width           =   5010
      Begin VB.ComboBox cmb_series 
         Height          =   315
         Left            =   3300
         TabIndex        =   38
         Top             =   360
         Width           =   795
      End
      Begin VB.TextBox txt_cantidad 
         Height          =   315
         Left            =   3315
         TabIndex        =   15
         Top             =   735
         Width           =   1575
      End
      Begin VB.TextBox txt_codigo 
         Height          =   315
         Left            =   780
         TabIndex        =   14
         Top             =   750
         Width           =   1650
      End
      Begin VB.TextBox txt_folio 
         Enabled         =   0   'False
         Height          =   375
         Left            =   780
         TabIndex        =   13
         Top             =   315
         Width           =   1650
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   2820
         TabIndex        =   39
         Top             =   420
         Width           =   405
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   810
         Width           =   600
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   345
         Width           =   600
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   2595
         TabIndex        =   27
         Top             =   795
         Width           =   675
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Artículos "
      Height          =   4425
      Index           =   0
      Left            =   90
      TabIndex        =   23
      Top             =   2865
      Width           =   11490
      Begin MSComctlLib.ListView lv_existencias 
         Height          =   4140
         Left            =   45
         TabIndex        =   24
         Top             =   210
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   7303
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
            Text            =   "Código"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Disponibles       "
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Facturar        "
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Faltan            "
            Object.Width           =   2646
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Salida a Vistas "
      Height          =   990
      Left            =   75
      TabIndex        =   21
      Top             =   570
      Width           =   5040
      Begin VB.TextBox txt_clave_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1035
         TabIndex        =   6
         Top             =   600
         Width           =   1035
      End
      Begin VB.TextBox txt_nombre_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2085
         TabIndex        =   7
         Top             =   600
         Width           =   2880
      End
      Begin VB.TextBox txt_numero_salida 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1035
         TabIndex        =   5
         Top             =   210
         Width           =   1695
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   225
         TabIndex        =   25
         Top             =   660
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   225
         TabIndex        =   22
         Top             =   255
         Width           =   600
      End
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   12450
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2430
      Width           =   2100
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   0
      TabIndex        =   17
      Top             =   330
      Width           =   11655
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   3075
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfact_merc_vistas.frx":120A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfact_merc_vistas.frx":1AE4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfact_merc_vistas.frx":23BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfact_merc_vistas.frx":2C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfact_merc_vistas.frx":3572
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfact_merc_vistas.frx":3B0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfact_merc_vistas.frx":43EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfact_merc_vistas.frx":4CC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfact_merc_vistas.frx":559E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfact_merc_vistas.frx":56B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfact_merc_vistas.frx":57C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfact_merc_vistas.frx":58D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfact_merc_vistas.frx":59E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfact_merc_vistas.frx":5AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfact_merc_vistas.frx":5C7A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3990
      Top             =   -30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfact_merc_vistas.frx":5D8C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmfact_merc_vistas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_descuento_1 As Double
Dim var_descuento_2 As Double
Dim var_descuento_3 As Double
Dim var_subimporte As Double
Dim var_imp_total_desc_1 As Double
Dim var_imp_total_desc_2 As Double
Dim var_imp_total_desc_3 As Double
Dim var_imp_descuento_1 As Double
Dim var_imp_descuento_2 As Double
Dim var_imp_descuento_3 As Double
Dim var_total As Double
Dim var_importe_iva As Double
Dim var_piezas As Double
Dim var_primera_vez As Boolean
Dim var_almacen_origen As String
Dim var_almacen_Destino As String
Dim var_numero_folio As Double
Dim var_cantidad_Salida As Double
Dim var_cantidad_posible As Integer
Dim var_nombre_articulo As String
Dim var_costo As Double
Dim var_precio As Double
Dim var_referencia As String
Dim var_orden_surtido As Integer
Dim var_clave_agente As String
Dim var_clave_establecimiento As String
Dim var_clave_titular As String
Dim var_clave_cliente As String
Dim var_clave_ruta As String
Dim var_plazo As Integer
Dim var_iva As Double
Dim var_agrupador As String
Dim var_folio_enviado As Double
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_estatus_movimiento As String
Dim var_numero_factura As Integer
Dim var_movimiento_dependencia As String
Dim var_clave_moneda As String
Dim var_moneda_local As Integer
Dim var_tipo_Cambio As Double
Dim var_lista_precios As String
Dim var_canal_venta As String
Dim var_serie As String
Dim var_tipo_lista As Integer

Private Sub cmb_clientes_Click()
      txt_cliente = Obtener_llave(cnn, rs, "TB_CLIENTES", "VCHA_CLI_NOMBRE", cmb_clientes, 0, "T")
      var_clave_cliente = txt_cliente
      rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
         var_moneda_local = IIf(IsNull(rs!inte_mon_moneda_local), 1, rs!inte_mon_moneda_local)
         var_clave_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
         txt_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
         var_descuento_1 = IIf(IsNull(rs!floa_gac_Descuento_1), 0, rs!floa_gac_Descuento_1)
         var_descuento_2 = IIf(IsNull(rs!FLOA_GAC_DESCUENTO_2), 0, rs!FLOA_GAC_DESCUENTO_2)
         var_descuento_3 = IIf(IsNull(rs!floa_gac_descuento_3), 0, rs!floa_gac_descuento_3)
         txt_descuentos = Str(var_descuento_1) + "% +" + Str(var_descuento_2) + "% "
         var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
         var_lista_precios = IIf(IsNull(rs!vcha_LIS_LISTA_iD), "", rs!vcha_LIS_LISTA_iD)
         var_canal_venta = IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id)
      End If
      rs.Close
      If Trim(txt_cliente) <> "" Then
         txt_codigo.Enabled = True
      Else
         txt_codigo.Enabled = False
      End If
End Sub

Private Sub cmb_clientes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_codigo.Enabled = True
      txt_codigo.SetFocus
   End If
End Sub

Private Sub cmb_clientes_LostFocus()
   If txt_cliente <> "" Then
      cmb_clientes.Enabled = False
      txt_cliente.Enabled = False
      txt_codigo.Enabled = True
   Else
      cmb_clientes.Enabled = True
      txt_cliente.Enabled = True
      txt_codigo.Enabled = False
   End If
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_titular = lv_lista.selectedItem
            txt_nombre_titular = lv_lista.selectedItem.SubItems(1)
         Else
            txt_titular = ""
            txt_nombre_titular = ""
         End If
         txt_titular.SetFocus
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_establecimiento = lv_lista.selectedItem
            txt_nombre_establecimiento = lv_lista.selectedItem.SubItems(1)
         Else
            txt_establecimiento = ""
            txt_nombre_establecimiento = ""
         End If
         txt_establecimiento.SetFocus
      End If
      If var_tipo_lista = 3 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_cliente = lv_lista.selectedItem
            txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
         Else
            txt_cliente = ""
            txt_nombre_cliente = ""
         End If
         txt_cliente.SetFocus
      End If
   End If
   If keyasii = 27 Then
      If var_tipo_lista = 1 Then
         txt_titular.SetFocus
      End If
      If var_tipo_lista = 2 Then
         txt_establecimiento.SetFocus
      End If
      If var_tipo_lista = 3 Then
         txt_cliente.SetFocus
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct * from vw_establecimientos where vcha_age_agente_id = '" + txt_clave_agente + "' and vcha_tit_titular_id = '" + txt_titular + "' order by vcha_ESB_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      var_tipo_lista = 3
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_establecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct * from vw_establecimientos where vcha_age_agente_id = '" + txt_clave_agente + "' and vcha_tit_titular_id = '" + txt_titular + "' order by vcha_ESB_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ESB_ESTABLECIMIENTO_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ESTABLECIMIENTOS"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_establecimiento_LostFocus()
   If txt_establecimiento <> "" Then
      txt_nombre_establecimiento.Enabled = False
      txt_establecimiento.Enabled = False
      txt_nombre_cliente.Enabled = True
      txt_cliente.Enabled = True
   Else
      txt_nombre_establecimiento.Enabled = True
      txt_establecimiento.Enabled = True
      txt_nombre_cliente.Enabled = False
      txt_cliente.Enabled = False
   End If
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct * from vw_establecimientos where vcha_age_agente_id = '" + txt_clave_agente + "' and vcha_tit_titular_id = '" + txt_titular + "' order by vcha_ESB_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      var_tipo_lista = 3
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_cliente) <> "" Then
         txt_codigo.Enabled = True
         txt_codigo.SetFocus
      End If
   End If
End Sub

Private Sub txt_nombre_establecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct * from vw_establecimientos where vcha_age_agente_id = '" + txt_clave_agente + "' and vcha_tit_titular_id = '" + txt_titular + "' order by vcha_ESB_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ESB_ESTABLECIMIENTO_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ESTABLECIMIENTOS"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_establecimiento) <> "" Then
         txt_cliente.Enabled = True
         txt_nombre_cliente.Enabled = True
         txt_cliente.SetFocus
      End If
   End If
End Sub

Private Sub cmb_series_Click()
   var_serie = cmb_series
End Sub


Private Sub txt_nombre_titular_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct * from vw_titulares_1 where vcha_age_agente_id = '" + txt_clave_agente + "' order by vcha_tit_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tit_titular_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TITULARES"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_titular_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_establecimiento.Enabled = True
      txt_nombre_establecimiento.Enabled = True
      txt_establecimiento.SetFocus
   End If
End Sub

Private Sub txt_nombre_titular_LostFocus()
   If txt_titular <> "" Then
      txt_nombre_establecimiento.Enabled = True
      txt_establecimiento.Enabled = True
      txt_nombre_titular.Enabled = False
      txt_titular.Enabled = False
      txt_nombre_cliente.Enabled = False
      txt_cliente.Enabled = False
   Else
      txt_nombre_titular.Enabled = True
      txt_titular.Enabled = True
      txt_nombre_establecimiento.Enabled = False
      txt_establecimiento.Enabled = False
      txt_nombre_cliente.Enabled = False
      txt_cliente.Enabled = False
   End If
End Sub

Private Sub cmd_buscar_Click()
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
   Set TB_ENTRADAS_INSERTA = New TB_ENTRADAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_LIBERA_APARTADOS = New TB_LIBERA_APARTADOS
   Set TB_SALIDA_VISTAS_I = New TB_SALIDA_VISTAS_I
   Set TB_ENTRADAS_VISTAS_I = New TB_ENTRADAS_VISTAS_I
   Set TB_ENC_FACTURAS_I = New TB_ENC_FACTURAS_I
   Set TB_DET_FACTURAS_I = New TB_DET_FACTURAS_I
   Set TB_INCREMENTA_FACTURA = New TB_INCREMENTA_FACTURA
   Dim var_numero_renglones As Integer
   Dim var_conta As Integer
         frm_busqueda.Visible = True
         txt_busqueda_folio.SetFocus
End Sub

Private Sub cmd_imprimir_Click()
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
   Set TB_ENTRADAS_INSERTA = New TB_ENTRADAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_LIBERA_APARTADOS = New TB_LIBERA_APARTADOS
   Set TB_SALIDA_VISTAS_I = New TB_SALIDA_VISTAS_I
   Set TB_ENTRADAS_VISTAS_I = New TB_ENTRADAS_VISTAS_I
   Dim var_numero_renglones As Integer
   Dim var_moneda_local As Integer
   Dim var_conta As Integer
   Dim var_posible_tipo_cambio As Boolean



   Dim var_i As Integer
   Dim var_j As Integer
   Dim var_k As Integer
   Dim var_cliente_str As String
   Dim var_expedicion_str As String
   Dim var_domicilio_str As String
   Dim var_ciudad_str As String
   Dim var_agente_str As String
   Dim var_linea As String
   
   Dim var_cantidad_str As String
   Dim var_descuento_1_str As Double
   Dim var_descuento_2_str As Double
   Dim var_descuento_3_str As Double
   Dim var_precio As Double
   Dim var_precio_str As String
   Dim var_importe_str As String
   Dim var_subimporte_str As String
   Dim var_cantidad_letra_str As String
   Dim var_iva_str As String
   Dim var_rfc_str As String
   Dim var_dia_str As String
   Dim var_mes_str As String
   Dim var_año_str As String
   Dim var_porcentaje As Double
   Dim var_Archivo_str As String
   Dim var_relacion As String
   Dim var_marca_promocion As String
   Dim var_contador_promociones As Integer
   Dim var_cantidad_total As Double
   Dim var_cantidad_total_str As String
   Dim var_si As Integer
   Dim var_factura_inicio As Double
   If var_numero_folio > 0 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         MsgBox "El movimiento ya fue impreso", vbOKOnly, "ATENCION"
      Else
         rs.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
         var_factura_inicio = rs!inte_ser_factura
         rs.Close
         var_si = MsgBox("Se va a imprimir el movimiento en la factura " + Str(var_factura_inicio) + " prepare la impresora", vbOKCancel, "ATENCION")
         If var_si = 1 Then
            var_si = MsgBox("Confirmar la impresión de la factura " + CStr(var_factura_inicio), vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_moneda_local = 0
               rs.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
               If var_factura_inicio <> rs!inte_ser_factura Then
                  MsgBox "El número de factura a cambiado y el proceso de facturación se cancela", vbOKOnly, "ATENCION"
                  rs.Close
               Else
                  rs.Close
                  var_posible_tipo_cambio = False
                  rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_moneda_local = IIf(IsNull(rs!inte_mon_moneda_local), 0, rs!inte_mon_moneda_local)
                  End If
                  If var_moneda_local = 0 Then
                     rsaux2.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        var_posible_tipo_cambio = True
                        var_tipo_Cambio = rsaux2!floa_tca_importe
                     End If
                     rsaux2.Close
                  Else
                     var_tipo_Cambio = 1
                     var_posible_tipo_cambio = True
                  End If
                  rs.Close
                  If var_posible_tipo_cambio = True Then
                     cnn.BeginTrans
                     Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio)
                     rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     var_subimporte = 0
                     var_imp_total_desc_1 = 0
                     var_imp_total_desc_2 = 0
                     var_imp_total_desc_3 = 0
                     var_imp_descuento_1 = 0
                     var_imp_descuento_2 = 0
                     var_imp_descuento_3 = 0
                     var_descuento_1 = 0
                     var_descuento_2 = 0
                     var_descuento_3 = 0
                     var_total = 0
                     var_importe_iva = 0
                     var_piezas = 0
                     var_conta = 1
                     var_conta = 0
                     While Not rs.EOF
                        var_conta = var_conta + 1
                        var_inserta = False
                        var_inserta = TB_LIBERA_APARTADOS.Anadir(var_almacen_Destino, rs!vcha_Art_articulo_id, 0 - rs!FLOA_sAL_cANTIDAD)
                        var_inserta = False
                        Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2]) values ('" + rs(0).Value + "', '" + rs(1).Value + "', '" + rs(2).Value + "', '" + rs(3).Value + "', " + CStr(rs(4).Value) + ", '" + rs(5).Value + "' , " + Str(rs(6).Value) + ", " + CStr(rs(7).Value) + ", " + CStr(rs(8).Value * var_tipo_Cambio) + ", " + CStr(rs(9).Value) + ", " + CStr(rs(10).Value) + ",  " + CStr(rs(11).Value) + ")"
                        rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                        'var_inserta = TB_SALIDAS_INSERTA.Anadir(rs(0).Value, rs(1).Value, rs(2).Value, rs(3).Value, rs(4).Value, rs(5).Value, rs(6).Value, rs(7).Value, rs(8).Value * var_tipo_cambio, rs(9).Value)
                        var_inserta = False
                        var_inserta = TB_ENTRADAS_VISTAS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_folio_enviado, var_clave_movimiento, var_numero_folio, rs!vcha_Art_articulo_id, rs!FLOA_sAL_cANTIDAD, rs!floa_sal_costo, rs!floa_Sal_precio)
                        rs.MoveNext
                     Wend
                     rs.Close
                     var_estatus_movimiento = "I"
                     var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, "I", Now, var_tipo_Cambio)
                     rs.Open "execute FACTURA_MERCANCIA_VISTAS '" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + Str(var_numero_folio) + ",'' ,'','" + var_serie + "','FA'", cnn, adOpenDynamic, adLockOptimistic
                     txt_codigo.Enabled = False
                     txt_foco.Enabled = False
                     txt_cantidad.Enabled = False
                     
                     cnn.CommitTrans
                     'SE IMPRIME LA FACTURA
                     rsaux3.Open "select distinct vcha_car_documento, vcha_ser_Serie_id, inte_Car_numero from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + txt_folio, cnn, adOpenDynamic, adLockOptimistic
                     While Not rsaux3.EOF
                        rs.Open "select * from vw_facturas_vistas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + rsaux3!vcha_Car_documento + "' and vcha_ser_serie_id = '" + rsaux3!VCHA_SER_SERIE_ID + "' and inte_Car_numero = " + Str(rsaux3!inte_Car_numero), cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt") For Output As #1
                           Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                           'Print #1, Chr(27) + Chr(67) + Chr(44) + Chr(13)
                           'Print #1, ""
                           Print #1, Spc(92); Str(rs!inte_Car_numero)
                           Print #1, ""
                           Print #1, ""
                           Print #1, Spc(93); "FECHA: "; Format(rs!dtim_Car_fecha, "Short Date")
                           Print #1, ""
                           Print #1, Spc(92); Str(rs!INTE_CAR_PLAZO) + " DIAS DE VENCIMIENTO"
                           var_cliente_str = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                           For var_j = 1 + Len(Trim(var_cliente_str)) To 83
                               var_cliente_str = var_cliente_str + " "
                           Next var_j
                           var_cliente_str = var_cliente_str + "AGUASCALIENTES, AGS."
                           Print #1, ""
                           Print #1, Spc(10); var_cliente_str
                           var_domicilio_str = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                           For var_j = 1 + Len(Trim(var_domicilio_str)) To 83
                               var_domicilio_str = var_domicilio_str + " "
                           Next var_j
                           var_agente_str = ""
                           var_agente_str = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                           For var_j = 1 + Len(Trim(var_agente_str)) To 8
                               var_agente_str = var_agente_str + " "
                           Next var_j
                           var_agente_str = var_agente_str + IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
                           var_domicilio_str = var_domicilio_str
                           'Print #1, Spc(111); var_agente
                           Print #1, Spc(10); var_domicilio_str
                           var_ciudad_str = ""
                           var_ciudad_str = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                           For var_j = 1 + Len(Trim(var_ciudad_str)) To 37
                               var_ciudad_str = var_ciudad_str + " "
                           Next var_j
                           var_estado_str = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                           For var_j = 1 + Len(Trim(var_estado_str)) To 46
                               var_estado_str = var_estado_str + " "
                           Next var_j
                           var_ciudad_str = var_ciudad_str + var_estado_str
                           
                           For var_j = 1 + Len(Trim(var_ciudad_str)) To 14
                               var_ciudad_str = var_ciudad_str + " "
                           Next var_j
                        
                           var_ciudad_str = var_ciudad_str + var_agente_str
                           var_relacion = "RMV: " + CStr(IIf(IsNull(rs!INTE_EMO_NUMERO_ORIGEN), "", rs!INTE_EMO_NUMERO_ORIGEN))
                           Print #1, Spc(10); var_ciudad_str
                           var_rfc_str = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                           var_rfc_str = "RFC:  " + var_rfc_str
                           For var_j = 1 + Len(Trim(var_rfc_str)) To 89
                               var_rfc_str = var_rfc_str + " "
                           Next var_j
                           var_rfc_str = var_rfc_str + var_relacion
                           Print #1, Spc(4); var_rfc_str
                           Print #1, Spc(10); IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
                           Print #1, ""
                           Print #1, ""
                           var_importe_descuento_1 = 0
                           var_importe_descuento_2 = 0
                           var_importe_descuento_3 = 0
                           var_contador_promociones = 0
                           var_cantidad_total = 0
                           For var_k = 1 To var_renglones_factura
                               If Not rs.EOF Then
                                  var_linea = ""
                                  var_marca_promocion = " "
                                  var_promocion_1 = IIf(IsNull(rs!floa_sal_promocion_1), 0, rs!floa_sal_promocion_1)
                                  var_promocion_2 = IIf(IsNull(rs!FLOA_SAL_PROMOCION_2), 0, rs!FLOA_SAL_PROMOCION_2)
                                  If var_promocion_1 > 0 Then
                                     var_marca_promocion = "*"
                                     var_contador_promociones = var_contador_promociones + 1
                                  End If
                                  If var_promocion_2 > 0 Then
                                     var_marca_promocion = "*"
                                     var_contador_promociones = var_contador_promociones + 1
                                  End If
                                  var_linea = IIf(IsNull(rs!vcha_agr_agrupador_id), "", rs!vcha_agr_agrupador_id)
                                  For var_j = 1 + Len(Trim(var_linea)) To 15
                                      var_linea = var_linea + " "
                                  Next var_j
                                  var_linea = var_linea + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                  For var_j = 1 + Len(Trim(var_linea)) To 91
                                      var_linea = var_linea + " "
                                  Next var_j
                                  var_linea = var_linea + var_marca_promocion
                                  var_cantidad_str = Format(IIf(IsNull(rs!FLOA_sAL_cANTIDAD), 0, rs!FLOA_sAL_cANTIDAD), "###,###,##0.00")
                                  var_cantidad_total = var_cantidad_total + IIf(IsNull(rs!FLOA_sAL_cANTIDAD), 0, rs!FLOA_sAL_cANTIDAD)
                                  If Len(Trim(var_cantidad_str)) < 14 Then
                                     For var_j = 1 + Len(Trim(var_cantidad_str)) To 14
                                        var_cantidad_str = " " + var_cantidad_str
                                     Next var_j
                                  End If
                                  var_precio = IIf(IsNull(rs!floa_Sal_precio), 0, rs!floa_Sal_precio)
                                  var_descuento_1 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
                                  var_descuento_2 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2)
                                  var_descuento_3 = IIf(IsNull(rs!floa_car_porcentaje_descuento_3), 0, rs!floa_car_porcentaje_descuento_3)
                                  var_porcentaje = (100 - var_descuento_1) / 100
                                  var_precio = var_precio * var_porcentaje
                                  var_importe_descuento_1_2 = (IIf(IsNull(rs!floa_Sal_precio), 0, rs!floa_Sal_precio) - var_precio)
                                  var_importe_descuento_1 = var_importe_descuento_1 + (IIf(IsNull(rs!floa_Sal_precio), 0, rs!floa_Sal_precio) - var_precio)
                                  var_precio = var_precio * ((100 - var_descuento_2) / 100)
                                  var_importe_descuento_2 = var_importe_descuento_2 + (IIf(IsNull(rs!floa_Sal_precio), 0, rs!floa_Sal_precio) - (var_importe_descuento_1_2 + var_precio))
                                  var_precio = var_precio * ((100 - var_descuento_3) / 100)
                                  var_precio = var_precio / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                                  'var_precio_str = Format(var_precio / IIf(IsNull(rs!cantidad), 0, rs!cantidad), "###,###,##0.00")
                                  var_precio_str = Format(IIf(IsNull(rs!floa_Sal_precio), 0, rs!floa_Sal_precio) / IIf(IsNull(rs!FLOA_sAL_cANTIDAD), 0, rs!FLOA_sAL_cANTIDAD), "###,###,##0.00")
                                  If Len(Trim(var_precio_str)) < 14 Then
                                     For var_j = 1 + Len(Trim(var_precio_str)) To 14
                                         var_precio_str = " " + var_precio_str
                                     Next var_j
                                  End If
                                  var_linea = var_linea + var_cantidad_str + var_precio_str
                                  var_importe_str = Format((IIf(IsNull(rs!floa_Sal_precio), 0, rs!floa_Sal_precio)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                  If Len(Trim(var_importe_str)) < 14 Then
                                     For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                         var_importe_str = " " + var_importe_str
                                     Next var_j
                                  End If
                                  var_linea = var_linea + var_importe_str
                                   
                                  Print #1, var_linea
                                  rs.MoveNext
                               Else
                                  Print #1, ""
                               End If
                           Next var_k
                           Print #1, ""
                           Print #1, ""
                           rs.MoveFirst
                           var_cantidad_total_str = Format(var_cantidad_total, "###,###,##0.00")
                           var_cantidad_letra_str = rs!vcha_car_importe_letra
                           var_importe_descuento_1_str = Format(var_importe_descuento_1, "###,###,##0.00")
                             
                           If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                              For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                  var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                              Next var_j
                           End If
                           var_importe_descuento_2_str = Format(var_importe_descuento_2, "###,###,##0.00")
                           If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                              For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                  var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                              Next var_j
                           End If
                           var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                           If Len(Trim(var_linea)) < 120 Then
                              For var_j = 1 + Len(Trim(var_linea)) To 120
                                  var_linea = var_linea + " "
                              Next var_j
                           End If
                           Print #1, var_linea + var_importe_descuento_1_str
                           var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%"
                           If Len(Trim(var_linea)) < 120 Then
                              For var_j = 1 + Len(Trim(var_linea)) To 120
                                  var_linea = var_linea + " "
                              Next var_j
                           End If
                           var_linea = var_linea + var_importe_descuento_2_str
                           Print #1, var_linea
                           If var_contador_promociones > 0 Then
                              Print #1, "PROMOCION EN ARTICULOS MARCADOS CON *"
                           Else
                              Print #1, ""
                           End If
                           var_rfc_str = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                           var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                           If Len(Trim(var_linea)) < 120 Then
                              For var_j = 1 + Len(Trim(var_linea)) To 120
                                  var_x = var_j Mod 2
                                  If var_x >= 1 Then
                                     var_linea = " " + var_linea
                                  Else
                                     var_linea = var_linea + " "
                                  End If
                              Next var_j
                           End If
                           
                           If Len(Trim(var_rfc_str)) = 0 Then
                              var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                              If Len(Trim(var_subimporte_str)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                     var_subimporte_str = " " + var_subimporte_str
                                 Next var_j
                              End If
                              var_iva_str = "      -        "
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
                           var_linea = var_linea + var_iva_str
                        
                           If Len(Trim(var_subimporte_str)) < 14 Then
                              For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                  var_subimporte_str = " " + var_subimporte_str
                              Next var_j
                           End If
                           Print #1, Spc(101); var_cantidad_total_str; Spc(6); var_subimporte_str
                        
                           Print #1, var_linea
                           var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                           If Len(Trim(var_importe_str)) < 120 Then
                              For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                  var_importe_str = " " + var_importe_str
                              Next var_j
                           End If
                           var_linea = "                                             ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                            "
                           var_linea = var_linea + var_importe_str
                           Print #1, var_linea
                           Print #1, Spc(4); "AGUASCALIENTES, AGS"; Spc(3); Format(rs!dtim_Car_fecha, "Short Date")
                           var_linea = ""
                           Print #1, Spc(45); var_linea
                           var_dia_str = Day(rs!dtim_Car_fecha)
                           var_mes_str = Month(rs!dtim_Car_fecha)
                           var_año_str = Year(rs!dtim_Car_fecha)
                           var_linea = var_dia
                           If Len(Trim(var_linea)) < 14 Then
                              For var_j = 1 + Len(Trim(var_linea)) To 14
                                  var_linea = var_linea + " "
                              Next var_j
                           End If
                           var_linea = var_linea + var_mes_str
                           If Len(Trim(var_linea)) < 50 Then
                              For var_j = 1 + Len(Trim(var_linea)) To 50
                                  var_linea = var_linea + " "
                              Next var_j
                           End If
                           Print #1, Spc(70); var_linea
                           var_linea = ""
                           var_linea = var_año_str
                           If Len(Trim(var_linea)) < 15 Then
                              For var_j = 1 + Len(Trim(var_linea)) To 15
                                  var_linea = var_linea + " "
                              Next var_j
                           End If
                           var_linea = var_linea + var_importe_str
                           If Len(Trim(var_linea)) < 24 Then
                              For var_j = 1 + Len(Trim(var_linea)) To 24
                                  var_linea = " " + var_linea
                              Next var_j
                           End If
                           var_linea = var_linea + " " + var_cantidad_letra_str
                           Print #1, Spc(2); var_linea
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                           Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!VCHA_CLI_COLONIA), "", rs!VCHA_CLI_DIRECCION))
                           Print #1, Spc(5); Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre))
                           Close #1
                           Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat") For Output As #2
                           var_Archivo = App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat"
                           Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt lpt1"
                           Close #2
                           x = Shell(var_Archivo, vbHide)
                        '''factura vieja
                        End If
                        rs.Close
                        rsaux3.MoveNext
                     Wend
                     rsaux3.Close
                     'SE TERMINA DE IMPRIMIR LA FACTURA
                     MsgBox "Se a terminado de imprimir la factura", vbOKOnly, "ATENCION"
                  Else
                     MsgBox "No es posible facturar ya que no se a indicado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
                  End If
               End If
            Else
               MsgBox "Se a cancelado la impresión de la factura", vbOKOnly, "ATENCION"
            End If
         End If
      End If
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
   Set TB_ENTRADAS_INSERTA = New TB_ENTRADAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_LIBERA_APARTADOS = New TB_LIBERA_APARTADOS
   Set TB_SALIDA_VISTAS_I = New TB_SALIDA_VISTAS_I
   Set TB_ENTRADAS_VISTAS_I = New TB_ENTRADAS_VISTAS_I
   Set TB_ENC_FACTURAS_I = New TB_ENC_FACTURAS_I
   Set TB_DET_FACTURAS_I = New TB_DET_FACTURAS_I
   Set TB_INCREMENTA_FACTURA = New TB_INCREMENTA_FACTURA
   Dim var_numero_renglones As Integer
   Dim var_conta As Integer
   var_estatus_movimiento = ""
   txt_numero_salida.Enabled = True
   txt_numero_salida = ""
   txt_clave_agente = ""
   txt_nombre_agente = ""
   txt_cliente = ""
   cmb_clientes.Clear
   txt_rfc = ""
   txt_descuentos = ""
   txt_establecimiento = ""
   txt_titular = ""
   txt_nombre_titular.Clear
   txt_nombre_establecimiento.Clear
   lv_existencias.ListItems.Clear
   var_primera_vez = True
   var_almacen_Destino = ""
   var_almacen_origen = ""
   txt_foco.Enabled = False
   txt_cantidad.Enabled = False
   txt_codigo.Enabled = False
   txt_folio = ""
   txt_numero_salida.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = 116 Then
      frmexisten_rapidas.Show 1
   End If
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 66 Then
      cmd_buscar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   frm_lista.Visible = False
   var_cadena_seguridad = ""
   Top = 0
   Left = 0
   frm_eliminar.Visible = False
   frm_busqueda.Visible = False
   txt_foco.Enabled = False
   var_cantidad_Salida = 0
   var_primera_vez = True
   var_iva = 15
   var_clave_movimiento = "FV"
   txt_establecimiento = ""
   txt_colonia = ""
   txt_titular = ""
   txt_nombre_titular = ""
   txt_nombre_establecimiento = ""
   txt_titular.Enabled = False
   txt_nombre_titular.Enabled = False
   txt_nombre_establecimiento.Enabled = False
   'cmb_clientes.Enabled = False
   txt_establecimiento.Enabled = False
   txt_cliente.Enabled = False
   rs.Open "select vcha_ser_serie_id from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_contador_serie = 0
      While Not rs.EOF
         var_contador_serie = var_contador_serie + 1
         rs.MoveNext
      Wend
      rs.MoveFirst
      txt_codigo.Enabled = True
      txt_cantidad.Enabled = True
      Call RecsetToCombo(cmb_series.hwnd, rs, 0)
      If var_contador_serie > 1 Then
         cmb_series.Enabled = True
      Else
         cmb_series.Enabled = False
      End If
      rs.MoveFirst
      cmb_series = rs!VCHA_SER_SERIE_ID
      var_serie = rs!VCHA_SER_SERIE_ID
   Else
      MsgBox "No se a indicado una serie para esta Unidad organizacional", vbOKOnly, "ATENCION"
      txt_codigo.Enabled = False
      txt_cantidad.Enabled = False
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_fact_merc_vistas)
End Sub

Private Sub lv_existencias_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If Trim(var_estatus_movimiento) = "" Then
         frm_eliminar.Visible = True
         txt_cantidad_eliminar.SetFocus
      Else
         MsgBox "Ya no es posible modificar el movimiento", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub Text1_Change()

End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
   frm_datos.Visible = True
   txt_direccion.SetFocus
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   Dim var_numero_busqueda As Integer
   If KeyAscii = 13 Then
      If txt_busqueda_folio <> "" Then
         rs.Open "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id ='" + var_unidad_organizacional + "' and inte_emo_numero = " + txt_busqueda_folio + "and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_primera_vez = False
            var_estatus_movimiento = rs!char_Emo_estatus
            var_numero_busqueda = rs!INTE_EMO_NUMERO_ORIGEN
            var_numero_folio = txt_busqueda_folio
            var_clave_moneda = rs!vcha_mon_moneda_id
            txt_folio = var_numero_folio
            txt_numero_salida = var_numero_busqueda
            var_clave_establecimiento = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
            var_clave_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
            txt_establecimiento = var_clave_establecimiento
            txt_cliente = var_clave_cliente
            rs.Close
            rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               txt_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
               var_lista_precios = IIf(IsNull(rs!vcha_LIS_LISTA_iD), "", rs!vcha_LIS_LISTA_iD)
               If txt_titular <> "" Then
                  rsaux.Open "select * from tb_titulares where vcha_tit_titular_id = '" + txt_titular + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     txt_nombre_titular.Text = rsaux!VCHA_TIT_NOMBRE
                  Else
                     txt_nombre_titular.Text = ""
                  End If
                  rsaux.Close
               End If
            End If
            rs.Close
            If var_clave_establecimiento <> "" Then
               rs.Open "select * from tb_establecimientos where vcha_esb_establecimiento_id = '" + var_clave_establecimiento + "'", cnn, adOpenDynamic, adLockOptimistic
               txt_nombre_establecimiento = rs!VCHA_ESB_NOMBRE
               rs.Close
            End If
            rs.Open "select * from vw_facturacion_vistas where inte_com_numero = " + Str(var_numero_busqueda), cnn, adOpenDynamic, adLockOptimistic
            lv_existencias.ListItems.Clear
            If Not rs.EOF Then
               var_folio_enviado = txt_numero_salida
               txt_clave_agente = rs!VCHA_AGE_AGENTE_ID
               txt_nombre_agente = rs!VCHA_AGE_NOMBRE
               var_almacen_origen = rs!VCHA_EMO_ALMACEN_DESTINO
               var_almacen_Destino = rs!VCHA_ALM_ALMACEN_ID
               var_referencia = rs!vcha_com_Referencia
               var_movimiento_dependencia = IIf(IsNull(rs!vcha_mov_movimiento_dependencia), "", rs!vcha_mov_movimiento_dependencia)
               While Not rs.EOF
                  Set list_item = lv_existencias.ListItems.Add(, , rs!vcha_Art_articulo_id)
                  list_item.SubItems(1) = Trim(rs!vcha_art_nombre_español)
                  list_item.SubItems(2) = Format(rs!FLOA_COM_CANTIDAD_ENVIADA - rs!floa_com_cantidad_recibida, "###,###,##0.00")
                  rsaux.Open "select floa_sal_cantidad from tb_temporal_Salidas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id ='" + var_unidad_organizacional + "' and inte_sal_numero = " + txt_busqueda_folio + " and vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "' and vcha_mov_movimiento_id ='" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     list_item.SubItems(3) = Format(rsaux!FLOA_sAL_cANTIDAD, "###,###,##0.00")
                     list_item.SubItems(4) = Format(rs!FLOA_COM_CANTIDAD_ENVIADA - rs!floa_com_cantidad_recibida, "###,###,##0.00")
                  Else
                     list_item.SubItems(3) = Format(0, "###,###,##0.00")
                     list_item.SubItems(4) = Format(rs!FLOA_COM_CANTIDAD_ENVIADA - rs!floa_com_cantidad_recibida, "###,###,##0.00")
                  End If
                  rsaux.Close
                  var_contador = var_contador + 1
                  rs.MoveNext:
               Wend
               rs.MoveFirst
               txt_numero_salida.Enabled = False
               rs.Close
               rsaux.Open "select * from vw_clientes where vcha_cli_clave_id = '" + txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_clave_titular = IIf(IsNull(rsaux!vcha_tit_titular_id), "", rsaux!vcha_tit_titular_id)
                  txt_rfc = IIf(IsNull(rsaux!VCHA_CLI_RFC), "", rsaux!VCHA_CLI_RFC)
                  cmb_clientes = IIf(IsNull(rsaux!VCHA_CLI_NOMBRE), "", rsaux!VCHA_CLI_NOMBRE)
                  var_plazo = IIf(IsNull(rsaux!inte_pla_dias), 0, rsaux!inte_pla_dias)
                  var_descuento_1 = IIf(IsNull(rsaux!floa_gac_Descuento_1), 0, rsaux!floa_gac_Descuento_1)
                  var_descuento_2 = IIf(IsNull(rsaux!FLOA_GAC_DESCUENTO_2), 0, rsaux!FLOA_GAC_DESCUENTO_2)
                  var_descuento_3 = IIf(IsNull(rsaux!floa_gac_descuento_3), 0, rsaux!floa_gac_descuento_3)
                  txt_descuentos = Str(var_descuento_1) + "% +" + Str(var_descuento_2) + "% "
               End If
               rsaux.Close
               If var_estatus_movimiento = "I" Then
                  rsaux2.Open "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + txt_busqueda_folio, cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     var_descuento_1 = IIf(IsNull(rsaux2!floa_emo_descuento_1), 0, rsaux2!floa_emo_descuento_1)
                     var_descuento_2 = IIf(IsNull(rsaux2!floa_emo_descuento_2), 0, rsaux2!floa_emo_descuento_2)
                     var_descuento_3 = IIf(IsNull(rsaux2!FLOA_EMO_DESCUENTO_3), 0, rsaux2!FLOA_EMO_DESCUENTO_3)
                     txt_descuentos = Str(var_descuento_1) + "% +" + Str(var_descuento_2) + "% "
                  Else
                  End If
                  txt_codigo = ""
                  txt_codigo.Enabled = False
                  txt_cantidad.Enabled = False
                  txt_foco.Enabled = False
                  rsaux2.Close
               Else
                  txt_codigo = ""
                  txt_codigo.Enabled = True
                  txt_cantidad.Enabled = False
                  txt_foco.Enabled = False
               End If
               txt_nombre_establecimiento.Enabled = False
               txt_establecimiento.Enabled = False
               cmb_clientes.Enabled = False
               txt_cliente.Enabled = False
               txt_titular.Enabled = False
               txt_nombre_titular.Enabled = False
            End If
         Else
            rs.Close
            MsgBox "El número de movimiento no existe", vbOKOnly, "ATENCION"
         End If
         frm_busqueda.Visible = False
      End If
   End If
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   Set TB_ARCH_COMPARACION_M = New TB_ARCH_COMPARACION_M
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Dim var_cantidad_eliminar As Double
   Dim var_codigo As String
   Dim var_numero_salida As Double
   var_numero_salida = txt_numero_salida
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If txt_cantidad_eliminar <> "" Then
         var_cantidad_posible = lv_existencias.selectedItem.SubItems(3)
         var_cantidad_eliminar = txt_cantidad_eliminar
         If var_cantidad_eliminar <= var_cantidad_posible Then
            var_codigo = lv_existencias.selectedItem
            var_actualiza = False
            var_actualiza = TB_ARCH_COMPARACION_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_movimiento_dependencia, var_numero_salida, "G", txt_clave_agente, var_codigo, 0 - var_cantidad_eliminar, var_referencia)
            var_actualiza = False
            var_actualiza = TB_TEMPORAL_SALIDAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, var_codigo, 0 - var_cantidad_eliminar)
            lv_existencias.selectedItem.SubItems(3) = Format(lv_existencias.selectedItem.SubItems(3) - var_cantidad_eliminar, "###,###,##0.00")
            lv_existencias.selectedItem.SubItems(4) = Format(lv_existencias.selectedItem.SubItems(4) + var_cantidad_eliminar, "###,###,##0.00")
            frm_eliminar.Visible = False
         Else
            MsgBox "No es posible eliminar esta cantidad", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   If KeyAscii = 27 Then
      frm_eliminar.Visible = False
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   txt_cantidad_eliminar = ""
   frm_eliminar.Visible = False
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_cantidad) = "" Then
         txt_cantidad = 0
      End If
      var_cantidad_Salida = txt_cantidad
      valor = txt_codigo
      Set itmfound = lv_existencias.findItem(valor, lvwText, , lvwPartial)
      itmfound.EnsureVisible
      itmfound.Selected = True
      var_cantidad_posible = lv_existencias.selectedItem.SubItems(2)
      If var_cantidad_Salida > var_cantidad_posible Then
         MsgBox "La cantidad exede a la cantidad posible as facturar", vbOKOnly, "ATENCION"
      Else
         var_tipo_Cambio = 1
         If var_moneda_local = 0 Then
            rs.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_tipo_Cambio = IIf(IsNull(rs!mone_tca_importe), 1, rs!mone_tca_importe)
               var_posible_tipo_cambio = True
            Else
               var_posible_tipo_cambio = False
            End If
            rs.Close
         Else
            var_posible_tipo_cambio = True
         End If
         If var_posible_tipo_cambio = True Then
            txt_foco.Enabled = True
            txt_foco.SetFocus
            txt_cantidad.Enabled = False
         Else
            MsgBox "No se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub


Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      rsaux.Open "select distinct * from vw_establecimientos where vcha_age_agente_id = '" + txt_clave_agente + "' and vcha_tit_titular_id = '" + txt_titular + "' AND VCHA_ESB_ESTABLECIMIENTO_ID = '" + txt_establecimiento + "' AND VCHA_CLI_CLAVE_ID = '" + txt_cliente + "'", cnn, adOpenDynamic, adLockBatchOptimistic
      If Not rsaux.EOF Then
         rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
            var_moneda_local = IIf(IsNull(rs!inte_mon_moneda_local), 1, rs!inte_mon_moneda_local)
            var_clave_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
            txt_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
            var_descuento_1 = IIf(IsNull(rs!floa_gac_Descuento_1), 0, rs!floa_gac_Descuento_1)
            var_descuento_2 = IIf(IsNull(rs!FLOA_GAC_DESCUENTO_2), 0, rs!FLOA_GAC_DESCUENTO_2)
            var_descuento_3 = IIf(IsNull(rs!floa_gac_descuento_3), 0, rs!floa_gac_descuento_3)
            txt_descuentos = Str(var_descuento_1) + "% +" + Str(var_descuento_2) + "% "
            var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
            var_lista_precios = IIf(IsNull(rs!vcha_LIS_LISTA_iD), "", rs!vcha_LIS_LISTA_iD)
            var_canal_venta = IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id)
         End If
      Else
         var_clave_moneda = ""
         var_moneda_local = 1
         var_clave_titular = ""
         txt_rfc = ""
         var_descuento_1 = 0
         var_descuento_2 = 0
         var_descuento_3 = 0
         txt_descuentos = ""
         var_clave_moneda = ""
         var_lista_precios = ""
         var_canal_venta = ""
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
      If Trim(txt_cliente) <> "" Then
         txt_codigo.Enabled = True
      Else
         txt_codigo.Enabled = False
      End If
   End If
End Sub

Private Sub txt_codigo_GotFocus()
   txt_codigo = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Dim var_posible As Boolean
   If KeyAscii = 13 Then
      var_verificador = True
      If Len(Trim(txt_codigo)) = 12 Then
         Call calcula_verificador(Trim(txt_codigo))
      End If
      If var_verificador = True Then
         If txt_codigo <> "" Then
            rs.Open "select * from tb_Articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rs.Close
            Else
               rs.Close
               rs.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_posible = True
                     txt_codigo = rs!vcha_Art_articulo_id
                     rsaux.Close
                     rs.Close
                  Else
                     var_posible = False
                     rsaux.Close
                     rs.Close
                  End If
               Else
                  rs.Close
               End If
            End If
            Cadena = "select * from vw_facturacion_vistas where inte_com_numero = " + txt_numero_salida + " and vcha_art_articulo_id = '" + txt_codigo + "'"
            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_precio = rs!floa_Sal_precio
               var_costo = rs!floa_sal_costo
               var_nombre_articulo = Trim(rs!vcha_art_nombre_español)
               rs.Close
               var_cantidad_Salida = 1
               valor = txt_codigo
               Set itmfound = lv_existencias.findItem(valor, lvwText, , lvwPartial)
               itmfound.EnsureVisible
               itmfound.Selected = True
               var_cantidad_posible = lv_existencias.selectedItem.SubItems(2)
               If var_cantidad_posible > 0 Then
                  var_cantidad_Salida = 1
                  txt_cantidad = var_cantidad_Salida
                  txt_cantidad.Enabled = True
                  txt_cantidad.SetFocus
               Else
                  MsgBox "No es posible seleccionar mas de este artículo", vbOKOnly, "ATENCION"
               End If
            Else
               txt_codigo = ""
               rs.Close
               MsgBox "El artículo no se encuentra dentro de la salida", vbOKOnly, "ATENCION"
            End If
         End If
      Else
         MsgBox "Error en código", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_cliente = ""
      txt_nombre_cliente = ""
      txt_rfc = ""
      txt_descuentos = ""
      var_clave_establecimiento = txt_establecimiento
      If Trim(txt_establecimiento) <> "" Then
         txt_cliente.Enabled = True
         txt_nombre_cliente.Enabled = True
      Else
         txt_cliente.Enabled = False
         txt_nombre_cliente.Enabled = False
      End If
      rsaux.Open "select * from vw_establecimientos where vcha_age_agente_id = '" + txt_clave_agente + "' and vcha_esb_Establecimiento_id = '" + txt_establecimiento + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      If Not rsaux.EOF Then
         txt_nombre_establecimiento = rsaux!VCHA_ESB_NOMBRE
         rsaux.Close
         txt_establecimiento.Enabled = False
         var_clave_establecimiento = txt_establecimiento
         txt_nombre_establecimiento.Enabled = False
         txt_cliente.Enabled = True
         cmb_clientes.Enabled = True
         txt_cliente.SetFocus
      Else
         MsgBox "El establecimiento no existe", vbOKOnly, "ATENCION"
         rsaux.Close
         txt_nombre_establecimiento.SetFocus
      End If
   End If
End Sub

Private Sub txt_numero_salida_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_numero_salida) <> "" Then
         Dim list_item As ListItem
         Dim var_contador As Integer
         var_contador = 0
         'rs.Open "select * from vw_facturacion_vistas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'DA' and inte_com_numero = " + txt_numero_salida, cnn, adOpenDynamic, adLockOptimistic
         rs.Open "select * from vw_facturacion_vistas where inte_com_numero = " + txt_numero_salida, cnn, adOpenDynamic, adLockOptimistic
         lv_existencias.ListItems.Clear
         If Not rs.EOF Then
            While Not rs.EOF
                  Set list_item = lv_existencias.ListItems.Add(, , rs!vcha_Art_articulo_id)
                  list_item.SubItems(1) = Trim(rs!vcha_art_nombre_español)
                  list_item.SubItems(2) = Format(rs!FLOA_COM_CANTIDAD_ENVIADA - rs!floa_com_cantidad_recibida, "###,###,##0.00")
                  list_item.SubItems(3) = Format(0, "###,###,##0.00")
                  list_item.SubItems(4) = Format(rs!FLOA_COM_CANTIDAD_ENVIADA - rs!floa_com_cantidad_recibida, "###,###,##0.00")
                  var_contador = var_contador + 1
                  rs.MoveNext:
            Wend
            rs.MoveFirst
            var_movimiento_dependencia = IIf(IsNull(rs!vcha_mov_movimiento_dependencia), "", rs!vcha_mov_movimiento_dependencia)
            var_folio_enviado = txt_numero_salida
            txt_clave_agente = rs!VCHA_AGE_AGENTE_ID
            var_clave_agente = rs!VCHA_AGE_AGENTE_ID
            txt_nombre_agente = rs!VCHA_AGE_NOMBRE
            var_almacen_origen = rs!VCHA_EMO_ALMACEN_DESTINO
            var_almacen_Destino = rs!VCHA_ALM_ALMACEN_ID
            var_referencia = rs!vcha_com_Referencia
            var_orden_surtido = rs!INTE_EMO_NUMERO_ORIGEN
            txt_codigo.Enabled = False
            txt_numero_salida.Enabled = False
            txt_nombre_titular.Enabled = True
            txt_titular.Enabled = True
            txt_titular.SetFocus
         Else
            MsgBox "El número de salida no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   Dim var_actualiza As Boolean
   Dim var_inserta As Boolean
   Dim bandera_suma As Boolean
   Dim var_cantidad As Double
   Dim var_precio_pedido As Double
   Dim var_catalogo As String
   Dim var_otorga_oferta As Boolean
   Dim var_numero_dias As Integer
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_DET_ORDEN_SURTIDO_M = New TB_DET_ORDEN_SURTIDO_M
   Set TB_ARCH_COMPARACION_M = New TB_ARCH_COMPARACION_M
   Dim var_numero_salida As Double
   var_numero_salida = txt_numero_salida
   If Trim(var_lista_precios) <> "" Then
      If Trim(var_clave_moneda) <> "" Then
         If Trim(txt_codigo.Text) <> "" Then
            ''' SE CALCULA EL PRECIO
            rs.Open "select * from vw_lista_precios_clientes where vcha_cli_clave_id = '" + txt_cliente + "' and vcha_lis_lista_id = '" + var_lista_precios + "' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            var_promocion_1 = 0
            var_promocion_2 = 0
            If Not rs.EOF Then
               var_precio_pedido = rs!floa_dli_precio
               var_catalogo = rs!vcha_cat_catalogo_id
               var_otorga_oferta = False
               
''
               If Not IsNull(rs!dtim_vig_fecha_fin) Then
                  var_numero_dias = Date - rs!dtim_vig_fecha_fin
                  var_otorga_oferta = True
               Else
                  var_otorga_oferta = False
               End If

''
               
               rs.Close
               rs.Open "select * from vw_descuentos_promociones_vigentes where vcha_can_canal_venta_id = '" + var_canal_venta + "' and vcha_art_Articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_promocion_1 = IIf(IsNull(rs!floa_dpr_desCuento), 0, rs!floa_dpr_desCuento)
                  var_precio_pedido = var_precio_pedido - (var_precio_pedido * (rs!floa_dpr_desCuento / 100))
                  rs.Close
               Else
                  rs.Close
                  If var_otorga_oferta = True Then
                     rs.Open "select * from tb_descuentos_catalogos where vcha_can_canal_venta_id = '" + var_canal_venta + "' and inte_des_limite_inferior <= " + Str(var_numero_dias) + " and inte_des_limite_superior >= " + Str(var_numero_dias), cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_promocion_2 = IIf(IsNull(rs!FLOA_DES_DESCUENTO), 0, rs!FLOA_DES_DESCUENTO)
                        var_precio_pedido = var_precio_pedido - (var_precio_pedido * (rs!FLOA_DES_DESCUENTO / 100))
                     End If
                     rs.Close
                  End If
               End If
                 ''' FIN DEL CALCULO DEL PRECIO
               var_precio = var_precio_pedido
               bandera_suma = False
               cnn.BeginTrans
               If var_primera_vez = True Then
                  var_inserta = False
                  var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, var_numero_salida, txt_cliente, "", var_almacen_origen, "", "", var_clave_usuario_global, fun_NombrePc, 0, "", "", txt_establecimiento, "B", var_clave_titular, var_clave_agente, var_descuento_1, var_descuento_2, var_descuento_3, var_clave_moneda, var_tipo_Cambio)
                  var_numero_folio = var_numero_folio_regreso
                  txt_folio = var_numero_folio
                  var_primera_vez = False
               End If
               rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  rsaux.Close
                  valor = txt_codigo
                  Set itmfound = lv_existencias.findItem(valor, lvwText, , lvwPartial)
                  itmfound.EnsureVisible
                  itmfound.Selected = True
                  var_cantidad_posible = lv_existencias.selectedItem.SubItems(2)
                  lv_existencias.selectedItem.SubItems(3) = Format(lv_existencias.selectedItem.SubItems(3) + var_cantidad_Salida, "###,###,##0.00")
                  lv_existencias.selectedItem.SubItems(4) = Format(lv_existencias.selectedItem.SubItems(4) - var_cantidad_Salida, "###,###,##0.00")
                  bandera_suma = True
                  var_actualiza = TB_ARCH_COMPARACION_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_movimiento_dependencia, var_numero_salida, "G", txt_clave_agente, txt_codigo, var_cantidad_Salida, var_referencia)
               Else
                  MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                  rsaux.Close
               End If
               If bandera_suma = True Then
                  Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = " + txt_codigo
                  rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_inserta = False
                     var_inserta = TB_TEMPORAL_SALIDAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_Salida)
                     rs.Close
                  Else
                     var_inserta = False
                     var_inserta = TB_TEMPORAL_SALIDAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_Salida, var_costo, var_precio, "0", var_promocion_1, var_promocion_2)
                     rs.Close
                  End If
                  bandera_suma = False
                  cnn.CommitTrans
               End If
            Else
               MsgBox "El artículo no se encuentra dentro de la lista de precios del cliente", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El cliente no tiene una moneda relacionada", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El cliente no tiene relacionada una lista de precios", vbOKOnly, "ATENCION"
      End If
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_rfc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_titular_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct * from vw_titulares_1 where vcha_age_agente_id = '" + txt_clave_agente + "' order by vcha_tit_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tit_titular_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TITULARES"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_titular_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      rsaux.Open "select * from vw_titulares_1 where vcha_tit_titular_id = '" + txt_titular + "'", cnn, adOpenDynamic, adLockBatchOptimistic
      If Not rsaux.EOF Then
         txt_nombre_titular = rsaux!VCHA_TIT_NOMBRE
         rsaux.Close
         var_clave_titular = txt_titular
         txt_titular.Enabled = False
         txt_nombre_titular.Enabled = False
         txt_establecimiento.Enabled = True
         txt_nombre_establecimiento.Enabled = True
         txt_establecimiento.SetFocus
      Else
         MsgBox "El establecimiento no existe", vbOKOnly, "ATENCION"
         rsaux.Close
         txt_nombre_establecimiento.Enabled = True
         txt_nombre_establecimiento.SetFocus
      End If
   End If
End Sub

Private Sub txt_titular_LostFocus()
   If txt_titular <> "" Then
      txt_nombre_establecimiento.Enabled = True
      txt_establecimiento.Enabled = True
      txt_nombre_titular.Enabled = False
      txt_titular.Enabled = False
      txt_nombre_cliente.Enabled = False
      txt_cliente.Enabled = False
   Else
      txt_nombre_titular.Enabled = True
      txt_titular.Enabled = True
      txt_nombre_establecimiento.Enabled = False
      txt_establecimiento.Enabled = False
      txt_nombre_cliente.Enabled = False
      txt_cliente.Enabled = False
   End If
End Sub
