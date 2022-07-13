VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcargapedidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Pedidos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11700
   Icon            =   "frmcargapedidos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11700
   Begin VB.Frame frm_agentes 
      Height          =   3630
      Left            =   1305
      TabIndex        =   47
      Top             =   225
      Width           =   6075
      Begin VB.Frame Frame4 
         Height          =   60
         Left            =   15
         TabIndex        =   57
         Top             =   705
         Width           =   6045
      End
      Begin VB.CommandButton cmd_aceptar_pedidos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         Picture         =   "frmcargapedidos.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton cmd_cancelar_pedidos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmcargapedidos.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3780
         Picture         =   "frmcargapedidos.frx":0B5E
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3120
         Picture         =   "frmcargapedidos.frx":0D74
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Marcar (Enter)"
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3450
         Picture         =   "frmcargapedidos.frx":0FBE
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton Command9 
         Height          =   315
         Left            =   2460
         Picture         =   "frmcargapedidos.frx":1090
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2790
         Picture         =   "frmcargapedidos.frx":1192
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   375
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_agentes 
         Height          =   2775
         Left            =   45
         TabIndex        =   49
         Top             =   795
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1605
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Agente"
            Object.Width           =   8380
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Tipo"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000D&
         Caption         =   " Agentes"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   48
         Top             =   120
         Width           =   6000
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmcargapedidos.frx":13A8
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Pedido anterior"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_pedio_anterior 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   750
      Picture         =   "frmcargapedidos.frx":14AA
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Pedido anterior"
      Top             =   45
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame frm_carga_archivos 
      Height          =   3660
      Left            =   495
      TabIndex        =   3
      Top             =   450
      Width           =   8115
      Begin VB.CommandButton cmd_agrega_uno 
         Height          =   300
         Left            =   3885
         Picture         =   "frmcargapedidos.frx":15AC
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1545
         Width           =   315
      End
      Begin VB.CommandButton cmd_agrega_todos 
         Height          =   300
         Left            =   3885
         Picture         =   "frmcargapedidos.frx":16AE
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1905
         Width           =   315
      End
      Begin VB.CommandButton cmd_quita_todos 
         Height          =   300
         Left            =   3885
         Picture         =   "frmcargapedidos.frx":17B0
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2265
         Width           =   315
      End
      Begin VB.CommandButton cmd_quitar_uno 
         Height          =   300
         Left            =   3885
         Picture         =   "frmcargapedidos.frx":18B2
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2640
         Width           =   315
      End
      Begin VB.CommandButton cmd_cancelar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmcargapedidos.frx":19B4
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   450
         Width           =   330
      End
      Begin VB.CommandButton cmd_aceptar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         Picture         =   "frmcargapedidos.frx":1AFE
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   450
         Width           =   330
      End
      Begin VB.ListBox lst_origen 
         Height          =   2595
         Left            =   150
         TabIndex        =   8
         Top             =   930
         Width           =   3660
      End
      Begin VB.ListBox lst_archivos 
         Height          =   2595
         Left            =   4305
         TabIndex        =   7
         Top             =   930
         Width           =   3660
      End
      Begin VB.Frame Frame2 
         Height          =   60
         Left            =   0
         TabIndex        =   6
         Top             =   765
         Width           =   8115
      End
      Begin VB.FileListBox fil_archivos 
         Height          =   2625
         Left            =   135
         Pattern         =   "*.dbf"
         TabIndex        =   5
         Top             =   900
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Archivos que se cargaran para pedidos"
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   30
         TabIndex        =   4
         Top             =   120
         Width           =   8040
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2310
      Left            =   135
      TabIndex        =   12
      Top             =   930
      Width           =   11415
      Begin VB.TextBox txt_cantidad 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   225
         Width           =   2055
      End
      Begin VB.TextBox txt_importe_total 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   570
         Width           =   2055
      End
      Begin VB.TextBox txt_iva 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1590
         Width           =   2055
      End
      Begin VB.TextBox txt_importe_neto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1935
         Width           =   2040
      End
      Begin VB.TextBox txt_subimporte 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1260
         Width           =   2055
      End
      Begin VB.TextBox txt_descuento 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   915
         Width           =   2055
      End
      Begin VB.TextBox txt_descuento2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5490
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1935
         Width           =   1020
      End
      Begin VB.TextBox txt_descuento1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1935
         Width           =   1020
      End
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1590
         Width           =   5145
      End
      Begin VB.TextBox txt_establecimiento 
         Height          =   315
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1245
         Width           =   5145
      End
      Begin VB.TextBox txt_titular 
         Height          =   315
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   900
         Width           =   5145
      End
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   555
         Width           =   5145
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   210
         Width           =   1485
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   9195
         Top             =   270
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   20
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":1C48
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":2522
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":2DFC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":3398
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":3C74
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":454E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":4E28
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":4F3A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":504C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":515E
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":5270
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":5382
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":5494
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":59D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":5F18
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":602A
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":613C
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":624E
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":6358
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcargapedidos.frx":646A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   6795
         TabIndex        =   38
         Top             =   285
         Width           =   675
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Importe Total"
         Height          =   195
         Left            =   6795
         TabIndex        =   36
         Top             =   630
         Width           =   930
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "IVA:"
         Height          =   195
         Left            =   6795
         TabIndex        =   35
         Top             =   1650
         Width           =   300
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Importe Neto:"
         Height          =   195
         Left            =   6795
         TabIndex        =   34
         Top             =   1995
         Width           =   960
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Subimporte:"
         Height          =   195
         Left            =   6795
         TabIndex        =   33
         Top             =   1320
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Descuento:"
         Height          =   195
         Left            =   6795
         TabIndex        =   32
         Top             =   975
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Descuento por Pago Correcto:"
         Height          =   195
         Left            =   3270
         TabIndex        =   31
         Top             =   1980
         Width           =   2160
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Descuento por Volumen:"
         Height          =   195
         Left            =   150
         TabIndex        =   30
         Top             =   1980
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   150
         TabIndex        =   29
         Top             =   1650
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Left            =   150
         TabIndex        =   28
         Top             =   1305
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   615
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   270
         Width           =   600
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmcargapedidos.frx":657C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ver Archivos Alt + V"
      Top             =   45
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmcargapedidos.frx":667E
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11160
      Picture         =   "frmcargapedidos.frx":6780
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   45
      Width           =   330
   End
   Begin MSComctlLib.ListView lv_pedidos 
      Height          =   3855
      Left            =   150
      TabIndex        =   0
      Top             =   3345
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Numero"
         Object.Width           =   1605
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Agente"
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Titular"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Establecimiento"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Cliente     "
         Object.Width           =   3695
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Cantidad"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Importe     "
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Descuento 1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Importe Descuento 1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Descuento 2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Importe Descuento 2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Descuento 3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Importe Descuento 3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "IVA"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Importe IVA"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8760
      Top             =   -60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcargapedidos.frx":6DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcargapedidos.frx":7694
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcargapedidos.frx":7F6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcargapedidos.frx":850A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcargapedidos.frx":8DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcargapedidos.frx":96C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcargapedidos.frx":9F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcargapedidos.frx":A0AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcargapedidos.frx":A1BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcargapedidos.frx":A2D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcargapedidos.frx":A3E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcargapedidos.frx":A4F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcargapedidos.frx":A606
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcargapedidos.frx":AB48
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   60
      TabIndex        =   1
      Top             =   285
      Width           =   11460
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   105
      TabIndex        =   2
      Top             =   495
      Width           =   11685
   End
End
Attribute VB_Name = "frmcargapedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tabla As ADODB.Connection
Dim var_Archivo As String
Dim var_tabla_correcta As Boolean
Dim var_opcion_carga As Integer
Dim var_tipo_pedido As String
Dim var_resurtible As Integer
Dim var_especiales As Integer
Dim var_descuento1 As Variant
Dim var_descuento2 As Variant
Dim var_descuento3 As Variant
Dim var_titular As String
Dim var_cliente As String
Dim var_agente As String
Dim var_establecimiento As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_almacen As String


Private Sub llenar_lista()
   Dim var_descuento_1 As Double
   Dim var_descuento_2 As Double
   Dim var_descuento_3 As Double
   Dim var_importe_descuento_1 As Double
   Dim var_importe_descuento_2 As Double
   Dim var_importe_descuento_3 As Double
   Dim var_importe_total As Double
   Dim var_subimporte As Double
   If rs.State = 1 Then
      rs.Close
   End If
   cnn.CommandTimeout = 360
   var_dia = CStr(Day(Date))
   var_mes = CStr(Month(Date))
   var_año = CStr(Year(Date))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes)) = 1 Then
      var_mes = "0" + var_mes
   End If
   var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"

   rs.Open "select * from vw_suma_pedidos where dtim_ped_fecha >= " + var_fecha_fin + " and dtim_ped_fecha <= " + var_fecha_fin + "+1", cnn, adOpenDynamic, adLockOptimistic
   lv_pedidos.ListItems.Clear
   Dim list_item As ListItem
   While Not rs.EOF
      If Year(rs!dtim_ped_fecha) = Year(Date) And Month(rs!dtim_ped_fecha) = Month(Date) And Day(rs!dtim_ped_fecha) = Day(Date) Then
         Set list_item = lv_pedidos.ListItems.Add(, , rs!inte_ped_numero)
         list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
         list_item.SubItems(2) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
         list_item.SubItems(3) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
         list_item.SubItems(4) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         If IsNull(rs!Cantidad) Then
            list_item.SubItems(5) = Format(0, "###,###,##0.00")
         Else
            list_item.SubItems(5) = Format(rs!Cantidad, "###,###,##0.00")
         End If
         If IsNull(rs!Importe) Then
            list_item.SubItems(6) = Format(0, "###,###,##0.00")
            var_importe_total = 0
         Else
            list_item.SubItems(6) = Format(rs!Importe, "###,###,##0.00")
            var_importe_total = rs!Importe
         End If
         list_item.SubItems(7) = IIf(IsNull(rs!floa_ped_descuento_1), 0, rs!floa_ped_descuento_1)
         var_descuento_1 = IIf(IsNull(rs!floa_ped_descuento_1), 0, rs!floa_ped_descuento_1)
         If var_descuento_1 > 0 Then
            var_importe_descuento_1 = var_importe_total * (var_descuento_1 / 100)
         Else
            var_importe_descuento_1 = 0
         End If
         list_item.SubItems(8) = var_importe_descuento_1
         list_item.SubItems(9) = IIf(IsNull(rs!floa_ped_descuento_2), 0, rs!floa_ped_descuento_2)
         var_descuento_2 = IIf(IsNull(rs!floa_ped_descuento_2), 0, rs!floa_ped_descuento_2)
         If var_descuento_2 > 0 Then
            var_importe_descuento_2 = (var_importe_total - var_importe_descuento_1) * (var_descuento_2 / 100)
         Else
            var_importe_descuento_2 = 0
         End If
         list_item.SubItems(10) = var_importe_descuento_2
         list_item.SubItems(11) = IIf(IsNull(rs!floa_ped_Descuento_3), 0, rs!floa_ped_Descuento_3)
         var_descuento_3 = IIf(IsNull(rs!floa_ped_Descuento_3), 0, rs!floa_ped_Descuento_3)
         If var_descuento_3 > 0 Then
            var_importe_descuento_3 = (var_importe_total - var_importe_descuento_1 - var_importe_descuento_2) * (var_descuento_3 / 100)
         Else
            var_importe_descuento_3 = 0
         End If
         list_item.SubItems(12) = var_importe_descuento_3
         list_item.SubItems(13) = IIf(IsNull(rs!FLOA_TPE_IVA), 0, rs!FLOA_TPE_IVA)
         var_iva = IIf(IsNull(rs!FLOA_TPE_IVA), 0, rs!FLOA_TPE_IVA)
         If var_iva > 0 Then
            var_subimporte = var_importe_total - var_importe_descuento_1 - var_importe_descuento_2 - var_importe_descuento_3
            list_item.SubItems(14) = (var_subimporte * (1 + (var_iva / 100))) - var_subimporte
         Else
            list_item.SubItems(14) = 0
         End If
      End If
      rs.MoveNext:
   Wend
   rs.Close
   var_n = lv_pedidos.ListItems.Count
   var_numero_renglones = lv_pedidos.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_pedidos.ColumnHeaders(6).Width = 950
   Else
      lv_pedidos.ColumnHeaders(6).Width = 1150
   End If
End Sub
Private Sub cmd_aceptar_Click()
   Dim ok As Boolean
   Dim var_codigo As String
   Dim var_dias_condiciones As Integer
   Dim var_dias_caducidad As Integer
   Dim var_precio_pedido As Double
   Dim var_canal_venta As String
   Dim var_lista_precios As String
   Dim var_catalogo As String
   Dim var_numero_dias As Integer
   Dim var_otorga_oferta As Boolean
   Dim var_can_canal_venta As String
   Dim var_posible As Boolean
   Dim var_posible_moneda As Boolean
   Dim var_clave_moneda As String
   Dim var_sugerido As Integer
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   
   Set TB_ENC_PEDIDOS_I = New TB_ENC_PEDIDOS_I
   Set TB_DETALLE_PEDIDOS_I = New TB_DETALLE_PEDIDOS_I
   si = MsgBox("¿Se cargaran los pedidos?", vbYesNo, "ATENCION")
   If si = 6 Then
      n = lst_archivos.ListCount
      For i = 0 To n - 1
         lst_archivos.ListIndex = i
         var_Archivo = lst_archivos.List(i)
         var_bien = 1
         If var_bien = 1 Then
            Cadena = "select titular, cvecliente, establecim, numero, fecha, codigo, cantidad, precio, sugerido from " + Trim(var_Archivo)
            rs.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
            var_posible = True
            While Not rs.EOF
               var_codigo = Trim(rs!codigo)
               var_cliente = Trim(rs!cvecliente)
               rsaux1.Open "select * from tb_clientes where vcha_cli_clave_id = '" + var_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
               var_clave_moneda = IIf(IsNull(rsaux1!vcha_mon_moneda_id), "", rsaux1!vcha_mon_moneda_id)
               var_lista_precios = IIf(IsNull(rsaux1!vcha_LIS_LISTA_iD), "", rsaux1!vcha_LIS_LISTA_iD)
               rsaux1.Close
               rsaux1.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  var_codigo = rsaux1!vcha_Art_Articulo_id
               End If
               rsaux1.Close
               rsaux1.Open "select * from vw_lista_precios_clientes where vcha_cli_clave_id = '" + var_cliente + "' and vcha_lis_lista_id = '" + var_lista_precios + "' and vcha_art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux1.EOF Then
                  var_posible = False
               End If
               rsaux1.Close
               rs.MoveNext
            Wend
            rs.Close
            If Trim(var_clave_moneda) <> "" Then
               If var_posible = True Then
                  Cadena = "select distinct titular,establecim,cvecliente,numero,sugerido from " + Trim(var_Archivo)
                  rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
                  While Not rsaux.EOF
                     rsaux2.Open "select * from tb_encabezado_pedidos where vcha_cli_clave_id = '" + Trim(rsaux!cvecliente) + "' and INTE_PED_REFERENCIA = " + Str(rsaux!NUMERO), cnn, adOpenDynamic, adLockOptimistic
                     If rsaux2.EOF Then
                        Cadena = "select titular,cvecliente,establecim,numero,fecha,codigo,cantidad,precio, sugerido from " + Trim(var_Archivo) + " where  cvecliente = '" + Trim(rsaux!cvecliente) + "' and numero = " + Str(rsaux!NUMERO)
                        'cnn.BeginTrans
                        'rs.Open "SELECT * FROM VW_MAXIMO_PEDIDO", cnn, adOpenDynamic, adLockOptimistic
                        'If Not rs.EOF Then
                        '   If IsNull(rs(0).Value) Then
                        '      maximo_pedido = 1
                        '   Else
                        '      maximo_pedido = rs(0).Value + 1
                        '   End If
                        'Else
                        '   maximo_pedido = 1
                        'End If
                        'rs.Close
                        rs.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
                        var_titular = Trim(rs!titular)
                        var_sugerido = IIf(IsNull(rs!sugerido), 0, rs!sugerido)
                        var_cliente = Trim(rs!cvecliente)
                        var_establecimiento = Trim(rs!establecim)
                        rsaux3.Open "select * from vw_tipo_pedidos_1 where vcha_cli_clave_id = '" + rs!cvecliente + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_agente = rsaux3!VCHA_AGE_AGENTE_ID
                        var_tipo_pedido = rsaux3!char_tpe_tipo_pedido_id
                        var_resurtible = rsaux3!inte_tpe_resurtible
                        var_especiales = 0
                        var_descuento1 = rsaux3!floa_gac_Descuento_1
                        var_descuento2 = rsaux3!FLOA_GAC_DESCUENTO_2
                        var_descuento3 = rsaux3!floa_gac_descuento_3
                        var_dias_caducidad = rsaux3!inte_tpe_dias_caducidad
                        var_lista_precios = IIf(IsNull(rsaux3!vcha_LIS_LISTA_iD), "", rsaux3!vcha_LIS_LISTA_iD)
                        var_canal_venta = IIf(IsNull(rsaux3!vcha_can_canal_venta_id), "", rsaux3!vcha_can_canal_venta_id)
                        If IsNull(rsaux3!inte_pla_dias) Then
                           var_dias_condiciones = 0
                        Else
                           var_dias_condiciones = rsaux3!inte_pla_dias
                        End If
                        rsaux3.Close
                        ok = False
                        maximo_pedido = 0
                        ok = TB_ENC_PEDIDOS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, var_tipo_pedido, maximo_pedido, rs!NUMERO, Date, rs!fecha, var_agente, var_titular, var_cliente, var_establecimiento, var_resurtible, var_especiales, "I", var_descuento1, var_descuento2, var_descuento3, var_dias_condiciones, var_dias_caducidad, var_clave_usuario_global, fun_NombrePc, Date, var_clave_moneda, var_sugerido)
                        While Not rs.EOF
                           var_codigo = Trim(rs!codigo)
                           rsaux3.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux3.EOF Then
                              var_codigo = rsaux3!vcha_Art_Articulo_id
                           End If
                           rsaux3.Close
                           rsaux3.Open "select * from vw_lista_precios_clientes where vcha_cli_clave_id = '" + var_cliente + "' and vcha_lis_lista_id = '" + var_lista_precios + "' and vcha_art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux3.EOF Then
                              var_promocion_1 = 0
                              var_promocion_2 = 0
                              var_precio_pedido = rsaux3!floa_dli_precio
                              var_catalogo = rsaux3!vcha_cat_catalogo_id
                              var_otorga_oferta = False
                              If Not IsNull(rsaux3!dtim_vig_fecha_fin) Then
                                 var_numero_dias = Date - rsaux3!dtim_vig_fecha_fin
                                 var_otorga_oferta = True
                              Else
                                 var_otorga_oferta = False
                              End If
                              rsaux3.Close
                              rsaux3.Open "select * from vw_descuentos_promociones_vigentes where vcha_can_canal_venta_id = '" + var_canal_venta + "' and vcha_art_Articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 var_promocion_1 = IIf(IsNull(rsaux3!floa_dpr_desCuento), 0, rsaux3!floa_dpr_desCuento)
                                 var_precio_pedido = var_precio_pedido - (var_precio_pedido * (IIf(IsNull(rsaux3!floa_dpr_desCuento), 0, rsaux3!floa_dpr_desCuento) / 100))
                                 rsaux3.Close
                              Else
                                 rsaux3.Close
                                 If var_otorga_oferta = True Then
                                    rsaux3.Open "select * from tb_descuentos_catalogos where vcha_can_canal_venta_id = '" + var_canal_venta + "' and inte_des_limite_inferior <= " + Str(var_numero_dias) + " and inte_des_limite_superior >= " + Str(var_numero_dias), cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux3.EOF Then
                                       var_promocion_2 = IIf(IsNull(rsaux3!FLOA_DES_DESCUENTO), 0, rsaux3!FLOA_DES_DESCUENTO)
                                       var_precio_pedido = var_precio_pedido - (var_precio_pedido * (rsaux3!FLOA_DES_DESCUENTO / 100))
                                    End If
                                    rsaux3.Close
                                 End If
                              End If
''''
                              rsaux3.Open "select * from tb_detalle_pedidos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and inte_ped_numero = " + CStr(maximo_pedido) + " and vcha_art_articulo_id = '" + var_codigo + "' and floa_ped_precio = " + CStr(var_precio_pedido), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux3.Close
                                 rsaux3.Open "update tb_detalle_pedidos set floa_ped_cantidad = floa_ped_cantidad + " + CStr(rs!Cantidad) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and inte_ped_numero = " + CStr(maximo_pedido) + " and vcha_art_articulo_id = '" + var_codigo + "' and floa_ped_precio = " + CStr(var_precio_pedido), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 rsaux3.Close
                                 ok = TB_DETALLE_PEDIDOS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, maximo_pedido, var_codigo, var_precio_pedido, rs!Cantidad, 0, var_promocion_1, var_promocion_2, "P")
                              End If
                              
                           Else
                              MsgBox "El artículo " + var_codigo + " no se encuentra en la lista de precios relacionada al cliente", vbOKOnly, "ATENCION"
                              rsaux3.Close
                           End If
                           rs.MoveNext
                        Wend
                        rs.Close
                        'cnn.CommitTrans
                     End If
                     rsaux2.Close
                     rsaux.MoveNext
                  Wend
                  rsaux.Close
               Else
                  MsgBox "El archivo " + var_Archivo + " no puede ser cargado porque tiene artículos que no estan asignados a una lista de precios del cliente", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El cliente no tiene una moneda relacionada", vbOKOnly, "ATENCION"
            End If
         End If
      Next
      MsgBox "Se a terminado el proceso de carga de los pedidos", vbOKOnly, "ATENCION"
      Call llenar_lista
   Else
      MsgBox "Se a cancelado la carga de los pedidos", vbOKOnly, "ATENCION"
   End If
   frm_carga_archivos.Visible = False
End Sub

Private Sub cmd_aceptar_pedidos_Click()
   Dim ok As Boolean
   Dim var_codigo As String
   Dim var_dias_condiciones As Integer
   Dim var_dias_caducidad As Integer
   Dim var_precio_pedido As Double
   Dim var_canal_venta As String
   Dim var_lista_precios As String
   Dim var_catalogo As String
   Dim var_numero_dias As Double
   Dim var_otorga_oferta As Boolean
   Dim var_can_canal_venta As String
   Dim var_posible As Boolean
   Dim var_posible_moneda As Boolean
   Dim var_clave_moneda As String
   Dim var_sugerido As Integer
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   Dim var_ruta_agente As String
   
   Dim var_posible_agentes As Boolean
   Dim var_posible_cliente As Boolean
   Dim var_posible_establecimiento As Boolean
   Dim var_posible_titular As Boolean
   Dim var_cadena_agentes As String
   Dim var_cadena_clientes As String
   Dim var_cadena_titulares As String
   Dim var_cadena_establecimientos As String
   Dim var_bien As Boolean
   Dim var_contador_agentes As Double
   Dim var_veces As Double
   Dim var_durar As String
   Set TB_ENC_PEDIDOS_I = New TB_ENC_PEDIDOS_I
   Set TB_DETALLE_PEDIDOS_I = New TB_DETALLE_PEDIDOS_I
   si = MsgBox("¿Se cargaran los pedidos de los agentes?", vbYesNo, "ATENCION")
   var_z = 0
   If si = 6 Then
      
      var_durar = CStr(Now)
      var_contador_agentes = lv_agentes.ListItems.Count
      var_cadena_agentes = ""
      var_cadena_clientes = ""
      var_cadena_establecimientos = ""
      var_cadena_titulares = ""
      var_posible_agentes = True
      var_posible_cliente = True
      var_posible_titular = True
      var_posible_establecimiento = True
      rs.Open "DELETE FROM TB_TEMPORAL_CARGAR_PEDIDOS", cnn, adOpenDynamic, adLockOptimistic
      var_veces = 1
      var_bien = True
      var_z = 1
      While var_z <= lv_agentes.ListItems.Count
seguir:
            If Err.Number <> 0 Then
               Resume
            End If
             'On Error GoTo HELL
          lv_agentes.ListItems(var_z).Selected = True
          If lv_agentes.selectedItem.SubItems(2) = "*" Then
             var_veces = var_veces
             var_veces = var_veces
             rsaux4.Open "select * from tb_agentes where vcha_age_agente_id = '" + lv_agentes.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
             var_ruta_agente = IIf(IsNull(rsaux4!VCHA_AGE_RUTA_ARCHIVOS), "", rsaux4!VCHA_AGE_RUTA_ARCHIVOS)
             rsaux4.Close
             If Trim(var_ruta_agente) <> "" Then
                Frmmenu2.StatusBar1.Panels(1) = "Revisando la carpeta del agente " + lv_agentes.selectedItem.SubItems(1)
                var_bien = True
                Set var_tabla = CreateObject("ADODB.connection")
                var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta_agente + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
                If Me.lv_agentes.selectedItem.SubItems(3) = "TDA" Then
                   rsaux4.Open "select * from tb_clientes where vcha_age_agente_id = '" + lv_agentes.selectedItem + "' and inte_cli_cliente_pedido_tienda = 1", cnn, adOpenDynamic, adLockOptimistic
                   If Not rsaux4.EOF Then
                      var_clave_cliente = IIf(IsNull(rsaux4!vcha_cli_clave_id), "", rsaux4!vcha_cli_clave_id)
                      var_clave_titular = IIf(IsNull(rsaux4!vcha_tit_titular_id), "", rsaux4!vcha_tit_titular_id)
                   Else
                      var_clave_cliente = ""
                      var_clave_titular = ""
                   End If
                   rsaux4.Close
                   rsaux4.Open "select * from tb_detalle_establecimientos where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                   If Not rsaux4.EOF Then
                      var_clave_establecimiento = IIf(IsNull(rsaux4!vcha_ESB_ESTABLECIMIENTO_id), "", rsaux4!vcha_ESB_ESTABLECIMIENTO_id)
                   Else
                      var_clave_establecimiento = ""
                   End If
                   rsaux4.Close
                   Cadena = "select '" + var_clave_titular + "' as cvetitular, '" + var_clave_cliente + "' as cvecliente, '" + var_clave_establecimiento + "' cvetienda, 0 as numpedido, date() as fechapedido, cveestilo, round(cant1,2) as canpedi1, 0 as precio, '' tablcomive from datos"
                Else
                   Cadena = "select a.cvetitular, a.cvecliente, a.cvetienda, a.numpedido, a.fechapedid, a.cveestilo, a.canpedi1, a.precio, b.tablcomive from detalle a, general b where a.cvecliente = b.cvecliente and a.numpedido =  b.numpedido and a.cvetitular = b.cvetitular"
                End If
                rs.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
                var_posible = True
                While Not rs.EOF
                      var_codigo = Trim(rs!cveestilo)
                      var_cliente = Trim(rs!cvecliente)
                      var_establecimiento = Trim(rs!cvetienda)
                      var_titular = Trim(rs!cvetitular)
                      var_sugerido = 0
                      If Trim(rs!tablcomive) = "P" Then
                         var_sugerido = 0
                      End If
                      If Trim(rs!tablcomive) = "S" Then
                         var_sugerido = 1
                      End If
                      
                      var_multiplica_pedido = 1
                      var_plazo_pedido = 0
                      rsaux9.Open "select * from tb_codigos_partir_pedidos where vcha_tem_codigo = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                      If Not rsaux9.EOF Then
                         var_multiplica_pedido = IIf(IsNull(rsaux9!inte_tem_numero), 1, rsaux9!inte_tem_numero)
                         var_plazo_pedido = IIf(IsNull(rsaux9!inte_tem_plazo), 0, rsaux9!inte_tem_plazo)
                      Else
                         var_multiplica_pedido = 1
                         var_plazo_pedido = 0
                      End If
                      rsaux9.Close
                      rsaux.Open "insert into tb_temporal_Cargar_pedidos (inte_ped_numero, vcha_tit_titular_anterior_id, vcha_cli_clave_anterior_id, vcha_esb_establecimiento_anterior_id, vcha_Art_articulo_anterior_id, INTE_PED_CANTIDAD, INTE_PED_APLICADO,INTE_PED_SUGERIDO, VCHA_AGE_AGENTE_ID,inte_ped_plazo) VALUES (" + Trim(CDbl(rs!NUMPEDIDO * var_multiplica_pedido)) + ", '" + Trim(rs!cvetitular) + "','" + Trim(rs!cvecliente) + "','" + Trim(rs!cvetienda) + "', '" + Trim(rs!cveestilo) + "', " + CStr(rs!canpedi1) + ",0," + CStr(var_sugerido) + ",'" + lv_agentes.selectedItem + "', " + CStr(var_plazo_pedido) + ")", cnn, adOpenDynamic, adLockOptimistic
                      
                      rs.MoveNext
                Wend
                rs.Close
             Else
                var_posible_agentes = False
                var_cadena_agentes = var_cadena_agentes + IIf(IsNull(lv_agentes.selectedItem), "", lv_agentes.selectedItem)
             End If
             var_bien = True
          End If
HELL:
     var_z = var_z + 1
     If var_z > lv_agentes.ListItems.Count Then
        GoTo salir:
     End If
     lv_agentes.ListItems(var_z).Selected = True
     If lv_agentes.selectedItem.SubItems(2) = "*" Then
        Frmmenu2.StatusBar1.Panels(1) = "Revisando la carpeta del agente " + lv_agentes.selectedItem.SubItems(1)
        If rsaux4.State = 1 Then
           rsaux4.Close
        End If
        rsaux4.Open "select * from tb_agentes where vcha_age_agente_id = '" + lv_agentes.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
        var_ruta_agente = IIf(IsNull(rsaux4!VCHA_AGE_RUTA_ARCHIVOS), "", rsaux4!VCHA_AGE_RUTA_ARCHIVOS)
        rsaux4.Close
        If var_ruta_agente <> "" Then
           Set var_tabla = Nothing
           Set var_tabla = CreateObject("ADODB.connection")
           var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta_agente + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
           If Me.lv_agentes.selectedItem.SubItems(3) = "TDA" Then
              rsaux4.Open "select * from tb_clientes where vcha_age_agente_id = '" + lv_agentes.selectedItem + "'  and inte_cli_cliente_pedido_tienda = 1", cnn, adOpenDynamic, adLockOptimistic
              If Not rsaux4.EOF Then
                 var_clave_cliente = IIf(IsNull(rsaux4!vcha_cli_clave_id), "", rsaux4!vcha_cli_clave_id)
                 var_clave_titular = IIf(IsNull(rsaux4!vcha_tit_titular_id), "", rsaux4!vcha_tit_titular_id)
              Else
                 var_clave_cliente = ""
                 var_clave_titular = ""
              End If
              rsaux4.Close
              rsaux4.Open "select * from tb_detalle_establecimientos where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
              If Not rsaux4.EOF Then
                 var_clave_establecimiento = IIf(IsNull(rsaux4!vcha_ESB_ESTABLECIMIENTO_id), "", rsaux4!vcha_ESB_ESTABLECIMIENTO_id)
              Else
                 var_clave_establecimiento = ""
              End If
              rsaux4.Close
              Cadena = "select '" + var_clave_titular + "' as cvetitular, '" + var_clave_cliente + "' as cvecliente, '" + var_clave_establecimiento + "' cvetienda, 0 as numpedido, date() as fechapedido, cveestilo, cant1 as canpedi1, 0 as precio, '' tablcomive from datos"
           Else
              Cadena = "select a.cvetitular, a.cvecliente, a.cvetienda, a.numpedido, a.fechapedid, a.cveestilo, a.canpedi1, a.precio, b.tablcomive from detalle a, general b where a.cvecliente = b.cvecliente and a.numpedido =  b.numpedido and a.cvetitular = b.cvetitular"
           End If
        End If
     End If
     If rs.State = 1 Then
        rs.Close
     End If
      Wend
salir:

      var_cadena_agentes = ""
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select * from TB_TEMPORAL_CARGAR_PEDIDOS where floa_ped_precio = 0", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Set reporte = appl.OpenReport(App.Path + "\rep_tem_cargar_pedidos.rpt")
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
      
         reporte.RecordSelectionFormula = "{VW_TEMPORAL_CARGAR_PEDIDOS.FLOA_PED_PRECIO} = 0"
         frmvistasprevias.cr.ReportSource = reporte
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de artículos que no se encuentran en la lista de precios"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         If rsaux4.State = 1 Then
            rsaux4.Close
         End If
         rsaux4.Open "select distinct vcha_age_agente_id FROM TB_TEMPORAL_CARGAR_PEDIDOS where inte_ped_aplicado = 0 and floa_ped_precio = 0", cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux4.EOF
               var_cadena_agentes = var_cadena_agentes + " " + Trim(IIf(IsNull(rsaux4!VCHA_AGE_AGENTE_ID), "", rsaux4!VCHA_AGE_AGENTE_ID))
               'rsaux.Open "delete from tb_temporal_cargar_pedidos where vcha_age_agente_id = '" + IIf(IsNull(rsaux4!vcha_age_agente_id), "", rsaux4!vcha_age_agente_id) + "' and inte_ped_aplicado = 0", cnn, adOpenDynamic, adLockOptimistic
               rsaux.Open "update tb_temporal_cargar_pedidos set inte_ped_cantidad = 0  where vcha_age_agente_id = '" + IIf(IsNull(rsaux4!VCHA_AGE_AGENTE_ID), "", rsaux4!VCHA_AGE_AGENTE_ID) + "' and inte_ped_aplicado = 0 and floa_ped_precio = 0"
               rsaux4.MoveNext
         Wend
         rsaux4.Close
         If Trim(var_cadena_agentes) <> "" Then
            MsgBox "Existen problemas con los siguientes agentes " + var_cadena_agentes, vbOKOnly, "ATENCION"
         End If
      End If
      rs.Close
      rs.Open "exec SP_CARGA_TEMPORAL_PEDIDOS '" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen + "','" + var_clave_usuario_global + "', '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
      Call llenar_lista
      var_durar = var_durar + " " + CStr(Now)
      MsgBox var_durar, vbOKOnly, "atencion"
   End If
   frm_agentes.Visible = False
End Sub

Private Sub cmd_agrega_todos_Click()
   n = lst_origen.ListCount
   For i = 0 To n - 1
      var_Archivo = lst_origen.List(i)
      var_tabla_correcta = False
      var_opcion_carga = 2
      Call valida(var_Archivo)
      If var_tabla_correcta = True Then
         lst_archivos.AddItem lst_origen.List(i)
      End If
   Next
   m = lst_archivos.ListCount
   For i = 0 To m - 1
       var_Archivo = lst_archivos.List(i)
       n = lst_origen.ListCount
       For j = 0 To n - 1
           var_archivo_2 = lst_origen.List(j)
           If var_archivo_2 = var_Archivo Then
              lst_origen.RemoveItem (j)
           End If
       Next j
   Next i
End Sub

Private Sub cmd_agrega_uno_Click()
   var_opcion_carga = 1
   n = lst_origen.ListIndex
   If n >= 0 Then
      Dim var_Archivo As String
      n = lst_origen.ListIndex
      var_Archivo = lst_origen.List(n)
      var_tabla_correcta = False
      Call valida(var_Archivo)
      If var_tabla_correcta = True Then
         lst_archivos.AddItem lst_origen.List(n)
         lst_origen.RemoveItem (lst_origen.ListIndex)
      End If
   Else
      MsgBox "No se a seleccionado ningun archivo", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_cancelar_Click()
         frm_carga_archivos.Visible = False
End Sub

Private Sub cmd_cancelar_pedidos_Click()
   frm_agentes.Visible = False
End Sub

Private Sub cmd_imprimir_Click()
   Dim fecha1 As String
   Dim fecha2 As String
         Set reporte = appl.OpenReport(App.Path + "\rep_PEDIDos_2.rpt")
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         
         fecha1 = CDate(Date)
         reporte.RecordSelectionFormula = "Year ({VW_SUMA_PEDIDOS.DTIM_PED_FECHA}) = " + Str(Year(Date)) + " and Month ({VW_SUMA_PEDIDOS.DTIM_PED_FECHA}) = " + Str(Month(Date)) + " and Day ({VW_SUMA_PEDIDOS.DTIM_PED_FECHA})= " + Str(Day(Date))
         frmvistasprevias.cr.ReportSource = reporte
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Resumen de Pedidos"
         frmvistasprevias.Show 1
         Set reporte = Nothing
End Sub

Private Sub cmd_nuevo_Click()
   Dim fecha1 As String
   Dim fecha2 As String
   lst_origen.Clear
   rs.Open "select vcha_pri_ruta_pedidos from tb_principal", cnn, adOpenDynamic, adLockOptimistic
   ruta_pedidos = IIf(IsNull(rs!VCHA_PRI_RUTA_PEDIDOS), "", rs!VCHA_PRI_RUTA_PEDIDOS)
   rs.Close
   frm_carga_archivos.Visible = True
   fil_archivos.Path = ruta_pedidos
   n = fil_archivos.ListCount
   For i = 0 To n - 1
       fil_archivos.ListIndex = i
       lst_origen.AddItem fil_archivos.FileName
   Next
   lst_archivos.Clear
End Sub

Private Sub cmd_pedio_anterior_Click()
   Dim ok As Boolean
   Dim var_codigo As String
   Dim var_dias_condiciones As Integer
   Dim var_dias_caducidad As Integer
   Dim var_precio_pedido As Double
   Dim var_canal_venta As String
   Dim var_lista_precios As String
   Dim var_catalogo As String
   Dim var_numero_dias As Double
   Dim var_otorga_oferta As Boolean
   Dim var_can_canal_venta As String
   Dim var_posible As Boolean
   Dim var_posible_moneda As Boolean
   Dim var_clave_moneda As String
   Dim var_sugerido As Integer
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   Dim var_ruta_agente As String
   
   Dim var_posible_agentes As Boolean
   Dim var_posible_cliente As Boolean
   Dim var_posible_establecimiento As Boolean
   Dim var_posible_titular As Boolean
   Dim var_cadena_agentes As String
   Dim var_cadena_clientes As String
   Dim var_cadena_titulares As String
   Dim var_cadena_establecimientos As String
   Dim var_bien As Boolean
   Dim var_contador_agentes As Double
   Dim var_veces As Double
   Dim var_durar As String
   Set TB_ENC_PEDIDOS_I = New TB_ENC_PEDIDOS_I
   Set TB_DETALLE_PEDIDOS_I = New TB_DETALLE_PEDIDOS_I
   si = MsgBox("¿Se cargaran los pedidos de los agentes?", vbYesNo, "ATENCION")
   If si = 6 Then
      var_durar = CStr(Now)
      rsaux4.Open "select * from tb_agentes where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      var_contador_agentes = rsaux4.RecordCount
      var_cadena_agentes = ""
      var_cadena_clientes = ""
      var_cadena_establecimientos = ""
      var_cadena_titulares = ""
      var_posible_agentes = True
      var_posible_cliente = True
      var_posible_titular = True
      var_posible_establecimiento = True
      rs.Open "DELETE FROM TB_TEMPORAL_CARGAR_PEDIDOS WHERE INTE_PED_APLICADO = 0", cnn, adOpenDynamic, adLockOptimistic
      var_veces = 1
      var_bien = True
      While Not rsaux4.EOF
seguir:
            If Err.Number <> 0 Then
               Resume
            End If
            On Error GoTo HELL
            var_veces = var_veces
            var_veces = var_veces
            var_ruta_agente = IIf(IsNull(rsaux4!VCHA_AGE_RUTA_ARCHIVOS), "", rsaux4!VCHA_AGE_RUTA_ARCHIVOS)
            If Trim(var_ruta_agente) <> "" Then
               Frmmenu2.StatusBar1.Panels(1) = "Revisando la carpeta del agente " + rsaux4!VCHA_AGE_NOMBRE
               var_bien = True
               Set var_tabla = CreateObject("ADODB.connection")
               var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta_agente + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
               Cadena = "select a.cvetitular, a.cvecliente, a.cvetienda, a.numpedido, a.fechapedid, a.cveestilo, a.canpedi1, a.precio, b.tablcomive from detalle a, general b where a.cvecliente = b.cvecliente and a.numpedido =  b.numpedido and a.cvetitular = b.cvetitular"
               rs.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
               var_posible = True
               While Not rs.EOF
                     var_codigo = Trim(rs!cveestilo)
                     var_cliente = Trim(rs!cvecliente)
                     var_establecimiento = Trim(rs!cvetienda)
                     var_titular = Trim(rs!cvetitular)
                     var_sugerido = 0
                     If Trim(rs!tablcomive) = "P" Then
                        var_sugerido = 0
                     End If
                     If Trim(rs!tablcomive) = "S" Then
                        var_sugerido = 1
                     End If
                     'rsaux2.Open "select * from tb_temporal_cargar_pedidos where vcha_age_agente_id = '" + Trim(rsaux4!VCHA_AGE_AGENTE_ID) + "' and inte_ped_numero = " + CStr(rs!numpedido) + " and vcha_tit_titular_anterior_id = '" + Trim(rs!cvetitular) + "' and vcha_esb_establecimiento_anterior_id = '" + Trim(rs!cvetienda) + "' and vcha_cli_clave_anterior_id = '" + Trim(rs!cvecliente) + "'and vcha_art_articulo_anterior_id = '" + Trim(rs!cveestilo) + "'", cnn, adOpenDynamic, adLockOptimistic
                     'If rsaux2.EOF Then
                     '   Cadena = "insert into tb_temporal_Cargar_pedidos (vcha_age_agente_id, inte_ped_numero, vcha_tit_titular_anterior_id, vcha_cli_clave_anterior_id, vcha_esb_establecimiento_anterior_id, vcha_Art_articulo_anterior_id, INTE_PED_CANTIDAD, INTE_PED_APLICADO,INTE_PED_SUGERIDO) VALUES ('" + rsaux4!VCHA_AGE_AGENTE_ID + "'," + Trim(rs!numpedido) + ", '" + Trim(rs!cvetitular) + "','" + Trim(rs!cvecliente) + "','" + Trim(rs!cvetienda) + "', '" + Trim(rs!cveestilo) + "', " + CStr(rs!canpedi1) + ",0," + CStr(var_sugerido) + ")"
                        rsaux.Open "insert into tb_temporal_Cargar_pedidos (vcha_age_agente_id, inte_ped_numero, vcha_tit_titular_anterior_id, vcha_cli_clave_anterior_id, vcha_esb_establecimiento_anterior_id, vcha_Art_articulo_anterior_id, INTE_PED_CANTIDAD, INTE_PED_APLICADO,INTE_PED_SUGERIDO) VALUES ('" + rsaux4!VCHA_AGE_AGENTE_ID + "'," + Trim(rs!NUMPEDIDO) + ", '" + Trim(rs!cvetitular) + "','" + Trim(rs!cvecliente) + "','" + Trim(rs!cvetienda) + "', '" + Trim(rs!cveestilo) + "', " + CStr(rs!canpedi1) + ",0," + CStr(var_sugerido) + ")", cnn, adOpenDynamic, adLockOptimistic
                     'Else
                     '   If rsaux2!inte_ped_aplicado = 0 Then
                     '      rsaux.Open "update tb_temporal_Cargar_pedidos set INTE_PED_CANTIDAD = " + CStr(rs!canpedi1) + ", INTE_PED_APLICADO = 0 where vcha_age_agente_id = '" + rsaux4!VCHA_AGE_AGENTE_ID + "' and inte_ped_numero = " + CStr(rs!numpedido) + " and vcha_tit_titular_anterior_id = '" + Trim(rs!cvetitular) + "' and vcha_esb_establecimiento_anterior_id = '" + Trim(rs!cvetienda) + "' and vcha_cli_clave_anterior_id = '" + Trim(rs!cvecliente) + "'and vcha_art_articulo_anterior_id = '" + Trim(rs!cveestilo) + "'", cnn, adOpenDynamic, adLockOptimistic
                     '   End If
                     'End If
                     'rsaux2.Close
                     rs.MoveNext
               Wend
               rs.Close
            Else
               var_posible_agentes = False
               var_cadena_agentes = var_cadena_agentes + IIf(IsNull(rsaux4!VCHA_AGE_AGENTE_ID), "", rsaux4!VCHA_AGE_AGENTE_ID)
            End If
            var_bien = True
HELL:
            rsaux4.MoveNext
            If Not rsaux4.EOF Then
               Frmmenu2.StatusBar1.Panels(1) = "Revisando la carpeta del agente " + rsaux4!VCHA_AGE_NOMBRE
               var_ruta_agente = IIf(IsNull(rsaux4!VCHA_AGE_RUTA_ARCHIVOS), "", rsaux4!VCHA_AGE_RUTA_ARCHIVOS)
               If var_ruta_agente <> "" Then
                  Set var_tabla = Nothing
                  Set var_tabla = CreateObject("ADODB.connection")
                  var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta_agente + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
                  Cadena = "select a.cvetitular, a.cvecliente, a.cvetienda, a.numpedido, a.fechapedid, a.cveestilo, a.canpedi1, a.precio, b.tablcomive from detalle a, general b where a.cvecliente = b.cvecliente and a.numpedido =  b.numpedido and a.cvetitular = b.cvetitular"
               End If
            End If
      Wend
      rsaux4.Close
      var_cadena_agentes = ""
      rs.Open "select * from TB_TEMPORAL_CARGAR_PEDIDOS where floa_ped_precio = 0", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Set reporte = appl.OpenReport(App.Path + "\rep_tem_cargar_pedidos.rpt")
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
      
         reporte.RecordSelectionFormula = "{VW_TEMPORAL_CARGAR_PEDIDOS.FLOA_PED_PRECIO} = 0"
         frmvistasprevias.cr.ReportSource = reporte
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de artículos que no se encuentran en la lista de precios"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         rsaux4.Open "select distinct vcha_age_agente_id FROM TB_TEMPORAL_CARGAR_PEDIDOS where inte_ped_aplicado = 0 and floa_ped_precio = 0"
         While Not rsaux4.EOF
               var_cadena_agentes = var_cadena_agentes + " " + Trim(rsaux4!VCHA_AGE_AGENTE_ID)
               rsaux.Open "delete from tb_temporal_cargar_pedidos where vcha_age_agente_id = '" + rsaux4!VCHA_AGE_AGENTE_ID + "' and inte_ped_aplicado = 0", cnn, adOpenDynamic, adLockOptimistic
               rsaux4.MoveNext
         Wend
         rsaux4.Close
         If Trim(var_cadena_agentes) <> "" Then
            MsgBox "Existen problemas con los siguientes agentes " + var_cadena_agentes, vbOKOnly, "ATENCION"
         End If
      End If
      rs.Close
      rsaux4.Open "select distinct vcha_cli_clave_id, inte_ped_numero FROM TB_TEMPORAL_CARGAR_PEDIDOS where inte_ped_aplicado = 0", cnn, adOpenDynamic, adLockOptimistic
      While Not rsaux4.EOF
            rsaux2.Open "select * from tb_encabezado_pedidos where vcha_cli_clave_id = '" + Trim(rsaux4!vcha_cli_clave_id) + "' and INTE_PED_REFERENCIA = " + Str(rsaux4!inte_ped_numero), cnn, adOpenDynamic, adLockOptimistic
            If rsaux2.EOF Then
               Cadena = "select * FROM TB_TEMPORAL_CARGAR_PEDIDOS WHERE VCHA_CLI_CLAVE_ID = '" + Trim(rsaux4!vcha_cli_clave_id) + "' and INTE_PED_NUMERO = " + Str(rsaux4!inte_ped_numero) + " AND INTE_PED_APLICADO = 0"
               'cnn.BeginTrans
               rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
               var_titular = Trim(rs!vcha_tit_titular_id)
               var_sugerido = IIf(IsNull(rs!inte_ped_sugerido), 0, rs!inte_ped_sugerido)
              'var_sugerido = IIf(IsNull(rs!sugerido), 0, rs!sugerido)
               var_cliente = Trim(rs!vcha_cli_clave_id)
               var_establecimiento = Trim(rs!vcha_ESB_ESTABLECIMIENTO_id)
               rsaux3.Open "select * from vw_tipo_pedidos_1 where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
               var_agente = rsaux3!VCHA_AGE_AGENTE_ID
               var_tipo_pedido = rsaux3!char_tpe_tipo_pedido_id
               var_resurtible = rsaux3!inte_tpe_resurtible
               var_especiales = 0
               var_descuento1 = rsaux3!floa_gac_Descuento_1
               var_descuento2 = rsaux3!FLOA_GAC_DESCUENTO_2
               var_descuento3 = rsaux3!floa_gac_descuento_3
               var_clave_moneda = rsaux3!vcha_mon_moneda_id
               var_dias_caducidad = rsaux3!inte_tpe_dias_caducidad
               var_lista_precios = IIf(IsNull(rsaux3!vcha_LIS_LISTA_iD), "", rsaux3!vcha_LIS_LISTA_iD)
               var_canal_venta = IIf(IsNull(rsaux3!vcha_can_canal_venta_id), "", rsaux3!vcha_can_canal_venta_id)
               If IsNull(rsaux3!inte_pla_dias) Then
                  var_dias_condiciones = 0
               Else
                  var_dias_condiciones = rsaux3!inte_pla_dias
               End If
               rsaux3.Close
               ok = False
               maximo_pedido = 0
               ok = TB_ENC_PEDIDOS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, var_tipo_pedido, maximo_pedido, rs!inte_ped_numero, Date, Date, var_agente, var_titular, var_cliente, var_establecimiento, var_resurtible, var_especiales, "I", var_descuento1, var_descuento2, var_descuento3, var_dias_condiciones, var_dias_caducidad, var_clave_usuario_global, fun_NombrePc, Date, var_clave_moneda, var_sugerido)
               If var_tipo_pedido = "T" Then
                  rsaux3.Open "update tb_encabezado_pedidos set INTE_PED_AUTORIZO = 1, VCHA_PED_AUTORIZO = '" + var_clave_usuario_gloabal + "', DTIM_PED_AUTORIZO = getdate() where vcha_emp_empresa_id = '" + var_empresa + "' and inte_ped_numero = " + CStr(maximo_pedido), cnn, adOpenDynamic, adLockOptimistic
               End If
               While Not rs.EOF
                     var_codigo = Trim(rs!vcha_Art_Articulo_id)
                     rsaux3.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux3.EOF Then
                        var_codigo = rsaux3!vcha_Art_Articulo_id
                     End If
                     rsaux3.Close
                     rsaux3.Open "select * from vw_lista_precios_clientes where vcha_cli_clave_id = '" + var_cliente + "' and vcha_lis_lista_id = '" + var_lista_precios + "' and vcha_art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux3.EOF Then
                        var_promocion_1 = 0
                        var_promocion_2 = 0
                        var_precio_pedido = rsaux3!floa_dli_precio
                        var_catalogo = rsaux3!vcha_cat_catalogo_id
                        var_otorga_oferta = False
                        If Not IsNull(rsaux3!dtim_vig_fecha_fin) Then
                           var_numero_dias = Date - rsaux3!dtim_vig_fecha_fin
                           var_otorga_oferta = True
                        Else
                           var_otorga_oferta = False
                        End If
                        rsaux3.Close
                        rsaux3.Open "select * from vw_descuentos_promociones_vigentes where vcha_can_canal_venta_id = '" + var_canal_venta + "' and vcha_art_Articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           var_promocion_1 = IIf(IsNull(rsaux3!floa_dpr_desCuento), 0, rsaux3!floa_dpr_desCuento)
                           var_precio_pedido = var_precio_pedido - (var_precio_pedido * (IIf(IsNull(rsaux3!floa_dpr_desCuento), 0, rsaux3!floa_dpr_desCuento) / 100))
                           rsaux3.Close
                        Else
                           rsaux3.Close
                           If var_otorga_oferta = True Then
                              rsaux3.Open "select * from tb_descuentos_catalogos where vcha_can_canal_venta_id = '" + var_canal_venta + "' and inte_des_limite_inferior <= " + Str(var_numero_dias) + " and inte_des_limite_superior >= " + Str(var_numero_dias), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 var_promocion_2 = IIf(IsNull(rsaux3!FLOA_DES_DESCUENTO), 0, rsaux3!FLOA_DES_DESCUENTO)
                                 var_precio_pedido = var_precio_pedido - (var_precio_pedido * (rsaux3!FLOA_DES_DESCUENTO / 100))
                              End If
                              rsaux3.Close
                           End If
                        End If
                        rsaux3.Open "select * from tb_detalle_pedidos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and inte_ped_numero = " + CStr(maximo_pedido) + " and vcha_art_articulo_id = '" + var_codigo + "' and floa_ped_precio = " + CStr(var_precio_pedido), cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           rsaux3.Close
                           rsaux3.Open "update tb_detalle_pedidos set floa_ped_cantidad = floa_ped_cantidad + " + CStr(rs!Cantidad) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and inte_ped_numero = " + CStr(maximo_pedido) + " and vcha_art_articulo_id = '" + var_codigo + "' and floa_ped_precio = " + CStr(var_precio_pedido), cnn, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux3.Close
                           ok = TB_DETALLE_PEDIDOS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, maximo_pedido, var_codigo, var_precio_pedido, rs!INTE_PED_CANTIDAD, 0, var_promocion_1, var_promocion_2, "P")
                        End If
                     Else
                        var_promocion_1 = 0
                        var_promocion_2 = 0
                        var_precio_pedido = rs!FLOA_PED_PRECIO
                        ok = TB_DETALLE_PEDIDOS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, maximo_pedido, var_codigo, var_precio_pedido, rs!INTE_PED_CANTIDAD, 0, var_promocion_1, var_promocion_2, "P")
                        rsaux3.Close
                     End If
                     rs.MoveNext
               Wend
               rs.Close
               rs.Open "update TB_TEMPORAL_CARGAR_PEDIDOS set inte_ped_aplicado = 1 WHERE VCHA_CLI_CLAVE_ID = '" + Trim(rsaux4!vcha_cli_clave_id) + "' and INTE_PED_NUMERO = " + Str(rsaux4!inte_ped_numero), cnn, adOpenDynamic, adLockOptimistic
               'cnn.CommitTrans
            End If
            rsaux2.Close
            rsaux4.MoveNext
      Wend
      rsaux4.Close
      Call llenar_lista
      var_durar = var_durar + " " + CStr(Now)
      MsgBox var_durar, vbOKOnly, ""
   End If
End Sub

Private Sub cmd_quita_todos_Click()
   lst_origen.Clear
   fil_archivos.Path = ruta_pedidos
   n = fil_archivos.ListCount
   For i = 0 To n - 1
      fil_archivos.ListIndex = i
      lst_origen.AddItem fil_archivos.FileName
   Next
   lst_archivos.Clear
End Sub

Private Sub cmd_quitar_uno_Click()
   n = lst_archivos.ListIndex
   If n >= 0 Then
      Dim var_Archivo As String
      n = lst_archivos.ListIndex
      var_Archivo = lst_archivos.List(n)
      lst_origen.AddItem lst_archivos.List(n)
      lst_archivos.RemoveItem (lst_archivos.ListIndex)
   Else
      MsgBox "No se a seleccionado ningún archivo", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   lv_agentes.ListItems.Clear
   Dim list_item As ListItem
   rs.Open "select * from tb_agentes", cnn, adOpenDynamic, adLockOptimistic
   numero_items_catalogos = 0
   While Not rs.EOF
      Set list_item = lv_agentes.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      list_item.SubItems(2) = ""
      list_item.SubItems(3) = IIf(IsNull(rs!vcha_tag_tipoagente_id), "", rs!vcha_tag_tipoagente_id)
      rs.MoveNext:
      numero_items_catalogos = numero_items_catalogos + 1
   Wend
   rs.Close
   frm_agentes.Visible = True
   lv_agentes.SetFocus
End Sub

Private Sub Command10_Click()
   var_todos_lineas = 1
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_agentes.ListItems.Count
   For i = 1 To n
       lv_agentes.ListItems.item(i).SubItems(2) = "*"
       lv_agentes.ListItems.item(i).Bold = True
       lv_agentes.ListItems.item(i).ForeColor = &HFF0000
       lv_agentes.ListItems.item(i).ListSubItems(1).Bold = True
       lv_agentes.ListItems.item(i).ListSubItems(2).Bold = True
       lv_agentes.ListItems.item(i).ListSubItems(1).ForeColor = &HFF0000
       lv_agentes.ListItems.item(i).ListSubItems(2).ForeColor = &HFF0000
   Next
   lv_agentes.Refresh
End Sub

Private Sub Command6_Click()
   If var_todos_lineas = 1 Then
   Else
         var_todos_lineas = 0
   End If
   n = lv_agentes.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_agentes.ListItems.item(i).Selected = True
      If var_encontro = True And lv_agentes.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_agentes.selectedItem.SubItems(2) = "*"
         lv_agentes.ListItems.item(i).Bold = True
         lv_agentes.ListItems.item(i).ForeColor = &HFF0000
         lv_agentes.ListItems.item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_agentes.ListItems.item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_agentes.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_agentes.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub Command7_Click()
   var_todos_lineas = 0
   i = lv_agentes.selectedItem.Index
   If lv_agentes.selectedItem.SubItems(2) = "*" Then
      lv_agentes.selectedItem.SubItems(2) = ""
      lv_agentes.ListItems.item(i).Bold = False
      lv_agentes.ListItems.item(i).ForeColor = &H80000012
      lv_agentes.ListItems.item(i).ListSubItems(1).Bold = False
      lv_agentes.ListItems.item(i).ListSubItems(2).Bold = False
      lv_agentes.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
      lv_agentes.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
      lv_agentes.Refresh
   Else
      lv_agentes.selectedItem.SubItems(2) = "*"
      lv_agentes.ListItems.item(i).Bold = True
      lv_agentes.ListItems.item(i).ForeColor = &HFF0000
      lv_agentes.ListItems.item(i).ListSubItems(1).Bold = True
      lv_agentes.ListItems.item(i).ListSubItems(2).Bold = True
      lv_agentes.ListItems.item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_agentes.ListItems.item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_agentes.Refresh
   End If
End Sub

Private Sub Command8_Click()
   If var_todos_lineas = 1 Then
   Else
        var_todos_lineas = 0
   End If
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.item(i).Selected = True
      If lv_agentes.selectedItem.SubItems(2) = "*" Then
         lv_agentes.selectedItem.SubItems(2) = ""
         lv_agentes.ListItems.item(i).Bold = False
         lv_agentes.ListItems.item(i).ForeColor = &H80000012
         lv_agentes.ListItems.item(i).ListSubItems(1).Bold = False
         lv_agentes.ListItems.item(i).ListSubItems(2).Bold = False
         lv_agentes.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
         lv_agentes.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_agentes.selectedItem.SubItems(2) = "*"
         lv_agentes.ListItems.item(i).Bold = True
         lv_agentes.ListItems.item(i).ForeColor = &HFF0000
         lv_agentes.ListItems.item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_agentes.ListItems.item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub Command9_Click()
   var_todos_lineas = 0
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.item(i).Selected = True
      lv_agentes.selectedItem.SubItems(2) = ""
      lv_agentes.ListItems.item(i).Bold = False
      lv_agentes.ListItems.item(i).ForeColor = &H80000012
      lv_agentes.ListItems.item(i).ListSubItems(1).Bold = False
      lv_agentes.ListItems.item(i).ListSubItems(2).Bold = False
      lv_agentes.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
      lv_agentes.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_agentes.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 86 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
   If Shift = 4 And KeyCode = 65 Then
      cmd_aceptar_Click
   End If
   If Shift = 4 And KeyCode = 67 Then
      cmd_cancelar_Click
   End If
End Sub

Private Sub Form_Load()
   frm_agentes.Visible = False
   var_cadena_seguridad = ""
   Top = 0
   Left = 0
   If var_empresa = "03" Then
      rs.Open "select * from tb_almacenes where inte_alm_surtir = 1 and vcha_uor_unidad_id = '12'", cnn, adOpenDynamic, adLockOptimistic
   Else
      rs.Open "select * from tb_almacenes where inte_alm_surtir = 1 and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   If rs.EOF Then
      rs.Close
      'MsgBox ""
      Unload Me
   Else
      var_almacen = rs!VCHA_ALM_ALMACEN_ID
      rs.Close
   End If
   Label1.Caption = "Carga de Pedidos a la Fecha " + Format(Date, "long date")
   frm_carga_archivos.Visible = False
   rs.Open "select vcha_pri_ruta_pedidos from tb_principal", cnn, adOpenDynamic, adLockOptimistic
   var_ruta = rs(0).Value
   rs.Close
   Set var_tabla = CreateObject("ADODB.connection")
   var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
   var_tabla_correcta = False
   Call llenar_lista
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_cargapedidos)
End Sub

Private Sub lst_origen_DblClick()
   var_opcion_carga = 1
   n = lst_origen.ListIndex
   If n >= 0 Then
      Dim var_Archivo As String
      n = lst_origen.ListIndex
      var_Archivo = lst_origen.List(n)
      var_tabla_correcta = False
      Call valida(var_Archivo)
      If var_tabla_correcta = True Then
         lst_archivos.AddItem lst_origen.List(n)
         lst_origen.RemoveItem (lst_origen.ListIndex)
      End If
   Else
      MsgBox "No se a seleccionado ningun archivo", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub lst_origen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_opcion_carga = 1
      n = lst_origen.ListIndex
      If n >= 0 Then
         Dim var_Archivo As String
         n = lst_origen.ListIndex
         var_Archivo = lst_origen.List(n)
         var_tabla_correcta = False
         Call valida(var_Archivo)
         If var_tabla_correcta = True Then
            lst_archivos.AddItem lst_origen.List(n)
            lst_origen.RemoveItem (lst_origen.ListIndex)
         End If
      Else
         MsgBox "No se a seleccionado ningun archivo", vbOKOnly, "ATENCION"
      End If
   End If
End Sub



Sub valida(var_Archivo)
On Error GoTo mal:
   Cadena = "select titular,cvecliente,establecim,numero,fecha,codigo,cantidad,precio,sugerido from " + Trim(var_Archivo)
   var_bien = 0
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   var_tabla_correcta = True
   rs.Close
   If var_tabla_correcta = False And var_opcion_carga = 1 Then
mal:
      If var_opcion_carga = 1 Then
         MsgBox "El archivo " + Trim(var_Archivo) + " no es valido, precione enter para continuar", vbOKOnly, "atenciona"
      End If
   End If
End Sub

Private Sub lv_agentes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_agentes, ColumnHeader)
End Sub

Private Sub lv_agentes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim numero_lineas As Integer
      Dim numero_seleccionado1 As Integer
      Dim numero_seleccionado2 As Integer
      Dim primera_vez As Boolean
      Dim segunda_vez As Boolean
      Dim i As Integer
      Dim n As Integer
      Dim list_item As ListItem
      n = lv_agentes.ListItems.Count
      i = lv_agentes.selectedItem.Index
      If lv_agentes.ListItems.item(i).SubItems(2) = "*" Then
      lv_agentes.ListItems.item(i).SubItems(2) = " "
             lv_agentes.ListItems.item(i).Bold = False
             lv_agentes.ListItems.item(i).ForeColor = &H80000012
             lv_agentes.ListItems.item(i).ListSubItems(1).Bold = False
             lv_agentes.ListItems.item(i).ListSubItems(2).Bold = False
             lv_agentes.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
             lv_agentes.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
          Else
             lv_agentes.ListItems.item(i).SubItems(2) = "*"
             lv_agentes.ListItems.item(i).Bold = True
             lv_agentes.ListItems.item(i).ForeColor = &HFF0000
             lv_agentes.ListItems.item(i).ListSubItems(1).Bold = True
             lv_agentes.ListItems.item(i).ListSubItems(2).Bold = True
             lv_agentes.ListItems.item(i).ListSubItems(1).ForeColor = &HFF0000
             lv_agentes.ListItems.item(i).ListSubItems(2).ForeColor = &HFF0000
         End If
      lv_agentes.Refresh
   End If
   If KeyAscii = 27 Then
      frm_agentes.Visible = False
   End If
End Sub

Private Sub lv_pedidos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_pedidos, ColumnHeader)
End Sub

Private Sub lv_pedidos_GotFocus()
On Error GoTo salir:
   txt_numero = lv_pedidos.selectedItem
   txt_agente = lv_pedidos.selectedItem.SubItems(1)
   txt_titular = lv_pedidos.selectedItem.SubItems(2)
   txt_establecimiento = lv_pedidos.selectedItem.SubItems(3)
   txt_cliente = lv_pedidos.selectedItem.SubItems(4)
   txt_Cantidad = Format(lv_pedidos.selectedItem.SubItems(5), "###,###,##0.00")
   txt_importe_total = Format(lv_pedidos.selectedItem.SubItems(6), "###,###,##0.00")
   txt_descuento = Format((lv_pedidos.selectedItem.SubItems(8) * 1) + (lv_pedidos.selectedItem.SubItems(10) * 1) + (lv_pedidos.selectedItem.SubItems(12) * 1), "###,###,##0.00")
   txt_subimporte = Format((lv_pedidos.selectedItem.SubItems(6) * 1) - ((lv_pedidos.selectedItem.SubItems(8) * 1) + (lv_pedidos.selectedItem.SubItems(10) * 1) + (lv_pedidos.selectedItem.SubItems(12) * 1)), "###,###,##0.00")
   txt_iva = Format(lv_pedidos.selectedItem.SubItems(14), "###,###,##0.00")
   txt_importe_neto = Format(((lv_pedidos.selectedItem.SubItems(6) * 1) - ((lv_pedidos.selectedItem.SubItems(8) * 1) + (lv_pedidos.selectedItem.SubItems(10) * 1) + (lv_pedidos.selectedItem.SubItems(12) * 1)) + (lv_pedidos.selectedItem.SubItems(14) * 1)), "###,###,##0.00")
   txt_descuento1 = Format(lv_pedidos.selectedItem.SubItems(7), "###,###,##0.00")
   txt_descuento2 = Format(lv_pedidos.selectedItem.SubItems(9), "###,###,##0.00")
salir:
End Sub

Private Sub lv_pedidos_ItemClick(ByVal item As MSComctlLib.ListItem)
On Error GoTo salir:
   txt_numero = lv_pedidos.selectedItem
   txt_agente = lv_pedidos.selectedItem.SubItems(1)
   txt_titular = lv_pedidos.selectedItem.SubItems(2)
   txt_establecimiento = lv_pedidos.selectedItem.SubItems(3)
   txt_cliente = lv_pedidos.selectedItem.SubItems(4)
   txt_Cantidad = Format(lv_pedidos.selectedItem.SubItems(5), "###,###,##0.00")
   txt_importe_total = Format(lv_pedidos.selectedItem.SubItems(6), "###,###,##0.00")
   txt_descuento = Format((lv_pedidos.selectedItem.SubItems(8) * 1) + (lv_pedidos.selectedItem.SubItems(10) * 1) + (lv_pedidos.selectedItem.SubItems(12) * 1), "###,###,##0.00")
   txt_subimporte = Format((lv_pedidos.selectedItem.SubItems(6) * 1) - ((lv_pedidos.selectedItem.SubItems(8) * 1) + (lv_pedidos.selectedItem.SubItems(10) * 1) + (lv_pedidos.selectedItem.SubItems(12) * 1)), "###,###,##0.00")
   txt_iva = Format(lv_pedidos.selectedItem.SubItems(14), "###,###,##0.00")
   txt_importe_neto = Format(((lv_pedidos.selectedItem.SubItems(6) * 1) - ((lv_pedidos.selectedItem.SubItems(8) * 1) + (lv_pedidos.selectedItem.SubItems(10) * 1) + (lv_pedidos.selectedItem.SubItems(12) * 1)) + (lv_pedidos.selectedItem.SubItems(14) * 1)), "###,###,##0.00")
   txt_descuento1 = Format(lv_pedidos.selectedItem.SubItems(7), "###,###,##0.00")
   txt_descuento2 = Format(lv_pedidos.selectedItem.SubItems(9), "###,###,##0.00")
salir:
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub
