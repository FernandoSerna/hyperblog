VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_entradas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmoracle_entradas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Buscar Movimiento"
      Top             =   630
      Width           =   330
   End
   Begin VB.Frame frm_busqueda 
      Height          =   975
      Left            =   450
      TabIndex        =   42
      Top             =   915
      Width           =   2025
      Begin VB.TextBox txt_busqueda 
         Height          =   315
         Left            =   240
         TabIndex        =   43
         Top             =   495
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   " Busqueda de movimiento"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   6
         Left            =   30
         TabIndex        =   44
         Top             =   120
         Width           =   1950
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      Picture         =   "frmoracle_entradas.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir"
      Top             =   630
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmoracle_entradas.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   630
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11160
      Picture         =   "frmoracle_entradas.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   630
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   45
      TabIndex        =   18
      Top             =   510
      Width           =   11505
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   30
      TabIndex        =   17
      Top             =   870
      Width           =   11505
   End
   Begin VB.CommandButton cmd_mensaje_2 
      Caption         =   "mensaje 2"
      Height          =   195
      Left            =   1785
      TabIndex        =   2
      Top             =   675
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton cmd_mensaje_4 
      Caption         =   "mensaje 4"
      Height          =   195
      Left            =   1950
      TabIndex        =   1
      Top             =   675
      Visible         =   0   'False
      Width           =   75
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   60
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_entradas.frx":0940
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_entradas.frx":121A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_entradas.frx":1AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_entradas.frx":2090
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_entradas.frx":296C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_entradas.frx":3246
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_entradas.frx":3B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_entradas.frx":3C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_entradas.frx":3D44
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_entradas.frx":3E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_entradas.frx":3F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_entradas.frx":407A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_entradas.frx":418C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_entradas.frx":432E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_entradas.frx":5180
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_entradas.frx":5356
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_entradas.frx":5468
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Index           =   0
      Left            =   90
      TabIndex        =   19
      Top             =   915
      Width           =   6975
      Begin VB.TextBox txt_factura 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3450
         TabIndex        =   46
         Top             =   420
         Width           =   1395
      End
      Begin VB.TextBox txt_folio 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   420
         Width           =   1380
      End
      Begin VB.TextBox txt_nota 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   690
         TabIndex        =   20
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Factura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2565
         TabIndex        =   47
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4905
         TabIndex        =   41
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Nota:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   22
         Top             =   480
         Width           =   570
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   1
         Left            =   30
         TabIndex        =   21
         Top             =   120
         Width           =   6900
      End
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Index           =   3
      Left            =   7140
      TabIndex        =   23
      Top             =   915
      Width           =   2220
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   " Cantidad a Surtir"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   4
         Left            =   30
         TabIndex        =   25
         Top             =   120
         Width           =   2145
      End
      Begin VB.Label lbl_enviados 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   195
         TabIndex        =   24
         Top             =   420
         Width           =   1845
      End
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Index           =   4
      Left            =   9435
      TabIndex        =   26
      Top             =   915
      Width           =   2115
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   " Cantidad Surtida"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   5
         Left            =   30
         TabIndex        =   28
         Top             =   120
         Width           =   2040
      End
      Begin VB.Label lbl_recibidos 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   150
         TabIndex        =   27
         Top             =   420
         Width           =   1770
      End
   End
   Begin VB.Frame Frame3 
      Height          =   870
      Index           =   1
      Left            =   105
      TabIndex        =   29
      Top             =   1800
      Width           =   11460
      Begin VB.TextBox txt_proveedor 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6585
         TabIndex        =   31
         Top             =   420
         Width           =   4755
      End
      Begin VB.TextBox txt_destino 
         Enabled         =   0   'False
         Height          =   315
         Left            =   855
         TabIndex        =   30
         Top             =   435
         Width           =   4680
      End
      Begin VB.Label lbl_proveedor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Proveedor:"
         Height          =   195
         Left            =   5730
         TabIndex        =   34
         Top             =   480
         Width           =   780
      End
      Begin VB.Label label 
         BackColor       =   &H000000FF&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   30
         TabIndex        =   33
         Top             =   120
         Width           =   11385
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   32
         Top             =   495
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4665
      Left            =   90
      TabIndex        =   5
      Top             =   2610
      Width           =   11475
      Begin VB.CommandButton cmd_pasar_movimiento 
         Height          =   330
         Left            =   8880
         Picture         =   "frmoracle_entradas.frx":557A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   540
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txt_codigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1560
         TabIndex        =   11
         Top             =   465
         Width           =   3390
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   4440
         TabIndex        =   8
         Top             =   1575
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   75
            TabIndex        =   9
            Top             =   375
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H000000FF&
            Caption         =   "Cantidad a eliminar"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   10
            Top             =   15
            Width           =   2895
         End
      End
      Begin VB.TextBox txt_cantidad 
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
         Left            =   5865
         TabIndex        =   7
         Top             =   495
         Width           =   1890
      End
      Begin VB.TextBox txt_foco 
         Height          =   315
         Left            =   11655
         TabIndex        =   6
         Top             =   525
         Width           =   1650
      End
      Begin MSComctlLib.ListView lv_entradas 
         Height          =   3570
         Left            =   15
         TabIndex        =   13
         Top             =   1020
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   6297
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
         NumItems        =   29
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "   Código"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   8467
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Enviados"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Recibidos    "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Movimiento"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Faltan"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "shipment_header_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "shipment_line_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "item_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "to_organization_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "to_subinventory"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "receipt_source_code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "from_organization_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "ESTATUS"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "vendor_site_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "vendor_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Unit of measure"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "document_line_num"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "PROMISED_DATE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "deliver_to_person_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "LINE LOCATION_ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "ship_to_location_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "org_id "
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "country_of_origin_code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "UOM_CODE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "UNIT_PRICE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "customer_name"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "customer_ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "SITE_ID"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   5115
         TabIndex        =   16
         Top             =   615
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   615
         Width           =   1395
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   " Lectura de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   0
         Left            =   30
         TabIndex        =   14
         Top             =   120
         Width           =   11400
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp4 
      Height          =   75
      Left            =   10200
      TabIndex        =   39
      Top             =   480
      Visible         =   0   'False
      Width           =   30
      URL             =   "C:\sistemas\desarrollo\integral\type.wma"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   53
      _cy             =   132
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp3 
      Height          =   30
      Left            =   4725
      TabIndex        =   38
      Top             =   675
      Visible         =   0   'False
      Width           =   30
      URL             =   "C:\sistemas\desarrollo\integral\Articulo no en la OS.wav"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   53
      _cy             =   53
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   30
      Left            =   8505
      TabIndex        =   37
      Top             =   375
      Visible         =   0   'False
      Width           =   30
      URL             =   "C:\sistemas\desarrollo\integral\sound2.wav"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   53
      _cy             =   53
   End
   Begin VB.Label lblnombremovimiento 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   90
      TabIndex        =   36
      Top             =   30
      Width           =   11445
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp2 
      Height          =   135
      Left            =   1500
      TabIndex        =   35
      Top             =   735
      Visible         =   0   'False
      Width           =   30
      URL             =   "C:\sistemas\desarrollo\integral\CFFOUND.wav"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   53
      _cy             =   238
   End
End
Attribute VB_Name = "frmoracle_entradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_folio As Double
Dim var_primera_vez As Integer
Dim var_cantidad_leida As Double
Dim var_fecha_inicio As Date
Dim var_fecha_fin As Date
Dim var_renglon As Integer

Sub ilumina_grid()
   var_n = lv_entradas.ListItems.Count
   For var_i = 1 To var_n
       If var_i = var_renglon Then
          lv_entradas.ListItems.Item(var_i).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(3).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(4).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(5).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(6).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(7).Bold = True
          lv_entradas.ListItems.Item(var_i).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H8000&
       Else
          If (lv_entradas.ListItems.Item(var_i).ListSubItems(5) * 1) = 0 Then
             lv_entradas.ListItems.Item(var_i).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(1).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(2).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(3).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(4).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(5).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(6).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(7).Bold = False
             lv_entradas.ListItems.Item(var_i).ForeColor = &HFF&
             lv_entradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF&
             lv_entradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF&
             lv_entradas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF&
             lv_entradas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF&
             lv_entradas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF&
             lv_entradas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF&
             lv_entradas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HFF&
          Else
             lv_entradas.ListItems.Item(var_i).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(1).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(2).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(3).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(4).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(5).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(6).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(7).Bold = False
             lv_entradas.ListItems.Item(var_i).ForeColor = &H80000012
             lv_entradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000012
             lv_entradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000012
             lv_entradas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000012
             lv_entradas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000012
             lv_entradas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H80000012
             lv_entradas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H80000012
             lv_entradas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H80000012
          End If
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_entradas.ListItems.Item(var_renglon).Selected = True
      lv_entradas.selectedItem.EnsureVisible
   End If
   lv_entradas.Refresh
End Sub



Private Sub cmd_buscar_Click()
   Me.frm_busqueda.Visible = True
   Me.txt_busqueda = ""
   Me.txt_busqueda.SetFocus
End Sub

Private Sub cmd_imprimir_Click()
   Dim objConn As New ADODB.Connection
   Dim objCmd As New ADODB.Command
   Dim objParm As ADODB.Parameter
   If Me.lv_entradas.ListItems.Count > 0 Then
      If Me.txt_folio <> "" Then
         Dim clnt As New SoapClient30
         Dim var_arreglo() As String
         Dim var_s As String
         Dim var_paso As Boolean
         If Me.lv_entradas.ListItems.Count > 0 Then
            Me.lv_entradas.ListItems.Item(1).Selected = True
            VAR_ESTATUS = Me.lv_entradas.selectedItem.SubItems(13)
         Else
            VAR_ESTATUS = "I"
         End If
         If VAR_ESTATUS <> "I" Then
            var_si = MsgBox("Se va a imprimir la entrada y cerrar el movimiento", vbYesNo, "ATENCION")
            If var_si = 6 Then
               If var_clave_movimiento = "VENDOR" Then
                  If Me.lv_entradas.ListItems.Count > 0 Then
                     Me.lv_entradas.ListItems(1).Selected = True
                     var_fecha_compromiso = CDate(Me.lv_entradas.selectedItem.SubItems(18))
                     rs.Open "select * from xxvia_tb_recepciones WHERE shipment_num = '" + Me.txt_nota + "' and shipment_header_id = " + CStr(lv_entradas.selectedItem.SubItems(6)) + " and to_organization_id = " + CStr(lv_entradas.selectedItem.SubItems(9)) + " and receipt_source_code = '" + Me.lv_entradas.selectedItem.SubItems(11) + "' AND FOLIO = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        vendor_site_id = Me.lv_entradas.selectedItem.SubItems(14)
                        vendor_id = Me.lv_entradas.selectedItem.SubItems(15)
                        'clnt.MSSoapInit "http://intranet/wsebs12test/wsInterfacePO.asmx?wsdl"
                        'var_arreglo = clnt.crear_encabezado_recepcion(var_fecha_compromiso, var_unidad_organizacional, vendor_id, vendor_site_id, Me.txt_folio)
                        'Set clnt = Nothing
                        
                        rsaux1.Open "SELECT (next_receipt_num+1) idRec From rcv_parameters WHERE organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux1.EOF Then
                           var_next_receipt_num = rsaux1(0).Value
                        End If
                        rsaux1.Close
                        

                        
                        var_fecha_fin = Now
                        var_segundo_s = CStr(Second(var_fecha_fin))
                        var_minuto_s = CStr(Minute(var_fecha_fin))
                        var_hora_s = CStr(Hour(var_fecha_fin))
                        var_año_s = CStr(Year(var_fecha_fin))
                        var_mes_s = CStr(Month(var_fecha_fin))
                        var_dia_s = CStr(Day(var_fecha_fin))
                        If Len(var_segundo_s) = 1 Then
                           var_segundo_s = "0" + var_segundo_s
                        End If
                        If Len(var_minuto_s) = 1 Then
                           var_minuto_s = "0" + var_minuto_s
                        End If
                        If Len(var_hora_s) = 1 Then
                           var_hora_s = "0" + var_hora_s
                        End If
                        If Len(var_año_s) = 2 Then
                           var_año_s = "20" + var_año_s
                        End If
                        If Len(var_mes_s) = 1 Then
                           var_mes_s = "0" + var_mes_s
                        End If
                        If Len(var_dia_s) = 1 Then
                           var_dia_s = "0" + var_dia_s
                        End If
                        var_fecha_str_1 = var_dia_s + "/" + var_mes_s + "/" + var_año_s
                        var_fecha_str = var_año_s + "/" + var_mes_s + "/" + var_dia_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
                        
                        
                        var_concurrente = 0
                        objConn.Open var_conexion_oracle
                        With objCmd
                             objConn.BeginTrans
                             .ActiveConnection = objConn
                             .CommandText = "XXVIA_PK_INTERFACES_PO.crear_encabezado_recepcion2"
                             .CommandType = adCmdStoredProc
                                
                             Set objParm = .CreateParameter("p_expected_receipt_date", adVarChar, adParamInput, 50, var_fecha_str_1)
                             .Parameters.Append objParm
                          
                             Set objParm = .CreateParameter("p_ship_to_organization_id", adNumeric, adParamInput, 200, var_unidad_organizacional)
                             .Parameters.Append objParm
                                      
                             Set objParm = .CreateParameter("p_vendor_id", adNumeric, adParamInput, 200, vendor_id)
                             .Parameters.Append objParm
                                      
                             Set objParm = .CreateParameter("p_vendor_site_id", adNumeric, adParamInput, 200, vendor_site_id)
                             .Parameters.Append objParm
                                      
                             Set objParm = .CreateParameter("p_recepcion_sid", adNumeric, adParamInput, 200, CDbl(Me.txt_folio))
                             .Parameters.Append objParm
                                      
                             Set objParm = .CreateParameter("p_attrib12", adVarChar, adParamInput, 200, "SIDEC_" + Me.txt_folio)
                             .Parameters.Append objParm
                                      
                             Set objParm = .CreateParameter("p_attrib13", adVarChar, adParamInput, 200, "FACTURA: " + Me.txt_factura)
                             .Parameters.Append objParm
                                      
                             Set objParm = .CreateParameter("p_attrib14", adVarChar, adParamInput, 200, "")
                             .Parameters.Append objParm
                                      
                             Set objParm = .CreateParameter("x_header_interface_id", adNumeric, adParamOutput, 200, 0)
                             .Parameters.Append objParm
                                    
                             Set objParm = .CreateParameter("x_group_id", adNumeric, adParamOutput, 200, 0)
                             .Parameters.Append objParm
                                      
                             On Error GoTo SALIR
                             .execute
                         
                             var_header_interface_id = .Parameters("x_header_interface_id").Value
                             objConn.CommitTrans
                             
                             var_group_id = .Parameters("x_group_id").Value
                             'objConn.CommitTrans
                             
                        End With
                        Set objConn = Nothing
                        Set objCmd = Nothing
                        
                        var_cadena = "select xc.po_header_id oc_identificador, xc.num_oc oc_numero, xc.po_line_id oc_linea_identificador, xc.line_num oc_linea_numero, xc.item_number articulo_identificador, xc.item_description articulo_descripcion, xc.quantity cantidad_pendiente, xc.unit_meas_lookup_code oc_unidad_medida, xc.uom_code unidad_medida_primaria, xc.unit_price precio_unitario, xc.currency_code moneda, xc.vendor_name proveedor, xc.quantity+tolerance CANTIDAD_MAXIMA, xc.item_id, xc.closed_code, xc.vendor_id, xc.deliver_to_person_id, xc.line_num, xc.line_location_id, xc.ship_to_location_id, xc.country_of_origin_code, xc.vendor_site_id, xc.RELEASE_NUM,  xc.PO_RELEASE_ID, xc.TYPE_LOOKUP_CODE, cantidad, to_subinventory, factura  FROM xxvia_vw_recepcion_compra xc, xxvia_tb_recepciones xr"
                        var_cadena = var_cadena + " where xc.CLOSED_CODE = 'OPEN' AND xc.num_oc = '" + Me.txt_nota + "' AND xc.ship_to_organization_id =  " + var_unidad_organizacional + " AND xc.org_id  =  " + var_empresa + " and xc.org_id = xr.from_organization_id and xc.ship_to_organization_id = xr.to_organization_id and xc.num_oc = xr.shipment_num and xr.folio = " + Me.txt_folio + " and xc.po_line_id = xr.shipment_line_id  AND XR.LINE_LOCATION_ID = XC.LINE_LOCATION_ID"
                                                
                        rsaux1.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        While Not rsaux1.EOF
                              
                               rsaux2.Open "SELECT rcv_transactions_interface_s.NEXTVAL FROM DUAL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                               If Not rsaux2.EOF Then
                                  VAR_INTERFACE_TRANSACTION_ID = rsaux2(0).Value
                               End If
                              rsaux2.Close
                              
                              var_release_num = CStr(IIf(IsNull(rsaux1!RELEASE_NUM), 0, rsaux1!RELEASE_NUM))
                              If var_release_num = "0" Then
                                 var_release_num = "NULL"
                              End If
                              var_cadena = "Insert Into rcv_transactions_interface (interface_transaction_id, GROUP_ID,last_update_date, last_updated_by, creation_date, created_by, last_update_login, transaction_type,transaction_date, processing_status_code, processing_mode_code, transaction_status_code, quantity, unit_of_measure, item_id, employee_id, auto_transact_code, po_header_id, po_line_id, po_line_location_id, receipt_source_code, to_organization_code, source_document_code, document_num, destination_type_code, deliver_to_person_id, deliver_to_location_id, subinventory, header_interface_id, validation_flag, release_num)"
                              var_cadena = var_cadena + " VALUES (" + CStr(VAR_INTERFACE_TRANSACTION_ID) + "," + CStr(var_group_id) + ",SYSDATE, 1170, SYSDATE, 1170, 0, 'RECEIVE', SYSDATE, 'PENDING', 'BATCH', 'PENDING', " + CStr(IIf(IsNull(rsaux1!Cantidad), 0, rsaux1!Cantidad)) + ",'" + CStr(rsaux1!oc_unidad_medida) + "'," + CStr(rsaux1!ITEM_ID) + ",NULL,'DELIVER'," + CStr(rsaux1!oc_identificador) + "," + CStr(rsaux1!oc_linea_identificador) + "," + CStr(rsaux1!line_location_id) + ",'VENDOR','CDI','PO'," + Me.txt_nota + ",'INVENTORY'," + CStr(rsaux1!deliver_to_person_id) + "," + CStr(rsaux1!ship_to_location_id) + ",'" + rsaux1!TO_SUBINVENTORY + "'," + CStr(var_header_interface_id) + ",'Y'," + var_release_num + ")"
                              rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              rsaux1.MoveNext
                        Wend
                        rsaux1.Close
                        
                        
                        
                        
                        var_concurrente = 0
                        objConn.Open var_conexion_oracle
                        With objCmd
                             objConn.BeginTrans
                             .ActiveConnection = objConn
                             .CommandText = "XXVIA_PK_INVENTARIOS.XXVIA_SP_CONCURRENTE_MAT"
                             .CommandType = adCmdStoredProc
                                
                             Set objParm = .CreateParameter("x_concurrente", adNumeric, adParamOutput, 50, 0)
                             .Parameters.Append objParm
                          
                             Set objParm = .CreateParameter("p_tipo_movimiento", adVarChar, adParamInput, 200, "Traspasos")
                             .Parameters.Append objParm
                                      
                             Set objParm = .CreateParameter("p_organization_id", adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
                             .Parameters.Append objParm
                                      
                             Set objParm = .CreateParameter("p_group_id", adNumeric, adParamInput, 200, var_group_id)
                             .Parameters.Append objParm
                                      
                             On Error GoTo SALIR
                             .execute
                         
                             var_concurrente = .Parameters("x_concurrente").Value
                             objConn.CommitTrans
                             
                        End With
                        Set objConn = Nothing
                        Set objCmd = Nothing
                        
                        
                        
                        var_mensaje = ""
                        While var_mensaje <> "EXITO"
                        var_concurrente = 0
                        objConn.Open var_conexion_oracle
                        With objCmd
                             objConn.BeginTrans
                             .ActiveConnection = objConn
                             .CommandText = "XXVIA_PK_RECEPCIONES_MP.xxvia_sp_eje_concurr_0"
                             .CommandType = adCmdStoredProc
                                
                             Set objParm = .CreateParameter("p_application", adVarChar, adParamInput, 50, "PO")
                             .Parameters.Append objParm
                          
                             Set objParm = .CreateParameter("p_program", adVarChar, adParamInput, 200, "RCVLCMWS")
                             .Parameters.Append objParm
                                      
                             Set objParm = .CreateParameter("p_description", adVarChar, adParamInput, 200, "Integracion de costo extendido SID")
                             .Parameters.Append objParm
                                      
                             Set objParm = .CreateParameter("p_usuario", adNumeric, adParamInput, 200, 1170)
                             .Parameters.Append objParm
                                      
                             Set objParm = .CreateParameter("p_resp", adNumeric, adParamInput, 200, 20560)
                             .Parameters.Append objParm
                                      
                             Set objParm = .CreateParameter("p_app", adNumeric, adParamInput, 200, 706)
                             .Parameters.Append objParm
                                      
                             Set objParm = .CreateParameter("p_mensaje", adVarChar, adParamOutput, 200, "")
                             .Parameters.Append objParm
                             
                             Set objParm = .CreateParameter("p_concurrente", adNumeric, adParamOutput, 200, 0)
                             .Parameters.Append objParm
                                      
                                      
                             On Error GoTo SALIR
                             .execute
                         
                             var_concurrente = .Parameters("p_concurrente").Value
                             
                             var_mensaje = IIf(IsNull(.Parameters("p_mensaje").Value), "", .Parameters("p_mensaje").Value)
                             objConn.CommitTrans
                             
                        End With
                        Set objConn = Nothing
                        Set objCmd = Nothing
                        Wend
                        
                        
                        var_concurrente = 0
                        objConn.Open var_conexion_oracle
                        With objCmd
                             objConn.BeginTrans
                             .ActiveConnection = objConn
                             .CommandText = "XXVIA_PK_INVENTARIOS.XXVIA_SP_CONCURRENTE_MAT"
                             .CommandType = adCmdStoredProc
                                
                             Set objParm = .CreateParameter("x_concurrente", adNumeric, adParamOutput, 50, 0)
                             .Parameters.Append objParm
                          
                             Set objParm = .CreateParameter("p_tipo_movimiento", adVarChar, adParamInput, 200, "ImpotarIterface")
                             .Parameters.Append objParm
                                      
                             Set objParm = .CreateParameter("p_organization_id", adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
                             .Parameters.Append objParm
                                      
                             Set objParm = .CreateParameter("p_group_id", adNumeric, adParamInput, 200, var_group_id)
                             .Parameters.Append objParm
                                      
                             On Error GoTo SALIR
                             .execute
                         
                             var_concurrente = .Parameters("x_concurrente").Value
                             objConn.CommitTrans
                             
                        End With
                        Set objConn = Nothing
                        Set objCmd = Nothing
                        
                        
                        var_concurrente = 0
                        objConn.Open var_conexion_oracle
                        With objCmd
                             objConn.BeginTrans
                             .ActiveConnection = objConn
                             .CommandText = "XXVIA_PK_RECEPCIONES_MP.XXVIA_SP_ESPERA_DELIVER"
                             .CommandType = adCmdStoredProc
                                
                             Set objParm = .CreateParameter("p_header_interface_id", adNumeric, adParamInput, 200, var_header_interface_id)
                             .Parameters.Append objParm
                             
                             
                             On Error GoTo SALIR
                             .execute
                         
                        End With
                        Set objConn = Nothing
                        Set objCmd = Nothing
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        rsaux.Open "UPDATE xxvia_tb_recepciones SET ESTATUS = 'I' WHERE shipment_num = '" + Me.txt_nota + "' and shipment_header_id = " + CStr(lv_entradas.selectedItem.SubItems(6)) + " and to_organization_id = " + CStr(lv_entradas.selectedItem.SubItems(9)) + " and receipt_source_code = '" + Me.lv_entradas.selectedItem.SubItems(11) + "' AND FOLIO = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        Me.txt_codigo.Enabled = False
                        Me.txt_foco.Enabled = False
                     

                        rsaux.Open "UPDATE XXVIA_TB_RECEPCIONES SET ESTATUS = 'I', fecha_fin = to_date('" + var_fecha_str + "','yyyy/mm/dd hh24:mi:ss') WHERE shipment_num = '" + Me.txt_nota + "' and shipment_header_id = " + CStr(lv_entradas.selectedItem.SubItems(6)) + " and to_organization_id = " + CStr(lv_entradas.selectedItem.SubItems(9)) + " and receipt_source_code = '" + Me.lv_entradas.selectedItem.SubItems(11) + "' AND FOLIO = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        For var_j = 1 To lv_entradas.ListItems.Count
                            lv_entradas.ListItems.Item(var_j).Selected = True
                            Me.lv_entradas.selectedItem.SubItems(13) = "I"
                        Next var_j
                        Me.txt_codigo.Enabled = False
                  
                        rsaux10.Open "select * from rcv_shipment_headers where attribute12 =  'SIDEC_" + Me.txt_folio + "'", cnnoracle_4
                        If Not rsaux10.EOF Then
                           cnn.BeginTrans
                           rsaux11.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_RECEPCIONES", cnn, adOpenDynamic, adLockOptimistic
                           var_consecutivo = IIf(IsNull(rsaux11(0).Value), 0, rsaux11(0).Value) + 1
                           rsaux11.Close
                           rsaux11.Open "INSERT INTO TB_TEMP_ORACLE_RECEPCIONES (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                           cnn.CommitTrans
                           VAR_SHIPMENT_HEADER = rsaux10!shipment_header_id
                           var_cadena = "SELECT SEGMENT1 AS CODIGO, DESCRIPTION AS NOMBRE_aRTICULO, po_unit_price as precio, currency_conversion_rate as tipo_cambio,A.SHIPMENT_HEADER_ID, a.SHIPMENT_LINE_ID, a.ITEM_ID, SHIPMENT_NUM, A.quantity_shipped AS CANTIDAD_ENVIADA, A.quantity_received AS CANTIDAD_RECIBIDA, B.ship_to_org_id AS TO_ORGANIZATION_ID, B.organization_id AS FROM_ORGANIZATION_ID, '' AS to_subinventory, d.vendor_id, b.attribute13 FROM rcv_shipment_lines A, RCV_SHIPMENT_HEADERS B, xxvia_system_items_b C, RCV_transactions D Where a.SHIPMENT_HEADER_ID = " + CStr(VAR_SHIPMENT_HEADER) + " AND A.shipment_header_id =  B.shipment_header_id AND A.ITEM_ID = C.INVENTORY_ITEM_ID AND A.to_organization_id = C.organization_id and a.SHIPMENT_HEADER_ID = D.SHIPMENT_HEADER_ID and a.shipment_line_id = d.shipment_line_id and d.destination_type_code = 'INVENTORY'"
                           rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           While Not rsaux.EOF
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_RECEPCIONES (INTE_TEM_CONSECUTIVO, ORGANIZACION_DESTINO, ORGANIZACION_ORIGEN, SHIPMENT_NUM, SHIPMENT_HEADER_ID, SHIPMENT_LINE_ID,                                                SUBINVENTARIO, SEGMENT1, CANTIDAD_ENVIADA, CANTIDAD_RECIBIDA, DESCRIPCION, VENDOR_ID, PRECIO, TIPO_CAMBIO, referencia)"
                                 var_cadena = var_cadena + "VALUES (" + CStr(var_consecutivo) + "," + CStr(IIf(IsNull(rsaux!TO_organizaTion_ID), 0, rsaux!TO_organizaTion_ID)) + "," + CStr(IIf(IsNull(rsaux!FROM_ORGANIZATION_ID), 0, rsaux!FROM_ORGANIZATION_ID)) + ",'" + Me.txt_nota + "'," + CStr(rsaux!shipment_header_id) + "," + CStr(rsaux!SHIPMENT_LINE_ID) + ",'" + IIf(IsNull(rsaux!TO_SUBINVENTORY), "", rsaux!TO_SUBINVENTORY) + "','" + rsaux!CODIGO + "'," + CStr(rsaux!CANTIDAD_ENVIADA) + "," + CStr(rsaux!CANTIDAD_RECIBIDA) + ",'" + IIf(IsNull(rsaux!NOMBRE_ARTICULO), "", rsaux!NOMBRE_ARTICULO) + "'," + CStr(rsaux!vendor_id) + "," + CStr(IIf(IsNull(rsaux!Precio), 0, rsaux!Precio)) + "," + CStr(IIf(IsNull(rsaux!TIPO_CAMBIO), 1, rsaux!TIPO_CAMBIO)) + ",'" + IIf(IsNull(rsaux!attribute13), "", rsaux!attribute13) + "')"
                                 'MsgBox var_cadena
                                 rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                 rsaux.MoveNext
                           Wend
                           rsaux.Close
                           
                           
                           rsaux1.Open "select organizacion_Destino, organizacion_origen, subinventario from TB_TEMP_ORACLE_RECEPCIONES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and organizacion_Destino is not null", cnn, adOpenDynamic, adLockOptimistic
                   
                           var_nombre_unidad_destino = ""
                           var_nombre_unidad_origen = ""
                           var_nombre_almacen_subinventario = Me.txt_destino
                           If Not rsaux1.EOF Then
                              rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + CStr(rsaux1!organizacion_origen), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux.EOF Then
                                 var_nombre_unidad_origen = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                              End If
                              rsaux.Close
                              rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + CStr(rsaux1!organizacion_destino), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux.EOF Then
                                 var_nombre_unidad_destino = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                              End If
                              rsaux.Close
                              rsaux.Open "update TB_TEMP_ORACLE_RECEPCIONES set nombre_ORGANIZACION_destino = '" + var_nombre_unidad_destino + "', nombre_ORGANIZACION_origen = '" + var_nombre_unidad_origen + "', nombre_subinventario = '" + var_nombre_almacen_subinventario + "', NOMBRE_PROVEEDOR = '" + Me.txt_proveedor + "' where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                              rsaux1.Close
                   
                              rsaux1.Open "SELECT SUM(CANTIDAD) FROM XXVIA_TB_RECEPCIONES WHERE FOLIO = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux1.EOF Then
                                 var_cantidad_total = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
                              End If
                              rsaux1.Close
                              var_si = 0
                      
                              rsaux1.Open "SELECT * FROM XXVIA_TB_RECEPCIONES WHERE FOLIO = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux1.EOF Then
                                 rsaux.Open "UPDATE TB_TEMP_ORACLE_RECEPCIONES SET fecha_inicio = '" + CStr(rsaux1!FECHA_INiCIO) + "', USUARIO = '" + rsaux1!USUARIO + "',MAQUINA = '" + rsaux1!maquina + "', MOVIMIENTO = '" + Me.lblnombremovimiento + "', FOLIO = " + Me.txt_folio + " where inte_Tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux1.Close
                      
                              rsaux.Open "delete from TB_TEMP_ORACLE_RECEPCIONES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ORGANIZACION_DESTINO is null", cnn, adOpenDynamic, adLockOptimistic
                              Set reporte = appl.OpenReport(App.Path + "\REP_ORACLE_RECEPCIONES_ENTRADAS_POR_COMPRA.rpt")
                       
                              reporte.RecordSelectionFormula = "{VW_ORACLE_RECEPCIONES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                              frmvistasprevias.cr.ReportSource = reporte
                              For ntablas = 1 To reporte.Database.Tables.Count
                                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                              Next ntablas
                              frmvistasprevias.cr.ViewReport
                              frmvistasprevias.Caption = Me.lblnombremovimiento
                              frmvistasprevias.Show 1
                              Set reporte = Nothing
                           Else
                              MsgBox "El movimiento no se a terminado de generar en ORACLE, espere un momento por favor y vuelvalo a intentar", vbOKOnly, "ATENCION"
                              rsaux1.Close
                           End If
                        Else
                           MsgBox "El movimiento no se a terminado de generar en ORACLE, espere un momento por favor y vuelvalo a intentar", vbOKOnly, "ATENCION"
                        End If
                        rsaux10.Close
                     Else
                        MsgBox "El movimiento esta vacio", vbOKOnly, "ATENCION"
                     End If
                   rs.Close
                  End If
               Else
                  If var_clave_movimiento = "DC" Then
                     rsaux10.Open "select * from xxvia_tb_recepciones where folio = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     clnt.MSSoapInit var_webservice
                     var_arreglo() = clnt.generar_encabezado_devolucion(rsaux10!TO_organizaTion_ID, rsaux10!CUSTOMER_ID, Me.txt_folio)
                     Set clnt = Nothing
                     While Not rsaux10.EOF
                           clnt.MSSoapInit var_webservice
                           var_b = clnt.generar_linea_devolucion(var_arreglo(0), var_arreglo(1), rsaux10!Cantidad, CStr(Date), rsaux10!MEDIDA, rsaux10!UOM_CODE, CStr(rsaux10!FECHA_PROMESA), 0, rsaux10!ITEM_ID, rsaux10!ship_to_location_id, rsaux10!TO_organizaTion_ID, rsaux10!shipment_header_id, rsaux10!SHIPMENT_LINE_ID, rsaux10!CUSTOMER_ID, rsaux10!site_id, rsaux10!ORG_ID, rsaux10!TO_SUBINVENTORY)
                           Set clnt = Nothing
                           rsaux10.MoveNext
                     Wend
                     rsaux10.Close
                     
                     var_fecha_fin = Now
                     var_segundo_s = CStr(Second(var_fecha_fin))
                     var_minuto_s = CStr(Minute(var_fecha_fin))
                     var_hora_s = CStr(Hour(var_fecha_fin))
                     var_año_s = CStr(Year(var_fecha_fin))
                     var_mes_s = CStr(Month(var_fecha_fin))
                     var_dia_s = CStr(Day(var_fecha_fin))
                     If Len(var_segundo_s) = 1 Then
                        var_segundo_s = "0" + var_segundo_s
                     End If
                     If Len(var_minuto_s) = 1 Then
                        var_minuto_s = "0" + var_minuto_s
                     End If
                     If Len(var_hora_s) = 1 Then
                        var_hora_s = "0" + var_hora_s
                     End If
                     If Len(var_año_s) = 2 Then
                        var_año_s = "20" + var_año_s
                     End If
                     If Len(var_mes_s) = 1 Then
                        var_mes_s = "0" + var_mes_s
                     End If
                     If Len(var_dia_s) = 1 Then
                        var_dia_s = "0" + var_dia_s
                     End If
                     var_fecha_str = var_año_s + "/" + var_mes_s + "/" + var_dia_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
                     rs.Open "UPDATE XXVIA_TB_RECEPCIONES SET ESTATUS = 'I', fecha_fin = to_date('" + var_fecha_str + "','yyyy/mm/dd hh24:mi:ss') WHERE shipment_num = '" + Me.txt_nota + "' and shipment_header_id = " + CStr(lv_entradas.selectedItem.SubItems(6)) + " and to_organization_id = " + CStr(lv_entradas.selectedItem.SubItems(9)) + " and receipt_source_code = '" + Me.lv_entradas.selectedItem.SubItems(11) + "' AND FOLIO = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     For var_j = 1 To lv_entradas.ListItems.Count
                         lv_entradas.ListItems.Item(var_j).Selected = True
                         Me.lv_entradas.selectedItem.SubItems(13) = "I"
                     Next var_j
                     Me.txt_codigo.Enabled = False
                  Else
                     clnt.MSSoapInit "http://intranet/wsoracle/wsInterfaceINV.asmx?wsdl"
                     If var_clave_movimiento = "INVENTORY" Then
                        VAR_TIPO_M = "21"
                     End If
                     var_s = clnt.registraDatosRecordSet(CDbl(Me.txt_folio), VAR_TIPO_M)
                     Set clnt = Nothing
                     If var_s = "0" Then
                        var_fecha_fin = Now
                        var_segundo_s = CStr(Second(var_fecha_fin))
                        var_minuto_s = CStr(Minute(var_fecha_fin))
                        var_hora_s = CStr(Hour(var_fecha_fin))
                        var_año_s = CStr(Year(var_fecha_fin))
                        var_mes_s = CStr(Month(var_fecha_fin))
                        var_dia_s = CStr(Day(var_fecha_fin))
                        If Len(var_segundo_s) = 1 Then
                           var_segundo_s = "0" + var_segundo_s
                        End If
                        If Len(var_minuto_s) = 1 Then
                           var_minuto_s = "0" + var_minuto_s
                        End If
                        If Len(var_hora_s) = 1 Then
                           var_hora_s = "0" + var_hora_s
                        End If
                        If Len(var_año_s) = 2 Then
                           var_año_s = "20" + var_año_s
                        End If
                        If Len(var_mes_s) = 1 Then
                           var_mes_s = "0" + var_mes_s
                        End If
                        If Len(var_dia_s) = 1 Then
                           var_dia_s = "0" + var_dia_s
                        End If
                        var_fecha_str = var_año_s + "/" + var_mes_s + "/" + var_dia_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
                        rs.Open "UPDATE XXVIA_TB_RECEPCIONES SET ESTATUS = 'I', fecha_fin = to_date('" + var_fecha_str + "','yyyy/mm/dd hh24:mi:ss') WHERE shipment_num = '" + Me.txt_nota + "' and shipment_header_id = " + CStr(lv_entradas.selectedItem.SubItems(6)) + " and to_organization_id = " + CStr(lv_entradas.selectedItem.SubItems(9)) + " and receipt_source_code = '" + Me.lv_entradas.selectedItem.SubItems(11) + "' AND FOLIO = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        For var_j = 1 To lv_entradas.ListItems.Count
                            lv_entradas.ListItems.Item(var_j).Selected = True
                            Me.lv_entradas.selectedItem.SubItems(13) = "I"
                        Next var_j
                        Me.txt_codigo.Enabled = False
                  
                        If var_clave_movimiento = "VENDOR" Then
                           rsaux10.Open "select * from rcv_shipment_headers where attribute15 =  '" + Me.txt_folio + "'", cnnoracle_4
                           If Not rsaux10.EOF Then
                              VAR_SHIPMENT_HEADER = rsaux10!shipment_header_id
                              var_cadena = "SELECT SEGMENT1 AS CODIGO, DESCRIPTION AS NOMBRE_aRTICULO, po_unit_price as precio, currency_conversion_rate as tipo_cambio,A.SHIPMENT_HEADER_ID, a.SHIPMENT_LINE_ID, a.ITEM_ID, SHIPMENT_NUM, A.quantity_shipped AS CANTIDAD_ENVIADA, A.quantity_received AS CANTIDAD_RECIBIDA, B.ship_to_org_id AS TO_ORGANIZATION_ID, B.organization_id AS FROM_ORGANIZATION_ID, '' AS to_subinventory, d.vendor_id  FROM rcv_shipment_lines A, RCV_SHIPMENT_HEADERS B, xxvia_system_items_b C, RCV_transactions D Where a.SHIPMENT_HEADER_ID = " + CStr(VAR_SHIPMENT_HEADER) + " AND A.shipment_header_id =  B.shipment_header_id AND A.ITEM_ID = C.INVENTORY_ITEM_ID AND A.to_organization_id = C.organization_id and a.SHIPMENT_HEADER_ID = D.SHIPMENT_HEADER_ID and a.shipment_line_id = d.shipment_line_id and d.destination_type_code = 'INVENTORY'"
                              rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              While Not rsaux.EOF
                                    var_cadena = "INSERT INTO TB_TEMP_ORACLE_RECEPCIONES (INTE_TEM_CONSECUTIVO, ORGANIZACION_DESTINO, ORGANIZACION_ORIGEN, SHIPMENT_NUM, SHIPMENT_HEADER_ID, SHIPMENT_LINE_ID,                                                SUBINVENTARIO, SEGMENT1, CANTIDAD_ENVIADA, CANTIDAD_RECIBIDA, DESCRIPCION, VENDOR_ID, PRECIO, TIPO_CAMBIO)"
                                    var_cadena = var_cadena + "VALUES (" + CStr(var_consecutivo) + "," + CStr(IIf(IsNull(rsaux!TO_organizaTion_ID), 0, rsaux!TO_organizaTion_ID)) + "," + CStr(IIf(IsNull(rsaux!FROM_ORGANIZATION_ID), 0, rsaux!FROM_ORGANIZATION_ID)) + ",'" + Me.txt_nota + "'," + CStr(rsaux!shipment_header_id) + "," + CStr(rsaux!SHIPMENT_LINE_ID) + ",'" + IIf(IsNull(rsaux!TO_SUBINVENTORY), "", rsaux!TO_SUBINVENTORY) + "','" + rsaux!CODIGO + "'," + CStr(rsaux!CANTIDAD_ENVIADA) + "," + CStr(rsaux!CANTIDAD_RECIBIDA) + ",'" + IIf(IsNull(rsaux!NOMBRE_ARTICULO), "", rsaux!NOMBRE_ARTICULO) + "'," + CStr(rsaux!vendor_id) + "," + CStr(IIf(IsNull(rsaux!Precio), 0, rsaux!Precio)) + "," + CStr(IIf(IsNull(rsaux!TIPO_CAMBIO), 1, rsaux!TIPO_CAMBIO)) + ")"
                                    rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    rsaux.MoveNext
                              Wend
                              rsaux.Close
                              rsaux1.Open "select organizacion_Destino, organizacion_origen, subinventario from TB_TEMP_ORACLE_RECEPCIONES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and organizacion_Destino is not null", cnn, adOpenDynamic, adLockOptimistic
                   
                              var_nombre_unidad_destino = ""
                              var_nombre_unidad_origen = ""
                              var_nombre_almacen_subinventario = Me.txt_destino
                     
                              rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + CStr(rsaux1!organizacion_origen), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux.EOF Then
                                 var_nombre_unidad_origen = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                              End If
                              rsaux.Close
                              rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + CStr(rsaux1!organizacion_destino), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux.EOF Then
                                 var_nombre_unidad_destino = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                              End If
                              rsaux.Close
                              rsaux.Open "update TB_TEMP_ORACLE_RECEPCIONES set nombre_ORGANIZACION_destino = '" + var_nombre_unidad_destino + "', nombre_ORGANIZACION_origen = '" + var_nombre_unidad_origen + "', nombre_subinventario = '" + var_nombre_almacen_subinventario + "', NOMBRE_PROVEEDOR = '" + Me.txt_proveedor + "' where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                              rsaux1.Close
                
                              rsaux1.Open "SELECT SUM(CANTIDAD) FROM XXVIA_TB_RECEPCIONES WHERE FOLIO = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux1.EOF Then
                                 var_cantidad_total = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
                              End If
                              rsaux1.Close
                              var_si = 0
                      
                              rsaux1.Open "SELECT * FROM XXVIA_TB_RECEPCIONES WHERE FOLIO = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux1.EOF Then
                                 rsaux.Open "UPDATE TB_TEMP_ORACLE_RECEPCIONES SET fecha_inicio = '" + CStr(rsaux1!FECHA_INiCIO) + "', USUARIO = '" + rsaux1!USUARIO + "',MAQUINA = '" + rsaux1!maquina + "', MOVIMIENTO = '" + Me.lblnombremovimiento + "', FOLIO = " + Me.txt_folio + " where inte_Tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux1.Close
                         
                              rsaux.Open "delete from TB_TEMP_ORACLE_RECEPCIONES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ORGANIZACION_DESTINO is null", cnn, adOpenDynamic, adLockOptimistic
                              Set reporte = appl.OpenReport(App.Path + "\REP_ORACLE_RECEPCIONES_ENTRADAS_POR_COMPRA.rpt")
                       
                              reporte.RecordSelectionFormula = "{VW_ORACLE_RECEPCIONES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                              frmvistasprevias.cr.ReportSource = reporte
                              For ntablas = 1 To reporte.Database.Tables.Count
                                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                              Next ntablas
                              frmvistasprevias.cr.ViewReport
                              frmvistasprevias.Caption = Me.lblnombremovimiento
                              frmvistasprevias.Show 1
                              Set reporte = Nothing
                           Else
                              MsgBox "El movimiento no se a terminado de generar en ORACLE, espere un momento por favor y vuelvalo a intentar"
                           End If
                           rsaux10.Close
                        Else
                           cnn.BeginTrans
                           rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_RECEPCIONES", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux.EOF Then
                              var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
                           Else
                              var_consecutivo = 1
                           End If
                           rsaux.Close
                           rsaux.Open "INSERT INTO TB_TEMP_ORACLE_RECEPCIONES (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                           cnn.CommitTrans
                           Me.lv_entradas.ListItems.Item(1).Selected = True
                           VAR_SHIPMENT_HEADER = Me.lv_entradas.selectedItem.SubItems(6)
                           var_cadena = "SELECT SEGMENT1 AS CODIGO, DESCRIPTION AS NOMBRE_aRTICULO, A.SHIPMENT_HEADER_ID, SHIPMENT_LINE_ID, ITEM_ID, SHIPMENT_NUM, A.quantity_shipped AS CANTIDAD_ENVIADA, A.quantity_received AS CANTIDAD_RECIBIDA, B.ship_to_org_id AS TO_ORGANIZATION_ID, B.organization_id AS FROM_ORGANIZATION_ID, A.to_subinventory  FROM rcv_shipment_lines A, RCV_SHIPMENT_HEADERS B, xxvia_system_items_b C Where a.SHIPMENT_HEADER_ID = " + CStr(VAR_SHIPMENT_HEADER) + " AND B.SHIPMENT_NUM = '" + Me.txt_nota + "' AND A.shipment_header_id =  B.shipment_header_id AND A.ITEM_ID = C.INVENTORY_ITEM_ID AND A.to_organization_id = C.organization_id "
                           rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           While Not rsaux.EOF
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_RECEPCIONES (INTE_TEM_CONSECUTIVO, ORGANIZACION_DESTINO, ORGANIZACION_ORIGEN, SHIPMENT_NUM, SHIPMENT_HEADER_ID, SHIPMENT_LINE_ID,                                                SUBINVENTARIO, SEGMENT1, CANTIDAD_ENVIADA, CANTIDAD_RECIBIDA, DESCRIPCION)"
                                 var_cadena = var_cadena + "VALUES (" + CStr(var_consecutivo) + "," + CStr(IIf(IsNull(rsaux!TO_organizaTion_ID), 0, rsaux!TO_organizaTion_ID)) + "," + CStr(IIf(IsNull(rsaux!FROM_ORGANIZATION_ID), 0, rsaux!FROM_ORGANIZATION_ID)) + ",'" + Me.txt_nota + "'," + CStr(rsaux!shipment_header_id) + "," + CStr(rsaux!SHIPMENT_LINE_ID) + ",'" + IIf(IsNull(rsaux!TO_SUBINVENTORY), "", rsaux!TO_SUBINVENTORY) + "','" + rsaux!CODIGO + "'," + CStr(rsaux!CANTIDAD_ENVIADA) + "," + CStr(rsaux!CANTIDAD_RECIBIDA) + ",'" + IIf(IsNull(rsaux!NOMBRE_ARTICULO), "", rsaux!NOMBRE_ARTICULO) + "')"
                                 rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                 rsaux.MoveNext
                           Wend
                           rsaux.Close
                     
                           rsaux1.Open "select organizacion_Destino, organizacion_origen, subinventario from TB_TEMP_ORACLE_RECEPCIONES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and organizacion_Destino is not null", cnn, adOpenDynamic, adLockOptimistic
                           var_nombre_unidad_destino = ""
                           var_nombre_unidad_origen = ""
                           var_nombre_almacen_subinventario = ""
                        
                           rsaux.Open "SELECT * FROM PO_VENDORS WHERE VENDOR_ID = " + var_unidad_organizacional + " AND SECONDARY_INVENTORY_NAME = '" + IIf(IsNull(rsaux1!SUBINVENTARIO), "", rsaux1!SUBINVENTARIO) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux.EOF Then
                              var_nombre_almacen_subinventario = IIf(IsNull(rsaux!Description), "", rsaux!Description)
                           End If
                           rsaux.Close
                           rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + CStr(rsaux1!organizacion_origen), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux.EOF Then
                              var_nombre_unidad_origen = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                           End If
                           rsaux.Close
                           rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + CStr(rsaux1!organizacion_destino), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux.EOF Then
                              var_nombre_unidad_destino = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                           End If
                           rsaux.Close
                           rsaux.Open "update TB_TEMP_ORACLE_RECEPCIONES set nombre_ORGANIZACION_destino = '" + var_nombre_unidad_destino + "', nombre_ORGANIZACION_origen = '" + var_nombre_unidad_origen + "', nombre_subinventario = '" + var_nombre_almacen_subinventario + "' where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                           rsaux1.Close
                        
                           rsaux1.Open "SELECT SUM(CANTIDAD) FROM XXVIA_TB_RECEPCIONES WHERE FOLIO = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux1.EOF Then
                              var_cantidad_total = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
                           End If
                           rsaux1.Close
                           var_si = 0
                           While var_si = 0
                                 rsaux1.Open "SELECT sum(QUANTITY) AS CANTIDAD FROM RCV_transactions WHERE SHIPMENT_HEADER_ID  =" + CStr(VAR_SHIPMENT_HEADER) + " AND ATTRIBUTE1 = " + Me.txt_folio + " AND destination_type_code = 'RECEIVING'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux1.EOF Then
                                    var_cantidad_oracle = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
                                 Else
                                    var_cantidad_oracle = 0
                                 End If
                                 rsaux1.Close
                                 If var_cantidad_oracle = var_cantidad_total Then
                                    var_si = 1
                                 End If
                           Wend
                           rsaux1.Open "SELECT * FROM XXVIA_TB_RECEPCIONES WHERE FOLIO = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux1.EOF Then
                              rsaux.Open "UPDATE TB_TEMP_ORACLE_RECEPCIONES SET fecha_inicio = '" + CStr(rsaux1!FECHA_INiCIO) + "', FECHA_FIN = '" + CStr(rsaux1!FECHA_FIN) + "',USUARIO = '" + rsaux1!USUARIO + "',MAQUINA = '" + rsaux1!maquina + "', MOVIMIENTO = '" + Me.lblnombremovimiento + "', FOLIO = " + Me.txt_folio + " where inte_Tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux1.Close
                        
                           rsaux1.Open "SELECT shipment_header_id, shipment_line_id, QUANTITY AS CANTIDAD FROM RCV_transactions WHERE SHIPMENT_HEADER_ID  =" + CStr(VAR_SHIPMENT_HEADER) + " AND ATTRIBUTE1 = " + Me.txt_folio + " AND destination_type_code = 'RECEIVING'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           While Not rsaux1.EOF
                                 rsaux2.Open "UPDATE TB_TEMP_ORACLE_RECEPCIONES SET CANTIDAD_MOVIMIENTO = " + CStr(IIf(IsNull(rsaux1!Cantidad), 0, rsaux1!Cantidad)) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND SHIPMENT_HEADER_ID = " + CStr(rsaux1!shipment_header_id) + " AND SHIPMENT_LINE_ID = " + CStr(rsaux1!SHIPMENT_LINE_ID) + " AND SHIPMENT_NUM = '" + Me.txt_nota + "'", cnn, adOpenDynamic, adLockOptimistic
                                 rsaux1.MoveNext
                           Wend
                           rsaux1.Close
                           
                          rsaux.Open "delete from TB_TEMP_ORACLE_RECEPCIONES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ORGANIZACION_DESTINO is null", cnn, adOpenDynamic, adLockOptimistic
                           Set reporte = appl.OpenReport(App.Path + "\REP_ORACLE_RECEPCIONES.rpt")
                           
                           reporte.RecordSelectionFormula = "{VW_ORACLE_RECEPCIONES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                           frmvistasprevias.cr.ReportSource = reporte
                           For ntablas = 1 To reporte.Database.Tables.Count
                               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                           Next ntablas
                           frmvistasprevias.cr.ViewReport
                           frmvistasprevias.Caption = Me.lblnombremovimiento
                           frmvistasprevias.Show 1
                           Set reporte = Nothing
                        End If
                     Else
                        If var_s = "False" Then
                           MsgBox "No se afecto el movimiento, intentelo nuevamente", vbOKOnly, "ATENCION"
                        Else
                           MsgBox var_s, vbOKOnly, "ATENCION"
                        End If
                     End If
                  End If
               End If
            End If
         Else
            cnn.BeginTrans
            rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_RECEPCIONES", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rsaux.Close
            rsaux.Open "INSERT INTO TB_TEMP_ORACLE_RECEPCIONES (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            If var_clave_movimiento = "VENDOR" Then
               If rsaux10.State = 1 Then
                  rsaux10.Close
               End If
               rsaux10.Open "select * from rcv_shipment_headers where attribute12 =  'SIDEC_" + Me.txt_folio + "'", cnnoracle_4
               If Not rsaux10.EOF Then
                  VAR_SHIPMENT_HEADER = rsaux10!shipment_header_id
                  var_cadena = "SELECT SEGMENT1 AS CODIGO, DESCRIPTION AS NOMBRE_aRTICULO, po_unit_price as precio, currency_conversion_rate as tipo_cambio,A.SHIPMENT_HEADER_ID, a.SHIPMENT_LINE_ID, a.ITEM_ID, SHIPMENT_NUM, A.quantity_shipped AS CANTIDAD_ENVIADA, A.quantity_received AS CANTIDAD_RECIBIDA, B.ship_to_org_id AS TO_ORGANIZATION_ID, B.organization_id AS FROM_ORGANIZATION_ID, '' AS to_subinventory, d.vendor_id, b.attribute13  FROM rcv_shipment_lines A, RCV_SHIPMENT_HEADERS B, xxvia_system_items_b C, RCV_transactions D Where a.SHIPMENT_HEADER_ID = " + CStr(VAR_SHIPMENT_HEADER) + " AND A.shipment_header_id =  B.shipment_header_id AND A.ITEM_ID = C.INVENTORY_ITEM_ID AND A.to_organization_id = C.organization_id and a.SHIPMENT_HEADER_ID = D.SHIPMENT_HEADER_ID and a.shipment_line_id = d.shipment_line_id and d.destination_type_code = 'INVENTORY'"
                  rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rsaux.EOF
                        var_cadena = "INSERT INTO TB_TEMP_ORACLE_RECEPCIONES (INTE_TEM_CONSECUTIVO, ORGANIZACION_DESTINO, ORGANIZACION_ORIGEN, SHIPMENT_NUM, SHIPMENT_HEADER_ID, SHIPMENT_LINE_ID,                                                SUBINVENTARIO, SEGMENT1, CANTIDAD_ENVIADA, CANTIDAD_RECIBIDA, DESCRIPCION, VENDOR_ID, PRECIO, TIPO_CAMBIO, referencia)"
                        var_cadena = var_cadena + "VALUES (" + CStr(var_consecutivo) + "," + CStr(IIf(IsNull(rsaux!TO_organizaTion_ID), 0, rsaux!TO_organizaTion_ID)) + "," + CStr(IIf(IsNull(rsaux!FROM_ORGANIZATION_ID), 0, rsaux!FROM_ORGANIZATION_ID)) + ",'" + Me.txt_nota + "'," + CStr(rsaux!shipment_header_id) + "," + CStr(rsaux!SHIPMENT_LINE_ID) + ",'" + IIf(IsNull(rsaux!TO_SUBINVENTORY), "", rsaux!TO_SUBINVENTORY) + "','" + rsaux!CODIGO + "'," + CStr(rsaux!CANTIDAD_ENVIADA) + "," + CStr(rsaux!CANTIDAD_RECIBIDA) + ",'" + IIf(IsNull(rsaux!NOMBRE_ARTICULO), "", rsaux!NOMBRE_ARTICULO) + "'," + CStr(rsaux!vendor_id) + "," + CStr(IIf(IsNull(rsaux!Precio), 0, rsaux!Precio)) + "," + CStr(IIf(IsNull(rsaux!TIPO_CAMBIO), 1, rsaux!TIPO_CAMBIO)) + ",'" + IIf(IsNull(rsaux!attribute13), "", rsaux!attribute13) + "')"
                        rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  If rsaux1.State = 1 Then
                     rsaux1.Close
                  End If
                  rsaux1.Open "select organizacion_Destino, organizacion_origen, subinventario from TB_TEMP_ORACLE_RECEPCIONES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and organizacion_Destino is not null", cnn, adOpenDynamic, adLockOptimistic
                   
                  var_nombre_unidad_destino = ""
                  var_nombre_unidad_origen = ""
                  var_nombre_almacen_subinventario = Me.txt_destino
                  If Not rsaux1.EOF Then
                     rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + CStr(rsaux1!organizacion_origen), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        var_nombre_unidad_origen = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                     End If
                     rsaux.Close
                     rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + CStr(rsaux1!organizacion_destino), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        var_nombre_unidad_destino = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                     End If
                     rsaux.Close
                     rsaux.Open "update TB_TEMP_ORACLE_RECEPCIONES set nombre_ORGANIZACION_destino = '" + var_nombre_unidad_destino + "', nombre_ORGANIZACION_origen = '" + var_nombre_unidad_origen + "', nombre_subinventario = '" + var_nombre_almacen_subinventario + "', NOMBRE_PROVEEDOR = '" + Me.txt_proveedor + "' where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                     rsaux1.Close
                   
                     rsaux1.Open "SELECT SUM(CANTIDAD) FROM XXVIA_TB_RECEPCIONES WHERE FOLIO = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux1.EOF Then
                        var_cantidad_total = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
                     End If
                     rsaux1.Close
                     var_si = 0
                    
                     rsaux1.Open "SELECT * FROM XXVIA_TB_RECEPCIONES WHERE FOLIO = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux1.EOF Then
                        rsaux.Open "UPDATE TB_TEMP_ORACLE_RECEPCIONES SET fecha_inicio = '" + CStr(rsaux1!FECHA_INiCIO) + "', USUARIO = '" + rsaux1!USUARIO + "',MAQUINA = '" + rsaux1!maquina + "', MOVIMIENTO = '" + Me.lblnombremovimiento + "', FOLIO = " + Me.txt_folio + " where inte_Tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux1.Close
                         
                     rsaux.Open "delete from TB_TEMP_ORACLE_RECEPCIONES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ORGANIZACION_DESTINO is null", cnn, adOpenDynamic, adLockOptimistic
                     Set reporte = appl.OpenReport(App.Path + "\REP_ORACLE_RECEPCIONES_ENTRADAS_POR_COMPRA.rpt")
                       
                     reporte.RecordSelectionFormula = "{VW_ORACLE_RECEPCIONES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                     frmvistasprevias.cr.ReportSource = reporte
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     frmvistasprevias.cr.ViewReport
                     frmvistasprevias.Caption = Me.lblnombremovimiento
                     frmvistasprevias.Show 1
                     Set reporte = Nothing
                  Else
                     MsgBox "El movimiento no se a terminado de generar en ORACLE, espere un momento por favor y vuelvalo a intentar"
                     rsaux1.Close
                  End If
               Else
                  MsgBox "El movimiento no se a terminado de generar en ORACLE, espere un momento por favor y vuelvalo a intentar"
               End If
               rsaux10.Close
            Else
               If var_clave_movimiento = "DC" Then
                  rsaux8.Open "select * from rcv_shipment_headers where attribute15 = '" + Me.txt_folio + "'", cnnoracle_4
                  If Not rsaux8.EOF Then
                     rsaux9.Open "select * from rcv_shipment_lines where shipment_header_id = " + CStr(rsaux8!shipment_header_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux9.EOF Then
                        rsaux7.Open "SELECT * FROM XXVIA_TB_rECEPCIONES WHERE FOLIO = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux7.EOF Then
                           var_nombre_cliente = rsaux7!customer_name
                           VAR_ESTABLECIMIENTO = rsaux7!site_id
                           VAR_CLIENTE = rsaux7!CUSTOMER_ID
                           VAR_AGENTE = rsaux7!vendor_id
                           VAR_USUARIO_MOV = rsaux7!USUARIO
                           MAQUINA_MOV = rsaux7!maquina
                           FECHA_INiCIO = rsaux7!FECHA_INiCIO
                           FECHA_FIN = rsaux7!FECHA_FIN
                        End If
                        rsaux7.Close
                        rsaux7.Open "SELECT address1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= " + VAR_ESTABLECIMIENTO + " AND HCSU.SITE_USE_ID = D.site_use_id AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                        If Not rsaux7.EOF Then
                           VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux7!address1), "", rsaux7!address1)
                        End If
                        rsaux7.Close
                        rsaux.Open "SELECT * FROM AR_COLLECTORS WHERE COLLECTOR_ID = '" + VAR_AGENTE + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           var_nombre_agente = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                        End If
                        rsaux.Close
                       
                        var_almacen = rsaux9!TO_SUBINVENTORY
                        rsaux.Open "SELECT * FROM mtl_secondary_inventories WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND SECONDARY_INVENTORY_NAME = '" + IIf(IsNull(rsaux9!TO_SUBINVENTORY), "", rsaux9!TO_SUBINVENTORY) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           var_nombre_almacen_subinventario = IIf(IsNull(rsaux!Description), "", rsaux!Description)
                        End If
                        rsaux.Close
                        rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + CStr(rsaux9!TO_organizaTion_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           var_nombre_unidad_origen = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                        End If
                        rsaux.Close
                        
                        While Not rsaux9.EOF
                              rsaux.Open "SELECT * FROM xxvia_system_items_b WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND INVENTORY_ITEM_ID = " + CStr(rsaux9!ITEM_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              VAR_SEGMENT2 = ""
                              If Not rsaux.EOF Then
                                 VAR_SEGMENT2 = IIf(IsNull(rsaux!SEGMENT1), "", rsaux!SEGMENT1)
                              End If
                              rsaux.Close
                              var_cadena = "INSERT INTO TB_TEMP_ORACLE_RECEPCIONES (INTE_TEM_CONSECUTIVO, ORGANIZACION_DESTINO, ORGANIZACION_ORIGEN, SHIPMENT_NUM, SHIPMENT_HEADER_ID, SHIPMENT_LINE_ID,                                                SUBINVENTARIO, SEGMENT1, CANTIDAD_ENVIADA, CANTIDAD_RECIBIDA, DESCRIPCION, NOMBRE_ORGANIZACION_DESTINO, NOMBRE_ORGANIZACION_ORIGEN,NOMBRE_SUBINVENTARIO,USUARIO, MAQUINA, FECHA_INICIO, FECHA_FIN, CLIENTE_ID, ESTABLECIMIENTO_ID,VENDOR_ID, NOMBRE_CLIENTE,NOMBRE_ESTABLECIMIENTO,NOMBRE_PROVEEDOR,MOVIMIENTO, FOLIO)"
                              var_cadena = var_cadena + "VALUES (" + CStr(var_consecutivo) + "," + CStr(IIf(IsNull(rsaux9!TO_organizaTion_ID), 0, rsaux9!TO_organizaTion_ID)) + "," + CStr(IIf(IsNull(rsaux9!FROM_ORGANIZATION_ID), 0, rsaux9!FROM_ORGANIZATION_ID)) + ",'" + Me.txt_nota + "',0,0,'" + IIf(IsNull(rsaux9!TO_SUBINVENTORY), "", rsaux9!TO_SUBINVENTORY) + "','" + VAR_SEGMENT2 + "'," + CStr(rsaux9!QUANTITY_RECEIVED) + "," + CStr(rsaux9!QUANTITY_RECEIVED) + ",'" + IIf(IsNull(rsaux9!item_description), "", rsaux9!item_description) + "','" + var_nombre_unidad_origen + "','" + var_nombre_unidad_origen + "','" + var_nombre_almacen_subinventario + "','" + var_clave_usuario_global + "','" + MAQUINA_MOV + "','" + CStr(FECHA_INiCIO) + "','" + CStr(FECHA_FIN) + "','" + VAR_CLIENTE + "','" + VAR_ESTABLECIMIENTO + "','" + VAR_AGENTE + "','" + var_nombre_cliente + "','" + VAR_NOMBRE_ESTABLECIMIENTO + "','" + var_nombre_agente + "','DEVOLUCION DE CLIENTES','" + Me.txt_folio + "')"
                              rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              rsaux9.MoveNext
                        Wend
                        rsaux.Open "delete from TB_TEMP_ORACLE_RECEPCIONES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and maquina is null", cnn, adOpenDynamic, adLockOptimistic
                        Set reporte = appl.OpenReport(App.Path + "\rep_oracle_recepciones_devoluciones_clientes.rpt")
                        reporte.RecordSelectionFormula = "{VW_ORACLE_RECEPCIONES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                        frmvistasprevias.cr.ReportSource = reporte
                        For ntablas = 1 To reporte.Database.Tables.Count
                            reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                        Next ntablas
                        frmvistasprevias.cr.ViewReport
                        frmvistasprevias.Caption = Me.lblnombremovimiento
                        frmvistasprevias.Show 1
                        Set reporte = Nothing
                     Else
                        MsgBox "El movimiento no se a terminado de generar en ORACLE, espere un momento por favor y vuelvalo a intentar"
                     End If
                     rsaux9.Close
                  Else
                     MsgBox "El movimiento no se a terminado de generar en ORACLE, espere un momento por favor y vuelvalo a intentar"
                  End If
                  rsaux8.Close
               Else
                  Me.lv_entradas.ListItems.Item(1).Selected = True
                  VAR_SHIPMENT_HEADER = Me.lv_entradas.selectedItem.SubItems(6)
                  var_cadena = "SELECT SEGMENT1 AS CODIGO, DESCRIPTION AS NOMBRE_aRTICULO, A.SHIPMENT_HEADER_ID, SHIPMENT_LINE_ID, ITEM_ID, SHIPMENT_NUM, A.quantity_shipped AS CANTIDAD_ENVIADA, A.quantity_received AS CANTIDAD_RECIBIDA, B.ship_to_org_id AS TO_ORGANIZATION_ID, B.organization_id AS FROM_ORGANIZATION_ID, A.to_subinventory  FROM rcv_shipment_lines A, RCV_SHIPMENT_HEADERS B, xxvia_system_items_b C Where a.SHIPMENT_HEADER_ID = " + CStr(VAR_SHIPMENT_HEADER) + " AND B.SHIPMENT_NUM = '" + Me.txt_nota + "' AND A.shipment_header_id =  B.shipment_header_id AND A.ITEM_ID = C.INVENTORY_ITEM_ID AND A.to_organization_id = C.organization_id "
                  rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rsaux.EOF
                        var_cadena = "INSERT INTO TB_TEMP_ORACLE_RECEPCIONES (INTE_TEM_CONSECUTIVO, ORGANIZACION_DESTINO, ORGANIZACION_ORIGEN, SHIPMENT_NUM, SHIPMENT_HEADER_ID, SHIPMENT_LINE_ID,                                                SUBINVENTARIO, SEGMENT1, CANTIDAD_ENVIADA, CANTIDAD_RECIBIDA, DESCRIPCION)"
                        var_cadena = var_cadena + "VALUES (" + CStr(var_consecutivo) + "," + CStr(IIf(IsNull(rsaux!TO_organizaTion_ID), 0, rsaux!TO_organizaTion_ID)) + "," + CStr(IIf(IsNull(rsaux!FROM_ORGANIZATION_ID), 0, rsaux!FROM_ORGANIZATION_ID)) + ",'" + Me.txt_nota + "'," + CStr(rsaux!shipment_header_id) + "," + CStr(rsaux!SHIPMENT_LINE_ID) + ",'" + IIf(IsNull(rsaux!TO_SUBINVENTORY), "", rsaux!TO_SUBINVENTORY) + "','" + rsaux!CODIGO + "'," + CStr(rsaux!CANTIDAD_ENVIADA) + "," + CStr(rsaux!CANTIDAD_RECIBIDA) + ",'" + IIf(IsNull(rsaux!NOMBRE_ARTICULO), "", rsaux!NOMBRE_ARTICULO) + "')"
                        rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  rsaux1.Open "select organizacion_Destino, organizacion_origen, subinventario from TB_TEMP_ORACLE_RECEPCIONES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and organizacion_Destino is not null", cnn, adOpenDynamic, adLockOptimistic
                   
                  var_nombre_unidad_destino = ""
                  var_nombre_unidad_origen = ""
                  var_nombre_almacen_subinventario = ""
                  
                  rsaux.Open "SELECT * FROM mtl_secondary_inventories WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND SECONDARY_INVENTORY_NAME = '" + IIf(IsNull(rsaux1!SUBINVENTARIO), "", rsaux1!SUBINVENTARIO) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_nombre_almacen_subinventario = IIf(IsNull(rsaux!Description), "", rsaux!Description)
                  End If
                  rsaux.Close
                  rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + CStr(rsaux1!organizacion_origen), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_nombre_unidad_origen = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                  End If
                  rsaux.Close
                  rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + CStr(rsaux1!organizacion_destino), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_nombre_unidad_destino = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                  End If
                  rsaux.Close
                  rsaux.Open "update TB_TEMP_ORACLE_RECEPCIONES set nombre_ORGANIZACION_destino = '" + var_nombre_unidad_destino + "', nombre_ORGANIZACION_origen = '" + var_nombre_unidad_origen + "', nombre_subinventario = '" + var_nombre_almacen_subinventario + "' where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Close
                   
                  rsaux1.Open "SELECT SUM(CANTIDAD) FROM XXVIA_TB_RECEPCIONES WHERE FOLIO = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                     var_cantidad_total = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
                  End If
                  rsaux1.Close
                  var_si = 0
                  While var_si = 0
                        rsaux1.Open "SELECT sum(QUANTITY) AS CANTIDAD FROM RCV_transactions WHERE SHIPMENT_HEADER_ID  =" + CStr(VAR_SHIPMENT_HEADER) + " AND ATTRIBUTE1 = " + Me.txt_folio + " AND destination_type_code = 'RECEIVING'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux1.EOF Then
                           var_cantidad_oracle = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
                        Else
                           var_cantidad_oracle = 0
                        End If
                        rsaux1.Close
                        If var_cantidad_oracle = var_cantidad_total Then
                          var_si = 1
                        End If
                  Wend
                  rsaux1.Open "SELECT * FROM XXVIA_TB_RECEPCIONES WHERE FOLIO = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                     rsaux.Open "UPDATE TB_TEMP_ORACLE_RECEPCIONES SET fecha_inicio = '" + CStr(rsaux1!FECHA_INiCIO) + "', FECHA_FIN = '" + CStr(rsaux1!FECHA_FIN) + "',USUARIO = '" + rsaux1!USUARIO + "',MAQUINA = '" + rsaux1!maquina + "', MOVIMIENTO = '" + Me.lblnombremovimiento + "', FOLIO = " + Me.txt_folio + " where inte_Tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux1.Close
                      
                  rsaux1.Open "SELECT shipment_header_id, shipment_line_id, QUANTITY AS CANTIDAD FROM RCV_transactions WHERE SHIPMENT_HEADER_ID  =" + CStr(VAR_SHIPMENT_HEADER) + " AND ATTRIBUTE1 = " + Me.txt_folio + " AND destination_type_code = 'RECEIVING'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "UPDATE TB_TEMP_ORACLE_RECEPCIONES SET CANTIDAD_MOVIMIENTO = " + CStr(IIf(IsNull(rsaux1!Cantidad), 0, rsaux1!Cantidad)) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND SHIPMENT_HEADER_ID = " + CStr(rsaux1!shipment_header_id) + " AND SHIPMENT_LINE_ID = " + CStr(rsaux1!SHIPMENT_LINE_ID) + " AND SHIPMENT_NUM = '" + Me.txt_nota + "'", cnn, adOpenDynamic, adLockOptimistic
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                      
                  rsaux.Open "delete from TB_TEMP_ORACLE_RECEPCIONES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ORGANIZACION_DESTINO is null", cnn, adOpenDynamic, adLockOptimistic
                  Set reporte = appl.OpenReport(App.Path + "\REP_ORACLE_RECEPCIONES.rpt")
                    
                  reporte.RecordSelectionFormula = "{VW_ORACLE_RECEPCIONES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = Me.lblnombremovimiento
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
               End If
            End If
         End If
      Else
         MsgBox "No se a seleccionado un movimiento", vbYesNo, "ATENCIO"
      End If
   Else
      MsgBox "No existe el movimiento", vbOKOnly, "ATENCION"
   End If
   Exit Sub
SALIR:
   If Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
       rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      Resume
    Else
      If rs.State = 1 Then
         rs.Close
      End If
      If rsaux.State = 1 Then
         rsaux.Close
      End If
      If rsaux1.State = 1 Then
         rsaux1.Close
      End If
      If rsaux2.State = 1 Then
         rsaux2.Close
      End If
      If rsaux3.State = 1 Then
         rsaux3.Close
      End If
      If rsaux4.State = 1 Then
         rsaux4.Close
      End If
      If rsaux5.State = 1 Then
         rsaux5.Close
      End If
      If rsaux6.State = 1 Then
         rsaux6.Close
      End If
      If rsaux7.State = 1 Then
         rsaux7.Close
      End If
      If rsaux8.State = 1 Then
         rsaux8.Close
      End If
      If rsaux9.State = 1 Then
         rsaux9.Close
      End If
      If rsaux10.State = 1 Then
         rsaux10.Close
      End If
    
    End If
End Sub

Private Sub cmd_mensaje_2_Click()
   Me.wmp2.Controls.play
End Sub

Private Sub cmd_mensaje_4_Click()
   Me.wmp4.Controls.play
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_factura = ""
   Me.txt_nota.Enabled = True
   Me.txt_nota = ""
   Me.txt_proveedor = ""
   Me.txt_destino = ""
   Me.txt_codigo = ""
   Me.txt_cantidad = ""
   Me.lbl_enviados = ""
   Me.lbl_recibidos = ""
   Me.lv_entradas.ListItems.Clear
   Me.txt_nota.SetFocus
   var_primera_vez = 1
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
   Me.Caption = var_descripcion_recepcion
   Me.lblnombremovimiento = var_descripcion_recepcion
   Me.frm_eliminar.Visible = False
   Me.txt_cantidad.Visible = False
   Me.lbl_cantidad.Visible = False
   Me.frm_busqueda.Visible = False
   If var_clave_movimiento = "DC" Then
      Me.lbl_proveedor = "Cliente:"
   End If
End Sub


Private Sub lv_entradas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_entradas, ColumnHeader)
End Sub

Private Sub lv_entradas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If Trim(Me.lv_entradas.selectedItem.SubItems(13)) = "" Then
         Me.frm_eliminar.Visible = True
         Me.txt_cantidad_eliminar = ""
         Me.txt_cantidad_eliminar.SetFocus
      Else
         MsgBox "El movimiento ya no puede ser modificado", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_busqueda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_busqueda) Then
         If var_clave_movimiento = "VENDOR" Then
            rs.Open "SELECT * FROM XXVIA_TB_rECEPCIONES WHERE FOLIO = " + Me.txt_busqueda + " AND RECEIPT_SOURCE_CODE = '" + var_clave_movimiento + "' AND TO_ORGANIZATION_ID = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               Me.txt_factura = IIf(IsNull(rs!FACTURA), "", rs!FACTURA)
               Me.txt_destino = ""
               var_almacen_global = rs!TO_SUBINVENTORY
               rsaux.Open "SELECT * FROM mtl_secondary_inventories WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND SECONDARY_INVENTORY_NAME = '" + IIf(IsNull(rs!TO_SUBINVENTORY), "", rs!TO_SUBINVENTORY) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  Me.txt_destino = IIf(IsNull(rsaux!Description), "", rsaux!Description)
               End If
               rsaux.Close
               Me.lv_entradas.ListItems.Clear
               Me.txt_folio = Me.txt_busqueda
               Me.txt_nota = rs!shipment_num
               
               Me.txt_proveedor = ""
               var_cantidad_enviada = 0
               var_cantidad_recibida = 0
               var_fecha_inicio = CDate(rs!FECHA_INiCIO)
               rsaux.Open "select deliver_to_person_id, PROMISED_DATE, unit_meas_lookup_code, line_num, vendor_id, vendor_site_id, num_oc SHIPMENT_NUM, PO_HEADER_ID SHIPMENT_HEADER_ID, PO_LINE_ID SHIPMENT_LINE_ID,  ITEM_ID, quantity AS QUANTITY_SHIPPED, 0 QUANTITY_RECEIVED, VENDOR_ID, VENDOR_NAME, item_number, ITEM_DESCRIPTION, ORG_ID , SHIP_TO_ORGANIZATION_ID, line_location_id, ship_to_location_id, country_of_origin_code, UOM_CODE, UNIT_PRICE  from xxvia_vw_recepcion_compra where num_oc = " + Me.txt_nota + " AND SHIP_TO_ORGANIZATION_ID = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
               While Not rsaux.EOF
                     Me.txt_proveedor = IIf(IsNull(rsaux!VENDOR_Name), "", rsaux!VENDOR_Name)
                     Set list_item = lv_entradas.ListItems.Add(, , rsaux!item_number)
                     list_item.SubItems(1) = IIf(IsNull(rsaux!item_description), "", rsaux!item_description)
                     list_item.SubItems(2) = Format(IIf(IsNull(rsaux!quantity_shipped), 0, rsaux!quantity_shipped), "###,###,##0.00")
                     list_item.SubItems(3) = Format(0, "###,###,##0.00")
                     var_cantidad_enviada = var_cantidad_enviada + IIf(IsNull(rsaux!quantity_shipped), 0, rsaux!quantity_shipped)
                     list_item.SubItems(4) = Format(0, "###,###,##0.00")
                     list_item.SubItems(5) = Format(IIf(IsNull(rsaux!quantity_shipped), 0, rsaux!quantity_shipped), "###,###,##0.00")
                     list_item.SubItems(6) = IIf(IsNull(rsaux!shipment_header_id), 0, rsaux!shipment_header_id)
                     list_item.SubItems(7) = IIf(IsNull(rsaux!SHIPMENT_LINE_ID), 0, rsaux!SHIPMENT_LINE_ID)
                     list_item.SubItems(8) = IIf(IsNull(rsaux!ITEM_ID), "", rsaux!ITEM_ID)
                     list_item.SubItems(9) = IIf(IsNull(rsaux!ship_to_organization_id), 0, rsaux!ship_to_organization_id)
                     list_item.SubItems(10) = var_almacen_global
                     list_item.SubItems(11) = "VENDOR"
                     list_item.SubItems(12) = IIf(IsNull(rsaux!ORG_ID), 0, rsaux!ORG_ID)
                     list_item.SubItems(14) = IIf(IsNull(rsaux!vendor_site_id), 0, rsaux!vendor_site_id)
                     list_item.SubItems(15) = IIf(IsNull(rsaux!vendor_id), "", rsaux!vendor_id)
                     list_item.SubItems(16) = IIf(IsNull(rsaux!unit_meas_lookup_code), "", rsaux!unit_meas_lookup_code)
                     list_item.SubItems(17) = IIf(IsNull(rsaux!line_num), "", rsaux!line_num)
                     list_item.SubItems(18) = IIf(IsNull(rsaux!PROMISED_DATE), Date, rsaux!PROMISED_DATE)
                     list_item.SubItems(19) = IIf(IsNull(rsaux!deliver_to_person_id), "", rsaux!deliver_to_person_id)
                     list_item.SubItems(20) = IIf(IsNull(rsaux!line_location_id), "", rsaux!line_location_id)
                     list_item.SubItems(21) = IIf(IsNull(rsaux!ship_to_location_id), "", rsaux!ship_to_location_id)
                     list_item.SubItems(22) = IIf(IsNull(rsaux!ORG_ID), "", rsaux!ORG_ID)
                     list_item.SubItems(23) = IIf(IsNull(rsaux!country_of_origin_code), "", rsaux!country_of_origin_code)
                     list_item.SubItems(24) = IIf(IsNull(rsaux!UOM_CODE), "", rsaux!UOM_CODE)
                     list_item.SubItems(25) = IIf(IsNull(rsaux!unit_price), 0, rsaux!unit_price)
                     rsaux.MoveNext
               Wend
               rsaux.Close
               Me.lbl_enviados = Format(var_cantidad_enviada, "###,###,##0.00")
               rsaux.Open "SELECT * FROM XXVIA_TB_RECEPCIONES WHERE shipment_num = '" + Me.txt_nota + "' and shipment_header_id = " + CStr(rs!shipment_header_id) + " and to_organization_id = " + CStr(rs!TO_organizaTion_ID) + " and receipt_source_code = '" + rs!receipt_source_code + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  While Not rsaux.EOF
                        For var_j = 1 To lv_entradas.ListItems.Count
                            lv_entradas.ListItems.Item(var_j).Selected = True
                            If rsaux!SEGMENT1 = lv_entradas.selectedItem And rsaux!line_location_id = CDbl(lv_entradas.selectedItem.SubItems(20)) Then
                               If CDbl(rsaux!folio) = CDbl(Me.txt_busqueda) Then
                                  Me.lv_entradas.selectedItem.SubItems(4) = Format(CDbl(Me.lv_entradas.selectedItem.SubItems(4)) + rsaux!Cantidad, "###,###,##0.00")
                                  VAR_ESTATUS = IIf(IsNull(rsaux!estatus), "", rsaux!estatus)
                               End If
                               var_cantidad_recibida = var_cantidad_recibida + rsaux!Cantidad
                               Me.lv_entradas.selectedItem.SubItems(3) = Format(CDbl(Me.lv_entradas.selectedItem.SubItems(3)) + rsaux!Cantidad, "###,###,##0.00")
                               Me.lv_entradas.selectedItem.SubItems(5) = Format(CDbl(Me.lv_entradas.selectedItem.SubItems(5)) - rsaux!Cantidad, "###,###,##0.00")
                            End If
                        Next var_j
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
               Else
                  
               End If
               Me.lbl_recibidos = Format(var_cantidad_recibida, "###,###,##0.00")
               var_primera_vez = 0
               If VAR_ESTATUS <> "" Then
                  Me.txt_codigo.Enabled = False
                  For var_j = 1 To lv_entradas.ListItems.Count
                      Me.lv_entradas.ListItems.Item(var_j).Selected = True
                      If VAR_ESTATUS <> "" Then
                         Me.lv_entradas.selectedItem.SubItems(13) = VAR_ESTATUS
                         Me.txt_codigo.Enabled = False
                      End If
                  Next var_j
               Else
                  Me.txt_codigo.Enabled = True
                  Me.txt_codigo.SetFocus
               End If
               Me.frm_busqueda.Visible = False
            Else
               MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            rs.Open "SELECT * FROM XXVIA_TB_rECEPCIONES WHERE FOLIO = " + Me.txt_busqueda + " AND RECEIPT_SOURCE_CODE = '" + var_clave_movimiento + "' AND TO_ORGANIZATION_ID = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               Me.lv_entradas.ListItems.Clear
               Me.txt_folio = Me.txt_busqueda
               Me.txt_nota = rs!shipment_num
               Me.txt_destino = ""
               Me.txt_proveedor = ""
               var_cantidad_enviada = 0
               var_cantidad_recibida = 0
               var_fecha_inicio = CDate(rs!FECHA_INiCIO)
               If var_clave_movimiento = "DC" Then
                  If IIf(IsNull(rs!estatus), "", rs!estatus) = "I" Then
                  'reporte de devolucion de clientes
                     cnn.BeginTrans
                     rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_RECEPCIONES", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
                     Else
                        var_consecutivo = 1
                     End If
                     rsaux.Close
                     rsaux.Open "INSERT INTO TB_TEMP_ORACLE_RECEPCIONES (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                     cnn.CommitTrans
                  
                  
                     rsaux8.Open "select * from rcv_shipment_headers where attribute15 = '" + Me.txt_folio + "'", cnnoracle_4
                  
                     If Not rsaux8.EOF Then
                        rsaux9.Open "select * from rcv_shipment_lines where shipment_header_id = " + CStr(rsaux8!shipment_header_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux9.EOF Then
                           rsaux7.Open "SELECT * FROM XXVIA_TB_rECEPCIONES WHERE FOLIO = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux7.EOF Then
                              var_nombre_cliente = rsaux7!customer_name
                              VAR_ESTABLECIMIENTO = rsaux7!site_id
                              VAR_CLIENTE = rsaux7!CUSTOMER_ID
                              VAR_AGENTE = rsaux7!vendor_id
                              VAR_USUARIO_MOV = rsaux7!USUARIO
                              MAQUINA_MOV = rsaux7!maquina
                              FECHA_INiCIO = rsaux7!FECHA_INiCIO
                              FECHA_FIN = rsaux7!FECHA_FIN
                           End If
                           rsaux7.Close
                           rsaux7.Open "SELECT address1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= " + VAR_ESTABLECIMIENTO + " AND HCSU.SITE_USE_ID = D.site_use_id AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                           If Not rsaux7.EOF Then
                              VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux7!address1), "", rsaux7!address1)
                           End If
                           rsaux7.Close
                           rsaux.Open "SELECT * FROM AR_COLLECTORS WHERE COLLECTOR_ID = '" + VAR_AGENTE + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux.EOF Then
                              var_nombre_agente = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                           End If
                           rsaux.Close
                          
                           var_almacen = rsaux9!TO_SUBINVENTORY
                           rsaux.Open "SELECT * FROM mtl_secondary_inventories WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND SECONDARY_INVENTORY_NAME = '" + IIf(IsNull(rsaux9!TO_SUBINVENTORY), "", rsaux9!TO_SUBINVENTORY) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux.EOF Then
                              var_nombre_almacen_subinventario = IIf(IsNull(rsaux!Description), "", rsaux!Description)
                           End If
                           rsaux.Close
                           rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + CStr(rsaux9!TO_organizaTion_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux.EOF Then
                              var_nombre_unidad_origen = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                           End If
                           rsaux.Close
                           
                           While Not rsaux9.EOF
                                 rsaux.Open "SELECT * FROM xxvia_system_items_b WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND INVENTORY_ITEM_ID = " + CStr(rsaux9!ITEM_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_SEGMENT2 = ""
                                 If Not rsaux.EOF Then
                                    VAR_SEGMENT2 = IIf(IsNull(rsaux!SEGMENT1), "", rsaux!SEGMENT1)
                                 End If
                                 rsaux.Close
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_RECEPCIONES (INTE_TEM_CONSECUTIVO, ORGANIZACION_DESTINO, ORGANIZACION_ORIGEN, SHIPMENT_NUM, SHIPMENT_HEADER_ID, SHIPMENT_LINE_ID,                                                SUBINVENTARIO, SEGMENT1, CANTIDAD_ENVIADA, CANTIDAD_RECIBIDA, DESCRIPCION, NOMBRE_ORGANIZACION_DESTINO, NOMBRE_ORGANIZACION_ORIGEN,NOMBRE_SUBINVENTARIO,USUARIO, MAQUINA, FECHA_INICIO, FECHA_FIN, CLIENTE_ID, ESTABLECIMIENTO_ID,VENDOR_ID, NOMBRE_CLIENTE,NOMBRE_ESTABLECIMIENTO,NOMBRE_PROVEEDOR,MOVIMIENTO, FOLIO)"
                                 var_cadena = var_cadena + "VALUES (" + CStr(var_consecutivo) + "," + CStr(IIf(IsNull(rsaux9!TO_organizaTion_ID), 0, rsaux9!TO_organizaTion_ID)) + "," + CStr(IIf(IsNull(rsaux9!FROM_ORGANIZATION_ID), 0, rsaux9!FROM_ORGANIZATION_ID)) + ",'" + Me.txt_nota + "',0,0,'" + IIf(IsNull(rsaux9!TO_SUBINVENTORY), "", rsaux9!TO_SUBINVENTORY) + "','" + VAR_SEGMENT2 + "'," + CStr(rsaux9!QUANTITY_RECEIVED) + "," + CStr(rsaux9!QUANTITY_RECEIVED) + ",'" + IIf(IsNull(rsaux9!item_description), "", rsaux9!item_description) + "','" + var_nombre_unidad_origen + "','" + var_nombre_unidad_origen + "','" + var_nombre_almacen_subinventario + "','" + var_clave_usuario_global + "','" + MAQUINA_MOV + "','" + CStr(FECHA_INiCIO) + "','" + CStr(FECHA_FIN) + "','" + VAR_CLIENTE + "','" + VAR_ESTABLECIMIENTO + "','" + VAR_AGENTE + "','" + var_nombre_cliente + "','" + VAR_NOMBRE_ESTABLECIMIENTO + "','" + var_nombre_agente + "','DEVOLUCION DE CLIENTES','" + Me.txt_folio + "')"
                                 rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                 rsaux9.MoveNext
                           Wend
                           rsaux.Open "delete from TB_TEMP_ORACLE_RECEPCIONES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and maquina is null", cnn, adOpenDynamic, adLockOptimistic
                           Set reporte = appl.OpenReport(App.Path + "\rep_oracle_recepciones_devoluciones_clientes.rpt")
                           reporte.RecordSelectionFormula = "{VW_ORACLE_RECEPCIONES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                           frmvistasprevias.cr.ReportSource = reporte
                           For ntablas = 1 To reporte.Database.Tables.Count
                               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                           Next ntablas
                           frmvistasprevias.cr.ViewReport
                           frmvistasprevias.Caption = Me.lblnombremovimiento
                           frmvistasprevias.Show 1
                           Set reporte = Nothing
                        Else
                           MsgBox "El movimiento no se a terminado de generar en ORACLE, espere un momento por favor y vuelvalo a intentar"
                        End If
                        rsaux9.Close
                     Else
                        MsgBox "El movimiento no se a terminado de generar en ORACLE, espere un momento por favor y vuelvalo a intentar"
                     End If
                     rsaux8.Close
                     Me.txt_folio = ""
                     Me.txt_nota = ""
                     Me.txt_codigo.Enabled = False
                  'fin reporte devolucion de clientes
                  Else
                     var_almacen_global = rs!TO_SUBINVENTORY
                  
                     rsaux.Open "SELECT * FROM mtl_secondary_inventories WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND SECONDARY_INVENTORY_NAME = '" + IIf(IsNull(rs!TO_SUBINVENTORY), "", rs!TO_SUBINVENTORY) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        Me.txt_destino = IIf(IsNull(rsaux!Description), "", rsaux!Description)
                     End If
                     rsaux.Close
                     var_cadena = "SELECT hcas.cust_account_id, oha.ship_to_org_id,oha.ship_to_org_id,HL.LOCATION_ID,OHA.ORG_ID, C.PRIMARY_UOM_CODE AS UOM_CODE, C.PRIMARY_UNIT_OF_MEASURE AS UOM_DESCRIPCION,A.LINE_ID,A.HEADER_ID,C.ORGANIZATION_ID , HCAS.CUST_ACCT_SITE_ID,"
                     var_cadena = var_cadena + " E.COLLECTOR_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id, OHA.REQUEST_DATE, OHA.ORDER_NUMBER, C.description, c.segment1, E.NAME, A.ORDERED_QUANTITY from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, OE_ORDER_LINES_ALL A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND OHA.order_number = '" + Me.txt_nota + "' AND A.HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND D.collector_id = e.collector_id AND LINE_CATEGORY_CODE = 'RETURN' "
                     var_cadena = var_cadena + " AND A.SHIP_FROM_ORG_ID = C.ORGANIZATION_ID AND A.FLOW_STATUS_CODE = 'AWAITING_RETURN' AND a.ship_from_org_id = " + var_unidad_organizacional
                     rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     While Not rsaux.EOF
                           Me.txt_proveedor = rsaux!customer_name
                           Set list_item = lv_entradas.ListItems.Add(, , rsaux!SEGMENT1)
                           list_item.SubItems(1) = IIf(IsNull(rsaux!Description), "", rsaux!Description)
                           list_item.SubItems(2) = Format(IIf(IsNull(rsaux!ORDERED_QUANTITY), 0, rsaux!ORDERED_QUANTITY), "###,###,##0.00")
                           list_item.SubItems(3) = Format(0, "###,###,##0.00")
                           var_cantidad_enviada = var_cantidad_enviada + IIf(IsNull(rsaux!ORDERED_QUANTITY), 0, rsaux!ORDERED_QUANTITY)
                           list_item.SubItems(4) = Format(0, "###,###,##0.00")
                           list_item.SubItems(5) = Format(IIf(IsNull(rsaux!ORDERED_QUANTITY), 0, rsaux!ORDERED_QUANTITY), "###,###,##0.00")
                           list_item.SubItems(6) = IIf(IsNull(rsaux!header_id), 0, rsaux!header_id)
                           list_item.SubItems(7) = IIf(IsNull(rsaux!line_id), 0, rsaux!line_id)
                           list_item.SubItems(8) = IIf(IsNull(rsaux!inventory_item_id), "", rsaux!inventory_item_id)
                           list_item.SubItems(9) = IIf(IsNull(rsaux!organization_id), 0, rsaux!organization_id)
                           list_item.SubItems(10) = var_almacen_global
                           list_item.SubItems(11) = "DC"
                           list_item.SubItems(12) = IIf(IsNull(rsaux!organization_id), 0, rsaux!organization_id)
                           list_item.SubItems(15) = IIf(IsNull(rsaux!collector_id), 0, rsaux!collector_id)
                           list_item.SubItems(27) = IIf(IsNull(rsaux!CUST_ACCOUNT_ID), 0, rsaux!CUST_ACCOUNT_ID)
                           list_item.SubItems(26) = IIf(IsNull(rsaux!customer_name), 0, rsaux!customer_name)
                           list_item.SubItems(24) = IIf(IsNull(rsaux!UOM_CODE), 0, rsaux!UOM_CODE)
                           list_item.SubItems(16) = IIf(IsNull(rsaux!UOM_DESCRIPCION), 0, rsaux!UOM_DESCRIPCION)
                           list_item.SubItems(21) = IIf(IsNull(rsaux!LOCATION_ID), 0, rsaux!LOCATION_ID)
                           list_item.SubItems(28) = IIf(IsNull(rsaux!ship_to_org_id), 0, rsaux!ship_to_org_id)
                           list_item.SubItems(22) = IIf(IsNull(rsaux!ORG_ID), 0, rsaux!ORG_ID)
                           rsaux.MoveNext
                     Wend
                     rsaux.Close
                  End If
               Else
                  rsaux.Open "SELECT rsh.shipment_num, rsh.shipment_header_id, rsl.shipment_line_id, rsl.creation_date, rsl.item_id, rsl.quantity_shipped, rsl.quantity_received, rsl.from_organization_id,rsl.to_organization_id, rsl.to_subinventory, rsh.receipt_source_code, a.segment1, a.description FROM rcv_shipment_headers rsh, rcv_shipment_lines rsl, xxvia_system_items_b a Where rsh.shipment_header_id = RSL.shipment_header_id AND RSL.item_id = A.inventory_item_id AND RSL.to_organization_id = A.organization_id  AND RSL.to_organization_id = " + var_unidad_organizacional + " AND rsh.shipment_num = '" + Me.txt_nota + "' AND receipt_source_code = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rsaux.EOF
                        Set list_item = lv_entradas.ListItems.Add(, , rsaux!SEGMENT1)
                        list_item.SubItems(1) = IIf(IsNull(rsaux!Description), "", rsaux!Description)
                        list_item.SubItems(2) = Format(IIf(IsNull(rsaux!quantity_shipped), 0, rsaux!quantity_shipped), "###,###,##0.00")
                        list_item.SubItems(3) = Format(0, "###,###,##0.00")
                        var_cantidad_enviada = var_cantidad_enviada + IIf(IsNull(rsaux!quantity_shipped), 0, rsaux!quantity_shipped)
                        list_item.SubItems(4) = Format(0, "###,###,##0.00")
                        list_item.SubItems(5) = Format(IIf(IsNull(rsaux!quantity_shipped), 0, rsaux!quantity_shipped), "###,###,##0.00")
                        list_item.SubItems(6) = IIf(IsNull(rsaux!shipment_header_id), 0, rsaux!shipment_header_id)
                        list_item.SubItems(7) = IIf(IsNull(rsaux!SHIPMENT_LINE_ID), 0, rsaux!SHIPMENT_LINE_ID)
                        list_item.SubItems(8) = IIf(IsNull(rsaux!ITEM_ID), "", rsaux!ITEM_ID)
                        list_item.SubItems(9) = IIf(IsNull(rsaux!TO_organizaTion_ID), 0, rsaux!TO_organizaTion_ID)
                        list_item.SubItems(10) = IIf(IsNull(rsaux!TO_SUBINVENTORY), 0, rsaux!TO_SUBINVENTORY)
                        list_item.SubItems(11) = IIf(IsNull(rsaux!receipt_source_code), 0, rsaux!receipt_source_code)
                        list_item.SubItems(12) = IIf(IsNull(rsaux!FROM_ORGANIZATION_ID), 0, rsaux!FROM_ORGANIZATION_ID)
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
               End If
               Me.lbl_enviados = Format(var_cantidad_enviada, "###,###,##0.00")
               If var_clave_movimiento <> "DC" Then
                  rsaux.Open "SELECT * FROM mtl_secondary_inventories WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND SECONDARY_INVENTORY_NAME = '" + IIf(IsNull(rs!TO_SUBINVENTORY), "", rs!TO_SUBINVENTORY) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     Me.txt_destino = IIf(IsNull(rsaux!Description), "", rsaux!Description)
                  End If
                  rsaux.Close
                  rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + CStr(rs!FROM_ORGANIZATION_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     Me.txt_proveedor = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                  End If
                  rsaux.Close
               End If
               
               
               rsaux.Open "SELECT * FROM XXVIA_TB_RECEPCIONES WHERE shipment_num = '" + Me.txt_nota + "' and shipment_header_id = " + CStr(rs!shipment_header_id) + " and to_organization_id = " + CStr(rs!TO_organizaTion_ID) + " and receipt_source_code = '" + rs!receipt_source_code + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               While Not rsaux.EOF
                     For var_j = 1 To lv_entradas.ListItems.Count
                         lv_entradas.ListItems.Item(var_j).Selected = True
                         If rsaux!SEGMENT1 = lv_entradas.selectedItem Then
                            If CDbl(rsaux!folio) = CDbl(Me.txt_busqueda) Then
                               Me.lv_entradas.selectedItem.SubItems(4) = Format(CDbl(Me.lv_entradas.selectedItem.SubItems(4)) + rsaux!Cantidad, "###,###,##0.00")
                               VAR_ESTATUS = IIf(IsNull(rsaux!estatus), "", rsaux!estatus)
                            End If
                            var_cantidad_recibida = var_cantidad_recibida + rsaux!Cantidad
                            Me.lv_entradas.selectedItem.SubItems(3) = Format(CDbl(Me.lv_entradas.selectedItem.SubItems(3)) + rsaux!Cantidad, "###,###,##0.00")
                            Me.lv_entradas.selectedItem.SubItems(5) = Format(CDbl(Me.lv_entradas.selectedItem.SubItems(5)) - rsaux!Cantidad, "###,###,##0.00")
                         End If
                     Next var_j
                     rsaux.MoveNext
               Wend
               rsaux.Close
               Me.lbl_recibidos = Format(var_cantidad_recibida, "###,###,##0.00")
               var_primera_vez = 0
               If VAR_ESTATUS <> "" Then
                  Me.txt_codigo.Enabled = False
                  For var_j = 1 To lv_entradas.ListItems.Count
                      Me.lv_entradas.ListItems.Item(var_j).Selected = True
                      If VAR_ESTATUS <> "" Then
                         Me.lv_entradas.selectedItem.SubItems(13) = VAR_ESTATUS
                         Me.txt_codigo.Enabled = False
                      End If
                  Next var_j
               Else
                  Me.txt_codigo.Enabled = True
                  Me.txt_codigo.SetFocus
               End If
               Me.frm_busqueda.Visible = False
            Else
               MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
            End If
            rs.Close
         End If
      Else
         MsgBox "Número de movimiento incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_busqueda.Visible = False
      If Me.txt_codigo.Enabled = True Then
         Me.txt_codigo.SetFocus
      End If
   End If
End Sub

Private Sub txt_busqueda_LostFocus()
   Me.frm_busqueda.Visible = False
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad_eliminar) Then
         If CDbl(Me.lv_entradas.selectedItem.SubItems(4)) - CDbl(Me.txt_cantidad_eliminar) >= 0 Then
            Me.lv_entradas.selectedItem.SubItems(3) = Format(CDbl(lv_entradas.selectedItem.SubItems(3)) - CDbl(Me.txt_cantidad_eliminar), "###,###,##0.00")
            Me.lv_entradas.selectedItem.SubItems(4) = Format(CDbl(lv_entradas.selectedItem.SubItems(4)) - CDbl(Me.txt_cantidad_eliminar), "###,###,##0.00")
            Me.lv_entradas.selectedItem.SubItems(5) = Format(CDbl(lv_entradas.selectedItem.SubItems(5)) + CDbl(Me.txt_cantidad_eliminar), "###,###,##0.00")
            rsaux1.Open "update xxvia_tb_recepciones set cantidad = cantidad -" + CStr(Me.txt_cantidad_eliminar) + " where folio = " + Me.txt_folio + " and shipment_num = '" + Me.txt_nota + "' and shipment_header_id = " + Me.lv_entradas.selectedItem.SubItems(6) + " and shipment_line_id = " + Me.lv_entradas.selectedItem.SubItems(7) + " and item_id = " + Me.lv_entradas.selectedItem.SubItems(8) + " and to_organization_id = " + Me.lv_entradas.selectedItem.SubItems(9), cnnoracle_4, adOpenDynamic, adLockOptimistic
            Me.lbl_recibidos = Format(CDbl(Me.lbl_recibidos) - CDbl(Me.txt_cantidad_eliminar), "###,###,##0.00")
            Me.lv_entradas.SetFocus
         Else
            MsgBox "La cantidad excede a la posible a eliminar", vbOKCancel, "ATENCION"
         End If
      End If
   End If
   If KeyAscii = 27 Then
      If Me.txt_codigo.Enabled = True Then
         Me.txt_codigo.SetFocus
      End If
      Me.frm_eliminar.Visible = False
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   Me.frm_eliminar.Visible = False
End Sub

Private Sub txt_codigo_GotFocus()
   'Me.txt_codigo = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_localizador_subinventario = " "
      var_encontro = 0
      var_cantidad_leida = 1
      rsaux8.Open "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, nvl(a.description,'') as localizador FROM mtl_cross_references_b A, xxvia_system_items_b B WHERE A.inventory_item_id = B.inventory_item_id AND B.organization_id = " + var_unidad_organizacional + " AND CROSS_REFERENCE = '" + Me.txt_codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rsaux8.EOF Then
         Me.txt_codigo = rsaux8!SEGMENT1
      Else
         rsaux9.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux9.EOF Then
            Me.txt_codigo = rsaux9!SEGMENT1
         Else
            Me.txt_codigo = ""
         End If
         rsaux9.Close
      End If
      rsaux8.Close
      
      var_encontro = 0
      If rsaux8.State = 1 Then
         rsaux8.Close
      End If
      rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rsaux8.EOF Then
         var_salida_masiva = IIf(IsNull(rsaux8!attribute10), "N", rsaux8!attribute10)
         If var_salida_masiva = "Y" Then
            var_codigo_global = Me.txt_codigo
            frmoracle_cantidad.Show 1
            var_cantidad_leida = var_cantidad_global
            Me.txt_codigo = var_codigo_global
         Else
            var_cantidad_leida = 1
         End If
         For var_j = 1 To Me.lv_entradas.ListItems.Count
             lv_entradas.ListItems.Item(var_j).Selected = True
             If Me.txt_codigo = lv_entradas.selectedItem And (CDbl(Me.lv_entradas.selectedItem.SubItems(5)) - var_cantidad_leida) >= 0 Then
                var_encontro = var_j
             End If
         Next var_j
         If var_encontro > 0 Then
            Me.lv_entradas.ListItems.Item(var_encontro).Selected = True
            If CDbl(Me.lv_entradas.selectedItem.SubItems(2)) >= CDbl(Me.lv_entradas.selectedItem.SubItems(3)) + var_cantidad_leida Then
               Me.txt_foco.Enabled = True
               Me.txt_foco.SetFocus
            Else
               Call cmd_mensaje_2_Click
               txt_codigo = ""
               frmmensaje.lbl_articulo = Me.lv_entradas.selectedItem.SubItems(1)
               frmmensaje.lbl_mensaje = "La cantidad supera a la enviada"
               frmmensaje.Show 1
            End If
         Else
            Call cmd_mensaje_2_Click
            txt_codigo = ""
            frmmensaje.lbl_articulo = ""
            frmmensaje.lbl_mensaje = "El artículo no se encuentra en la nota"
            frmmensaje.Show 1
         End If
      Else
         Call cmd_mensaje_2_Click
         txt_codigo = ""
         frmmensaje.lbl_articulo = ""
         frmmensaje.lbl_mensaje = "El artículo no existe"
         frmmensaje.Show 1
      End If
      rsaux8.Close
   End If
End Sub


Private Sub txt_factura_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_factura <> "" Then
         Me.txt_factura.Enabled = False
         Me.txt_codigo.Enabled = True
         Me.txt_codigo.SetFocus
      Else
         MsgBox "Debe de indicar una factura", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   If Me.txt_codigo <> "" Then
      If var_primera_vez = 1 Then
         var_primera_vez = 0
         cnnoracle_4.BeginTrans
         rs.Open "select * from xxvia_tb_folios_entradas", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_folio = IIf(IsNull(rs!folio), 0, rs!folio) + 1
            rsaux.Open "update xxvia_tb_folios_entradas SET FOLIO = " + CStr(var_folio), cnnoracle_4, adOpenDynamic, adLockOptimistic
         Else
            var_folio = 1
            rsaux.Open "insert INTO xxvia_tb_folios_entradas (folio) values (" + CStr(var_folio) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
         rs.Close
         Me.txt_folio = var_folio
         cnnoracle_4.CommitTrans
         var_fecha_inicio = Now
      End If
      rsaux.Open "select * from xxvia_tb_recepciones where folio = " + Me.txt_folio + " and shipment_num = '" + Me.txt_nota + "' and shipment_header_id = " + Me.lv_entradas.selectedItem.SubItems(6) + " and shipment_line_id = " + Me.lv_entradas.selectedItem.SubItems(7) + " and item_id = " + Me.lv_entradas.selectedItem.SubItems(8) + " and to_organization_id = " + Me.lv_entradas.selectedItem.SubItems(9), cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         rsaux1.Open "update xxvia_tb_recepciones set cantidad = cantidad +" + CStr(var_cantidad_leida) + " where folio = " + Me.txt_folio + " and shipment_num = '" + Me.txt_nota + "' and shipment_header_id = " + Me.lv_entradas.selectedItem.SubItems(6) + " and shipment_line_id = " + Me.lv_entradas.selectedItem.SubItems(7) + " and item_id = " + Me.lv_entradas.selectedItem.SubItems(8) + " and to_organization_id = " + Me.lv_entradas.selectedItem.SubItems(9), cnnoracle_4, adOpenDynamic, adLockOptimistic
      Else
         var_segundo_s = CStr(Second(var_fecha_inicio))
         var_minuto_s = CStr(Minute(var_fecha_inicio))
         var_hora_s = CStr(Hour(var_fecha_inicio))
         var_año_s = CStr(Year(var_fecha_inicio))
         var_mes_s = CStr(Month(var_fecha_inicio))
         var_dia_s = CStr(Day(var_fecha_inicio))
         If Len(var_segundo_s) = 1 Then
            var_segundo_s = "0" + var_segundo_s
         End If
         If Len(var_minuto_s) = 1 Then
            var_minuto_s = "0" + var_minuto_s
         End If
         If Len(var_hora_s) = 1 Then
            var_hora_s = "0" + var_hora_s
         End If
         If Len(var_año_s) = 2 Then
            var_año_s = "20" + var_año_s
         End If
         If Len(var_mes_s) = 1 Then
            var_mes_s = "0" + var_mes_s
         End If
         If Len(var_dia_s) = 1 Then
            var_dia_s = "0" + var_dia_s
         End If
         var_fecha_str = var_año_s + "/" + var_mes_s + "/" + var_dia_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
         var_cadena = "insert into xxvia_tb_Recepciones (folio, shipment_num, shipment_header_id, shipment_line_id,                                                                                                               item_id, cantidad, from_organization_id, to_organization_id, to_subinventory, receipt_source_code, segment1, usuario, maquina, fecha_inicio, vendor_site_id, vendor_id, medida, line_num, FECHA_PROMESA, DELIVER_TO_PERSON_ID, line_location_id,  ship_to_location_id, ORG_ID,  country_of_origin_code, UOM_CODE, UNIT_PRICE, CUSTOMER_ID, CUSTOMER_NAME,SITE_ID, FACTURA) "
         If var_clave_movimiento = "DC" Then
            var_cadena = var_cadena + " values (" + Me.txt_folio + ",'" + Me.txt_nota + "'," + Me.lv_entradas.selectedItem.SubItems(6) + "," + Me.lv_entradas.selectedItem.SubItems(7) + "," + Me.lv_entradas.selectedItem.SubItems(8) + "," + CStr(var_cantidad_leida) + "," + Me.lv_entradas.selectedItem.SubItems(12) + "," + Me.lv_entradas.selectedItem.SubItems(9) + ",'" + Me.lv_entradas.selectedItem.SubItems(10) + "','" + Me.lv_entradas.selectedItem.SubItems(11) + "','" + Me.lv_entradas.selectedItem + "','" + var_clave_usuario_global + "','" + fun_NombrePc + "',to_date('" + var_fecha_str + "','yyyy/mm/dd hh24:mi:ss'),'" + Me.lv_entradas.selectedItem.SubItems(14) + "','" + Me.lv_entradas.selectedItem.SubItems(15) + "','" + Me.lv_entradas.selectedItem.SubItems(16) + "','" + Me.lv_entradas.selectedItem.SubItems(17) + "', TO_DATE('" + Format(CDate(Date), "Short Date") + "','DD/MM/YYYY'),'" + Me.lv_entradas.selectedItem.SubItems(19) + "','" + Me.lv_entradas.selectedItem.SubItems(20)
            var_cadena = var_cadena + "','" + Me.lv_entradas.selectedItem.SubItems(21) + "','" + Me.lv_entradas.selectedItem.SubItems(22) + "','" + Me.lv_entradas.selectedItem.SubItems(23) + "','" + Me.lv_entradas.selectedItem.SubItems(24) + "',0,'" + Me.lv_entradas.selectedItem.SubItems(27) + "','" + Me.lv_entradas.selectedItem.SubItems(26) + "'," + Me.lv_entradas.selectedItem.SubItems(28) + ",'" + Me.txt_factura + "')"
         Else
            var_cadena = var_cadena + " values (" + Me.txt_folio + ",'" + Me.txt_nota + "'," + Me.lv_entradas.selectedItem.SubItems(6) + "," + Me.lv_entradas.selectedItem.SubItems(7) + "," + Me.lv_entradas.selectedItem.SubItems(8) + "," + CStr(var_cantidad_leida) + "," + Me.lv_entradas.selectedItem.SubItems(12) + "," + Me.lv_entradas.selectedItem.SubItems(9) + ",'" + Me.lv_entradas.selectedItem.SubItems(10) + "','" + Me.lv_entradas.selectedItem.SubItems(11) + "','" + Me.lv_entradas.selectedItem + "','" + var_clave_usuario_global + "','" + fun_NombrePc + "',to_date('" + var_fecha_str + "','yyyy/mm/dd hh24:mi:ss'),'" + Me.lv_entradas.selectedItem.SubItems(14) + "','" + Me.lv_entradas.selectedItem.SubItems(15) + "','" + Me.lv_entradas.selectedItem.SubItems(16) + "','" + Me.lv_entradas.selectedItem.SubItems(17) + "', TO_DATE('" + Format(CDate(Me.lv_entradas.selectedItem.SubItems(18)), "Short Date") + "','DD/MM/YYYY'),'" + Me.lv_entradas.selectedItem.SubItems(19) + "','" + Me.lv_entradas.selectedItem.SubItems(20)
            var_cadena = var_cadena + "','" + Me.lv_entradas.selectedItem.SubItems(21) + "','" + Me.lv_entradas.selectedItem.SubItems(22) + "','" + Me.lv_entradas.selectedItem.SubItems(23) + "','" + Me.lv_entradas.selectedItem.SubItems(24) + "'," + Me.lv_entradas.selectedItem.SubItems(25) + ",'','','','" + Me.txt_factura + "')"
         End If
         'MsgBox var_cadena
         rsaux1.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         
      End If
      Me.lbl_recibidos = Format(CDbl(Me.lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
      Me.lv_entradas.selectedItem.SubItems(3) = Format(CDbl(Me.lv_entradas.selectedItem.SubItems(3)) + var_cantidad_leida, "###,###,##0.00")
      Me.lv_entradas.selectedItem.SubItems(4) = Format(CDbl(Me.lv_entradas.selectedItem.SubItems(4)) + var_cantidad_leida, "###,###,##0.00")
      Me.lv_entradas.selectedItem.SubItems(5) = Format(CDbl(Me.lv_entradas.selectedItem.SubItems(5)) - var_cantidad_leida, "###,###,##0.00")
      var_renglon = Me.lv_entradas.selectedItem.Index
      Call ilumina_grid
      rsaux.Close
      Me.txt_codigo = ""
      Me.txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_nota_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim list_item As ListItem
      If Trim(Me.txt_nota) <> "" Then
         If var_clave_movimiento = "VENDOR" Then
            If rsaux8.State = 1 Then
               rsaux8.Close
            End If
            rsaux8.Open "select distinct folio as folio from xxvia_tb_recepciones where shipment_num = '" + Me.txt_nota + "' and to_organization_id = " + var_unidad_organizacional + " and receipt_source_code = '" + var_clave_movimiento + "' and (estatus <> 'I' or estatus is null)", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
               var_folio_existente = ""
               While Not rsaux8.EOF
                     If var_folio_existente = "" Then
                        var_folio_existente = CStr(IIf(IsNull(rsaux8!folio), "", rsaux8!folio))
                     Else
                        var_folio_existente = var_folio_existente + ", " + CStr(IIf(IsNull(rsaux8!folio), "", rsaux8!folio))
                     End If
                     rsaux8.MoveNext
               Wend
               MsgBox "Existe un movimiento abierto. Folios: " + var_folio_existente
            Else
               rs.Open "select deliver_to_person_id,  PROMISED_DATE, unit_meas_lookup_code, line_num, vendor_id, vendor_site_id, num_oc SHIPMENT_NUM, PO_HEADER_ID SHIPMENT_HEADER_ID, PO_LINE_ID SHIPMENT_LINE_ID,  ITEM_ID, quantity AS QUANTITY_SHIPPED, 0 QUANTITY_RECEIVED, VENDOR_ID, VENDOR_NAME, item_number, ITEM_DESCRIPTION, ORG_ID , SHIP_TO_ORGANIZATION_ID, line_location_id, ship_to_location_id, country_of_origin_code, UOM_CODE, UNIT_PRICE from xxvia_vw_recepcion_compra where num_oc = " + Me.txt_nota + " AND SHIP_TO_ORGANIZATION_ID = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_almacen_global = ""
                  frmoracle_subinventarios.Show 1
                  If var_almacen_global <> "" Then
                     Me.txt_destino = var_nombre_almacen_global
                     Me.lv_entradas.ListItems.Clear
                     Me.txt_folio = ""
                     Me.txt_proveedor = ""
                     var_cantidad_enviada = 0
                     var_cantidad_recibida = 0
                     While Not rs.EOF
                           Set list_item = lv_entradas.ListItems.Add(, , IIf(IsNull(rs!item_number), "", rs!item_number))
                           list_item.SubItems(1) = IIf(IsNull(rs!item_description), "", rs!item_description)
                           list_item.SubItems(2) = Format(IIf(IsNull(rs!quantity_shipped), 0, rs!quantity_shipped), "###,###,##0.00")
                           list_item.SubItems(3) = Format(0, "###,###,##0.00")
                           var_cantidad_enviada = var_cantidad_enviada + IIf(IsNull(rs!quantity_shipped), 0, rs!quantity_shipped)
                           list_item.SubItems(4) = Format(0, "###,###,##0.00")
                           list_item.SubItems(5) = Format(IIf(IsNull(rs!quantity_shipped), 0, rs!quantity_shipped), "###,###,##0.00")
                           list_item.SubItems(6) = IIf(IsNull(rs!shipment_header_id), 0, rs!shipment_header_id)
                           list_item.SubItems(7) = IIf(IsNull(rs!SHIPMENT_LINE_ID), 0, rs!SHIPMENT_LINE_ID)
                           list_item.SubItems(8) = IIf(IsNull(rs!ITEM_ID), "", rs!ITEM_ID)
                           list_item.SubItems(9) = IIf(IsNull(rs!ship_to_organization_id), 0, rs!ship_to_organization_id)
                           list_item.SubItems(10) = var_almacen_global
                           list_item.SubItems(11) = "VENDOR"
                           list_item.SubItems(12) = IIf(IsNull(rs!ORG_ID), 0, rs!ORG_ID)
                           list_item.SubItems(14) = IIf(IsNull(rs!vendor_site_id), "", rs!vendor_site_id)
                           list_item.SubItems(15) = IIf(IsNull(rs!vendor_id), "", rs!vendor_id)
                           list_item.SubItems(16) = IIf(IsNull(rs!unit_meas_lookup_code), "", rs!unit_meas_lookup_code)
                           list_item.SubItems(17) = IIf(IsNull(rs!line_num), "", rs!line_num)
                           list_item.SubItems(18) = IIf(IsNull(rs!PROMISED_DATE), Date, rs!PROMISED_DATE)
                           list_item.SubItems(19) = IIf(IsNull(rs!deliver_to_person_id), "", rs!deliver_to_person_id)
                           list_item.SubItems(20) = IIf(IsNull(rs!line_location_id), "", rs!line_location_id)
                           list_item.SubItems(21) = IIf(IsNull(rs!ship_to_location_id), "", rs!ship_to_location_id)
                           list_item.SubItems(22) = IIf(IsNull(rs!ORG_ID), "", rs!ORG_ID)
                           list_item.SubItems(23) = IIf(IsNull(rs!country_of_origin_code), "", rs!country_of_origin_code)
                           list_item.SubItems(24) = IIf(IsNull(rs!UOM_CODE), "", rs!UOM_CODE)
                           list_item.SubItems(25) = IIf(IsNull(rs!unit_price), 0, rs!unit_price)
                           rs.MoveNext
                     Wend
                     rs.MoveFirst
                     Me.lbl_enviados = Format(var_cantidad_enviada, "###,###,##0.00")
                     'rsaux.Open "SELECT * FROM mtl_secondary_inventories WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND SECONDARY_INVENTORY_NAME = '" + IIf(IsNull(rs!to_subinventory), "", rs!to_subinventory) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     'If Not rsaux.EOF Then
                     '   Me.txt_destino = IIf(IsNull(rsaux!Description), "", rsaux!Description)
                     'End If
                     'rsaux.Close
                     Me.txt_proveedor = IIf(IsNull(rs!VENDOR_Name), "", rs!VENDOR_Name)
                     rsaux.Open "SELECT * FROM XXVIA_TB_RECEPCIONES WHERE shipment_num = '" + Me.txt_nota + "' and shipment_header_id = " + CStr(rs!shipment_header_id) + " and to_organization_id = " + CStr(rs!ship_to_organization_id) + " and receipt_source_code = 'VENDOR'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_cantidad_recibida = 0
                     While Not rsaux.EOF
                           For var_j = 1 To lv_entradas.ListItems.Count
                               lv_entradas.ListItems.Item(var_j).Selected = True
                               If rsaux!SEGMENT1 = lv_entradas.selectedItem And rsaux!line_location_id = Me.lv_entradas.selectedItem.SubItems(20) Then
                                  var_cantidad_recibida = var_cantidad_recibida + rsaux!Cantidad
                                  Me.lv_entradas.selectedItem.SubItems(3) = Format(CDbl(Me.lv_entradas.selectedItem.SubItems(3)) + rsaux!Cantidad, "###,###,##0.00")
                                  Me.lv_entradas.selectedItem.SubItems(5) = Format(CDbl(Me.lv_entradas.selectedItem.SubItems(5)) - rsaux!Cantidad, "###,###,##0.00")
                               End If
                           Next var_j
                           rsaux.MoveNext
                     Wend
                     rsaux.Close
                     Me.lbl_recibidos = Format(var_cantidad_recibida, "###,###,##0.00")
                     Me.txt_nota.Enabled = False
                     Me.txt_factura.Enabled = True
                     Me.txt_factura.SetFocus
                     'Me.txt_codigo.Enabled = True
                     'Me.txt_codigo.SetFocus
                  Else
                     Me.txt_factura.Enabled = False
                     Me.txt_proveedor = ""
                     Me.lv_entradas.ListItems.Clear
                     Me.txt_codigo.Enabled = False
                     Me.txt_codigo = ""
                     Me.txt_destino = ""
                     MsgBox "No se se lecciono un almacén", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "La orden de compra no existe", vbOKOnly, "ATENCION"
               End If
               rs.Close
            End If
            rsaux8.Close
         End If
         If var_clave_movimiento = "INVENTORY" Or var_clave_movimiento = "INTERNAL ORDER" Then
            rsaux8.Open "select distinct folio as folio from xxvia_tb_recepciones where shipment_num = '" + Me.txt_nota + "' and to_organization_id = " + var_unidad_organizacional + " and receipt_source_code = '" + var_clave_movimiento + "' and (estatus <> 'I' or estatus is null)", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
               var_folio_existente = ""
               While Not rsaux8.EOF
                     If var_folio_existente = "" Then
                        var_folio_existente = CStr(IIf(IsNull(rsaux8!folio), "", rsaux8!folio))
                     Else
                        var_folio_existente = var_folio_existente + ", " + CStr(IIf(IsNull(rsaux8!folio), "", rsaux8!folio))
                     End If
                     rsaux8.MoveNext
               Wend
               MsgBox "Existe un movimiento abierto. Folios: " + var_folio_existente
            Else
               rs.Open "SELECT rsh.shipment_num, rsh.shipment_header_id, rsl.shipment_line_id, rsl.creation_date, rsl.item_id, rsl.quantity_shipped, rsl.quantity_received, rsl.from_organization_id,rsl.to_organization_id, rsl.to_subinventory, rsh.receipt_source_code, a.segment1, a.description FROM rcv_shipment_headers rsh, rcv_shipment_lines rsl, xxvia_system_items_b a Where rsh.shipment_header_id = RSL.shipment_header_id AND RSL.item_id = A.inventory_item_id AND RSL.to_organization_id = A.organization_id  AND RSL.to_organization_id = " + var_unidad_organizacional + " AND rsh.shipment_num = '" + Me.txt_nota + "' AND receipt_source_code = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  Me.lv_entradas.ListItems.Clear
                  Me.txt_folio = ""
                  Me.txt_destino = ""
                  Me.txt_proveedor = ""
                  var_cantidad_enviada = 0
                  var_cantidad_recibida = 0
                  While Not rs.EOF
                        Set list_item = lv_entradas.ListItems.Add(, , rs!SEGMENT1)
                        list_item.SubItems(1) = IIf(IsNull(rs!Description), "", rs!Description)
                        list_item.SubItems(2) = Format(IIf(IsNull(rs!quantity_shipped), 0, rs!quantity_shipped), "###,###,##0.00")
                        list_item.SubItems(3) = Format(0, "###,###,##0.00")
                        var_cantidad_enviada = var_cantidad_enviada + IIf(IsNull(rs!quantity_shipped), 0, rs!quantity_shipped)
                        list_item.SubItems(4) = Format(0, "###,###,##0.00")
                        list_item.SubItems(5) = Format(IIf(IsNull(rs!quantity_shipped), 0, rs!quantity_shipped), "###,###,##0.00")
                        list_item.SubItems(6) = IIf(IsNull(rs!shipment_header_id), 0, rs!shipment_header_id)
                        list_item.SubItems(7) = IIf(IsNull(rs!SHIPMENT_LINE_ID), 0, rs!SHIPMENT_LINE_ID)
                        list_item.SubItems(8) = IIf(IsNull(rs!ITEM_ID), "", rs!ITEM_ID)
                        list_item.SubItems(9) = IIf(IsNull(rs!TO_organizaTion_ID), 0, rs!TO_organizaTion_ID)
                        list_item.SubItems(10) = IIf(IsNull(rs!TO_SUBINVENTORY), 0, rs!TO_SUBINVENTORY)
                        list_item.SubItems(11) = IIf(IsNull(rs!receipt_source_code), 0, rs!receipt_source_code)
                        list_item.SubItems(12) = IIf(IsNull(rs!FROM_ORGANIZATION_ID), 0, rs!FROM_ORGANIZATION_ID)
                        rs.MoveNext
                  Wend
                  rs.MoveFirst
                  Me.lbl_enviados = Format(var_cantidad_enviada, "###,###,##0.00")
                  rsaux.Open "SELECT * FROM mtl_secondary_inventories WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND SECONDARY_INVENTORY_NAME = '" + IIf(IsNull(rs!TO_SUBINVENTORY), "", rs!TO_SUBINVENTORY) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     Me.txt_destino = IIf(IsNull(rsaux!Description), "", rsaux!Description)
                  End If
                  rsaux.Close
                  rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + CStr(rs!FROM_ORGANIZATION_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     Me.txt_proveedor = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                  End If
                  rsaux.Close
                  rsaux.Open "SELECT * FROM XXVIA_TB_RECEPCIONES WHERE shipment_num = '" + Me.txt_nota + "' and shipment_header_id = " + CStr(rs!shipment_header_id) + " and to_organization_id = " + CStr(rs!TO_organizaTion_ID) + " and receipt_source_code = '" + rs!receipt_source_code + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  var_cantidad_recibida = 0
                  While Not rsaux.EOF
                        For var_j = 1 To lv_entradas.ListItems.Count
                            lv_entradas.ListItems.Item(var_j).Selected = True
                            If rs!SEGMENT1 = lv_entradas.selectedItem Then
                               var_cantidad_recibida = var_cantidad_recibida + rsaux!Cantidad
                               Me.lv_entradas.selectedItem.SubItems(3) = Format(CDbl(Me.lv_entradas.selectedItem.SubItems(3)) + rsaux!Cantidad, "###,###,##0.00")
                               Me.lv_entradas.selectedItem.SubItems(5) = Format(CDbl(Me.lv_entradas.selectedItem.SubItems(5)) - rsaux!Cantidad, "###,###,##0.00")
                            End If
                        Next var_j
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  Me.lbl_recibidos = Format(var_cantidad_recibida, "###,###,##0.00")
                  Me.txt_nota.Enabled = False
                  Me.txt_codigo.Enabled = True
                  Me.txt_codigo.SetFocus
               Else
                  MsgBox "No se a encontrado la nota", vbOKOnly, "ATENCION"
               End If
               rs.Close
            End If
            rsaux8.Close
         End If
         If var_clave_movimiento = "DC" Then
            rsaux8.Open "select distinct folio as folio from xxvia_tb_recepciones where shipment_num = '" + Me.txt_nota + "' and to_organization_id = " + var_unidad_organizacional + " and receipt_source_code = '" + var_clave_movimiento + "' and (estatus <> 'I' or estatus is null)", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
               var_folio_existente = ""
               While Not rsaux8.EOF
                     If var_folio_existente = "" Then
                        var_folio_existente = CStr(IIf(IsNull(rsaux8!folio), "", rsaux8!folio))
                     Else
                        var_folio_existente = var_folio_existente + ", " + CStr(IIf(IsNull(rsaux8!folio), "", rsaux8!folio))
                     End If
                     rsaux8.MoveNext
               Wend
               MsgBox "Existe un movimiento abierto. Folios: " + var_folio_existente
            Else
               var_cadena = "SELECT e.collector_id, hcas.cust_account_id, oha.ship_to_org_id,oha.invoice_to_org_id,HL.LOCATION_ID, OHA.ORG_ID,C.PRIMARY_UOM_CODE AS UOM_CODE, C.PRIMARY_UNIT_OF_MEASURE AS UOM_DESCRIPCION, A.LINE_ID, A.HEADER_ID, C.ORGANIZATION_ID , HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id, OHA.REQUEST_DATE, OHA.ORDER_NUMBER, C.description, c.segment1, E.NAME, A.ORDERED_QUANTITY from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, OE_ORDER_LINES_ALL A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND OHA.order_number = '" + Me.txt_nota + "' AND A.HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND "
               var_cadena = var_cadena + " A.inventory_item_id  = c.inventory_item_id AND D.collector_id = e.collector_id AND LINE_CATEGORY_CODE = 'RETURN' "
               var_cadena = var_cadena + " AND A.SHIP_FROM_ORG_ID = C.ORGANIZATION_ID AND A.FLOW_STATUS_CODE = 'AWAITING_RETURN' AND a.ship_from_org_id = " + var_unidad_organizacional
               rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_almacen_global = ""
                  frmoracle_subinventarios.Show 1
                  If var_almacen_global <> "" Then
                     Me.txt_destino = var_nombre_almacen_global
                     Me.lv_entradas.ListItems.Clear
                     Me.txt_folio = ""
                     Me.txt_destino = ""
                     Me.txt_proveedor = ""
                     var_cantidad_enviada = 0
                     var_cantidad_recibida = 0
                     While Not rs.EOF
                           Set list_item = lv_entradas.ListItems.Add(, , rs!SEGMENT1)
                           list_item.SubItems(1) = IIf(IsNull(rs!Description), "", rs!Description)
                           list_item.SubItems(2) = Format(IIf(IsNull(rs!ORDERED_QUANTITY), 0, rs!ORDERED_QUANTITY), "###,###,##0.00")
                           list_item.SubItems(3) = Format(0, "###,###,##0.00")
                           var_cantidad_enviada = var_cantidad_enviada + IIf(IsNull(rs!ORDERED_QUANTITY), 0, rs!ORDERED_QUANTITY)
                           list_item.SubItems(4) = Format(0, "###,###,##0.00")
                           list_item.SubItems(5) = Format(IIf(IsNull(rs!ORDERED_QUANTITY), 0, rs!ORDERED_QUANTITY), "###,###,##0.00")
                           list_item.SubItems(6) = IIf(IsNull(rs!header_id), 0, rs!header_id)
                           list_item.SubItems(7) = IIf(IsNull(rs!line_id), 0, rs!line_id)
                           list_item.SubItems(8) = IIf(IsNull(rs!inventory_item_id), "", rs!inventory_item_id)
                           list_item.SubItems(9) = IIf(IsNull(rs!organization_id), 0, rs!organization_id)
                           list_item.SubItems(10) = var_almacen_global
                           list_item.SubItems(11) = "DC"
                           list_item.SubItems(12) = IIf(IsNull(rs!organization_id), 0, rs!organization_id)
                           list_item.SubItems(15) = IIf(IsNull(rs!collector_id), 0, rs!collector_id)
                           list_item.SubItems(27) = IIf(IsNull(rs!CUST_ACCOUNT_ID), 0, rs!CUST_ACCOUNT_ID)
                           list_item.SubItems(26) = IIf(IsNull(rs!customer_name), 0, rs!customer_name)
                           list_item.SubItems(24) = IIf(IsNull(rs!UOM_CODE), 0, rs!UOM_CODE)
                           list_item.SubItems(16) = IIf(IsNull(rs!UOM_DESCRIPCION), 0, rs!UOM_DESCRIPCION)
                           list_item.SubItems(21) = IIf(IsNull(rs!LOCATION_ID), 0, rs!LOCATION_ID)
                           list_item.SubItems(28) = IIf(IsNull(rs!ship_to_org_id), 0, rs!ship_to_org_id)
                           list_item.SubItems(22) = IIf(IsNull(rs!ORG_ID), 0, rs!ORG_ID)
                           list_item.SubItems(22) = IIf(IsNull(rs!ORG_ID), 0, rs!ORG_ID)
                           rs.MoveNext
                     Wend
                     rs.MoveFirst
                     Me.lbl_enviados = Format(var_cantidad_enviada, "###,###,##0.00")
                     rsaux.Open "SELECT * FROM mtl_secondary_inventories WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND SECONDARY_INVENTORY_NAME = '" + var_almacen_global + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        Me.txt_destino = IIf(IsNull(rsaux!Description), "", rsaux!Description)
                     End If
                     rsaux.Close
                     Me.txt_proveedor = rs!customer_name
                     rsaux.Open "SELECT * FROM XXVIA_TB_RECEPCIONES WHERE shipment_num = '" + Me.txt_nota + "' and shipment_header_id = " + CStr(rs!header_id) + " and to_organization_id = " + CStr(rs!organization_id) + " and receipt_source_code = 'DC'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_cantidad_recibida = 0
                     While Not rsaux.EOF
                           For var_j = 1 To lv_entradas.ListItems.Count
                               lv_entradas.ListItems.Item(var_j).Selected = True
                               If rsaux!SEGMENT1 = lv_entradas.selectedItem Then
                                  var_cantidad_recibida = var_cantidad_recibida + rsaux!Cantidad
                                  Me.lv_entradas.selectedItem.SubItems(3) = Format(CDbl(Me.lv_entradas.selectedItem.SubItems(3)) + rsaux!Cantidad, "###,###,##0.00")
                                  Me.lv_entradas.selectedItem.SubItems(5) = Format(CDbl(Me.lv_entradas.selectedItem.SubItems(5)) - rsaux!Cantidad, "###,###,##0.00")
                               End If
                           Next var_j
                           rsaux.MoveNext
                     Wend
                     rsaux.Close
                     Me.lbl_recibidos = Format(var_cantidad_recibida, "###,###,##0.00")
                     Me.txt_nota.Enabled = False
                     Me.txt_codigo.Enabled = True
                     Me.txt_codigo.SetFocus
                  Else
                     Me.txt_proveedor = ""
                     Me.lv_entradas.ListItems.Clear
                     Me.txt_codigo.Enabled = False
                     Me.txt_codigo = ""
                     Me.txt_destino = ""
                     MsgBox "No se se lecciono un almacén", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "No se a encontrado la nota", vbOKOnly, "ATENCION"
               End If
               rs.Close
            End If
            rsaux8.Close
         End If
      Else
         MsgBox "No se a seleccionado una nota", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

