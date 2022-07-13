VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmovimientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   ControlBox      =   0   'False
   Icon            =   "frmmovimientos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   180
      Picture         =   "frmmovimientos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   495
      Picture         =   "frmmovimientos.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   825
      Picture         =   "frmmovimientos.frx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1155
      Picture         =   "frmmovimientos.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1485
      Picture         =   "frmmovimientos.frx":0CA2
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11205
      Picture         =   "frmmovimientos.frx":0DA4
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   6255
      Left            =   5880
      TabIndex        =   24
      Top             =   960
      Width           =   5655
      Begin MSComctlLib.ListView lv_movimientos 
         Height          =   6045
         Left            =   30
         TabIndex        =   25
         Top             =   135
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   10663
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
         NumItems        =   25
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Afectacion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "hace referencia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "referencia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "requiere factura"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "folio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "documento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "causa devolucion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "dependencia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "clase"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "itercompañia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "relectura"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "aceptar de mas"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "tipo proveedor"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Tipo cliente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "promedia costo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Devolucion factura"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "reporte"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "ajuste"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "ultimo costo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "reempaque"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "ajuste reempaque"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "sobrantes"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "agurapador"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Movimientos "
      Height          =   6810
      Left            =   165
      TabIndex        =   15
      Top             =   420
      Width           =   5655
      Begin VB.TextBox txt_agrupador 
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   56
         Top             =   6000
         Width           =   4170
      End
      Begin VB.CheckBox chk_sobrante 
         Caption         =   "Sobrantes"
         Height          =   210
         Left            =   1380
         TabIndex        =   55
         Top             =   5760
         Width           =   1425
      End
      Begin VB.CheckBox chk_ajuste_reempaque 
         Caption         =   "Ajuste de Reempaque"
         Height          =   210
         Left            =   2820
         TabIndex        =   54
         Top             =   5490
         Width           =   2655
      End
      Begin VB.CheckBox chk_reempaque 
         Caption         =   "Reempaque"
         Height          =   210
         Left            =   1380
         TabIndex        =   53
         Top             =   5490
         Width           =   1425
      End
      Begin VB.CheckBox chk_ultimo_costo 
         Caption         =   "Ultimo Costo"
         Height          =   210
         Left            =   2820
         TabIndex        =   52
         Top             =   5235
         Width           =   2655
      End
      Begin VB.CheckBox chk_ajuste 
         Caption         =   "Ajuste"
         Height          =   210
         Left            =   1380
         TabIndex        =   51
         Top             =   5220
         Width           =   1425
      End
      Begin VB.TextBox txt_reporte 
         Height          =   315
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   49
         Top             =   4830
         Width           =   4140
      End
      Begin VB.CheckBox chk_devolucion_factura 
         Caption         =   "Devolucion de factura"
         Height          =   210
         Left            =   2850
         TabIndex        =   48
         Top             =   4605
         Width           =   2655
      End
      Begin VB.CheckBox chk_promedia_costo 
         Caption         =   "Promedia Costo"
         Height          =   210
         Left            =   1380
         TabIndex        =   47
         Top             =   4590
         Width           =   1425
      End
      Begin VB.ComboBox cmb_tipo_clientes 
         Height          =   315
         ItemData        =   "frmmovimientos.frx":13DE
         Left            =   2310
         List            =   "frmmovimientos.frx":13E0
         TabIndex        =   45
         Top             =   4230
         Width           =   3240
      End
      Begin VB.TextBox txt_tipo_cliente 
         Height          =   315
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   44
         Top             =   4215
         Width           =   900
      End
      Begin VB.ComboBox cmb_tipo_proveedores 
         Height          =   315
         ItemData        =   "frmmovimientos.frx":13E2
         Left            =   2310
         List            =   "frmmovimientos.frx":13E4
         TabIndex        =   42
         Top             =   3885
         Width           =   3240
      End
      Begin VB.TextBox txt_tipo_proveedor 
         Height          =   315
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   41
         Top             =   3900
         Width           =   900
      End
      Begin VB.CheckBox chk_aceptar_mas 
         Caption         =   "Aceptar de mas"
         Height          =   210
         Left            =   4080
         TabIndex        =   40
         Top             =   3675
         Width           =   1470
      End
      Begin VB.CheckBox chk_relectura 
         Caption         =   "Relectura"
         Height          =   210
         Left            =   2835
         TabIndex        =   39
         Top             =   3660
         Width           =   1005
      End
      Begin VB.CheckBox chk_requiere_factura 
         Caption         =   "Requiere indicar número de factura"
         Height          =   210
         Left            =   1380
         TabIndex        =   38
         Top             =   1830
         Width           =   3570
      End
      Begin VB.CommandButton cmd_almacenes 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5175
         Picture         =   "frmmovimientos.frx":13E6
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Almacenes para este Movimientos Alt + A"
         Top             =   6360
         Width           =   330
      End
      Begin VB.CheckBox chk_intercompañia 
         Caption         =   "Intercompañia"
         Height          =   210
         Left            =   1380
         TabIndex        =   14
         Top             =   3660
         Width           =   1395
      End
      Begin VB.TextBox txt_clase 
         Height          =   315
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   13
         Top             =   3300
         Width           =   900
      End
      Begin VB.TextBox txt_dependencia 
         Height          =   315
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   11
         Top             =   2970
         Width           =   900
      End
      Begin VB.ComboBox cmb_dependencias 
         Height          =   315
         ItemData        =   "frmmovimientos.frx":1970
         Left            =   2325
         List            =   "frmmovimientos.frx":1972
         TabIndex        =   12
         Top             =   2970
         Width           =   3240
      End
      Begin VB.TextBox txt_documento 
         Height          =   315
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   8
         Top             =   2400
         Width           =   900
      End
      Begin VB.ComboBox cmb_documentos 
         Height          =   315
         ItemData        =   "frmmovimientos.frx":1974
         Left            =   2325
         List            =   "frmmovimientos.frx":1984
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   2400
         Width           =   3240
      End
      Begin VB.CheckBox chk_causa_devolucion 
         Caption         =   "Causa Devolución"
         Height          =   210
         Left            =   1380
         TabIndex        =   10
         Top             =   2745
         Width           =   3570
      End
      Begin VB.CheckBox chk_hacer_referencia 
         Caption         =   "Hacer Referencia a Archivo"
         Height          =   210
         Left            =   1380
         TabIndex        =   5
         Top             =   1245
         Width           =   3570
      End
      Begin VB.ComboBox cmbmovimientos 
         Height          =   315
         ItemData        =   "frmmovimientos.frx":19B3
         Left            =   2310
         List            =   "frmmovimientos.frx":19C6
         TabIndex        =   4
         Top             =   900
         Width           =   3240
      End
      Begin VB.TextBox txt_movimiento 
         Height          =   315
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   900
      End
      Begin VB.TextBox txt_nombre_movimiento 
         Height          =   315
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   2
         Top             =   570
         Width           =   4140
      End
      Begin VB.TextBox txt_afectacion 
         Height          =   315
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   3
         Top             =   900
         Width           =   900
      End
      Begin VB.TextBox txt_referencia 
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1485
         Width           =   1110
      End
      Begin VB.TextBox txt_folio 
         Height          =   315
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   7
         Top             =   2070
         Width           =   1785
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Agrupador:"
         Height          =   195
         Index           =   11
         Left            =   180
         TabIndex        =   57
         Top             =   6015
         Width           =   780
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Reporte:"
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   50
         Top             =   4890
         Width           =   615
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cliente:"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   46
         Top             =   4305
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Proveedor:"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   43
         Top             =   3960
         Width           =   1140
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clase:"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   30
         Top             =   3360
         Width           =   435
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Dependencia:"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   29
         Top             =   3030
         Width           =   1005
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   28
         Top             =   2460
         Width           =   870
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Referencia:"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   26
         Top             =   1530
         Width           =   825
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   18
         Top             =   630
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Afectación:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   17
         Top             =   960
         Width           =   810
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   16
         Top             =   2130
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1455
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   7650
      Width           =   255
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   255
      Top             =   7380
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
            Picture         =   "frmmovimientos.frx":1A16
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":22F0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":2BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":34A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":3D7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":431A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":4BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":54D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":5DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":5EBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":5FCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":60E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":61F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   5865
      TabIndex        =   21
      Top             =   420
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   2085
         TabIndex        =   27
         Top             =   165
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3765
         TabIndex        =   22
         Top             =   165
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ir al primero"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un Registro Atras"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un registro adelante"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ir al ultimo"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de Movimiento:"
         Height          =   195
         Left            =   195
         TabIndex        =   23
         Top             =   210
         Width           =   1845
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   20
      Top             =   270
      Width           =   11475
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3585
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":6304
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":6BDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":74B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":7A54
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":8330
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":8C0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":94E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":95F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":9708
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":981A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":992C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmovimientos.frx":9A3E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmmovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_movimientos As Integer
Dim bitacora As Boolean




Private Sub Check4_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_aceptar_mas_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_ajuste_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_causa_devolucion_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_devolucion_factura_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_hacer_referencia_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_intercompañia_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_promedia_costo_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_reempaque_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_relectura_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_requiere_factura_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_sobrante_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_ultimo_costo_Click()
   var_hubo_cambios = True
End Sub

Private Sub cmb_dependencias_Click()
   txt_dependencia = Obtener_llave(cnn, rs, "TB_MOVIMIENTOS", "VCHA_MOV_NOMBRE", cmb_dependencias, 0, "T")
End Sub

Private Sub cmb_documentos_Click()
   If cmb_documentos = "NOTA DE ENVIO" Then
      txt_documento = "N"
   End If
   If cmb_documentos = "VISTAS" Then
      txt_documento = "V"
   End If
   If cmb_documentos = "DOCUMENTO" Then
      txt_documento = "D"
   End If
   If cmb_documentos = "FACTURA" Then
      txt_documento = "F"
   End If
End Sub

Private Sub cmb_tipo_clientes_Change()
   var_hubo_cambios = True
End Sub

Private Sub cmb_tipo_proveedores_Change()
   var_hubo_cambios = True
End Sub

Private Sub cmbmovimientos_Change()
   If cmbmovimientos = "POSITIVA" Then
      txt_afectacion = "+"
   End If
   If cmbmovimientos = "NEGATIVA" Then
      txt_afectacion = "-"
   End If
   If cmbmovimientos = "SALIDA PARA TRASPASO" Then
      txt_afectacion = "TS"
   End If
   If cmbmovimientos = "ENTRADA PARA TRASPASO" Then
      txt_afectacion = "TE"
   End If
   If cmbmovimientos = "TRASPASOS" Then
      txt_afectacion = "T"
   End If
End Sub

Private Sub cmbmovimientos_Click()
   var_hubo_cambios = True
   If cmbmovimientos = "POSITIVA" Then
      txt_afectacion = "+"
   End If
   If cmbmovimientos = "NEGATIVA" Then
      txt_afectacion = "-"
   End If
   If cmbmovimientos = "SALIDA PARA TRASPASO" Then
      txt_afectacion = "TS"
   End If
   If cmbmovimientos = "ENTRADA PARA TRASPASO" Then
      txt_afectacion = "TE"
   End If
   If cmbmovimientos = "TRASPASOS" Then
      txt_afectacion = "T"
   End If
End Sub

Private Sub cmbmovimientos_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub





Private Sub cmd_almacenes_Click()
   var_movimiento_almacen = txt_movimiento
   frmmov_almacenes.Caption = txt_nombre_movimiento
   Me.Enabled = False
   var_activa_forma_mov_almacenes = Me.Name
   frmmov_almacenes.Show
End Sub

Private Sub cmd_deshacer_Click()
       txt_movimiento.Enabled = False
       Call pro_textos

End Sub

Private Sub cmd_eliminar_Click()
   txt_movimiento.Enabled = False
   var_opcion_seguridad = 2
   var_acepta_seguridad = 1
   If var_global_permiso3 = 1 Then
      var_acepta_seguridad = 2
      If var_global_permiso4 = 1 Then
         frmpasswords2.Show 1
      Else
         frmpasswords.Show 1
      End If
   End If
   If var_acepta_seguridad = 1 Then
      Call pro_elimina_movimientos
      rs.Open "select * from tb_movimientos", cnn, adOpenDynamic, adLockOptimistic
      If rs.BOF Then
         cmd_guardar.Enabled = False
         cmd_deshacer.Enabled = False
         cmd_eliminar.Enabled = False
      Else
         cmd_guardar.Enabled = True
         cmd_deshacer.Enabled = True
         cmd_eliminar.Enabled = True
      End If
      rs.Close
   End If
End Sub

Private Sub cmd_guardar_Click()
Dim var_posible As Boolean
   var_posible = True
   If var_modifica_registro_movimiento = False Then
      rs.Open "select * from tb_movimientos where vcha_mov_movimiento_id = '" + Me.txt_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = False
      End If
      rs.Close
   End If
   If var_posible = True Then
      var_opcion_seguridad = 2
      var_acepta_seguridad = 1
      If var_global_permiso3 = 1 Then
         var_acepta_seguridad = 2
         If var_global_permiso4 = 1 Then
            frmpasswords2.Show 1
         Else
            frmpasswords.Show 1
         End If
      End If
      txt_movimiento.Enabled = False
      If var_acepta_seguridad = 1 Then
         If Trim(txt_folio) = "" Then
            txt_folio = 1
         End If
         Call pro_guardar_movimientos
         rs.Open "select * from tb_movimientos", cnn, adOpenDynamic, adLockOptimistic
         If rs.BOF Then
            cmd_guardar.Enabled = False
            cmd_deshacer.Enabled = False
            cmd_eliminar.Enabled = False
         Else
            cmd_guardar.Enabled = True
            cmd_deshacer.Enabled = True
            cmd_eliminar.Enabled = True
         End If
         rs.Close
      End If
   Else
      MsgBox "Clave de movimiento ya existe", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_movimientos, "LISTADO DE movimientos")
        End If

End Sub

Private Sub cmd_nuevo_Click()
   txt_movimiento.Enabled = True
   Call pro_limpiatextos(Me)
   txt_movimiento.Enabled = True
   txt_movimiento.SetFocus: var_modifica_registro_movimiento = False
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_movimiento = False Then
      var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
      If var_si <> 6 Then
         GoTo salir:
      End If
   Else
      If var_hubo_cambios = True Then
         var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
         If var_si <> 6 Then
            GoTo salir:
         End If
      End If
   End If
   Unload Me
   Exit Sub
salir:
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 71 Then
      cmd_guardar_Click
   End If
   If Shift = 4 And KeyCode = 68 Then
      cmd_deshacer_Click
   End If
   If Shift = 4 And KeyCode = 69 Then
      cmd_eliminar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
   If Shift = 4 And KeyCode = 65 Then
      cmd_almacenes_Click
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 0
   rs.Open "select * from tb_MOVIMIENTOS", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_dependencias.hwnd, rs, 1)
   rs.Close
   var_modifica_registro_movimiento = True
   lv_movimientos.SmallIcons = ImageList1
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_movimientos", cnn, adOpenDynamic, adLockOptimistic
   If rs.BOF Then
      cmd_guardar.Enabled = False
      cmd_deshacer.Enabled = False
      cmd_eliminar.Enabled = False
   Else
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
      cmd_eliminar.Enabled = True
   End If
   rs.Close
   txt_movimiento.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro_movimiento = False
   Call activa_forma(var_activa_forma_movimientos)
End Sub

Private Sub lv_movimientos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_movimientos, ColumnHeader)
End Sub

Private Sub lv_movimientos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txt_movimiento.Enabled = False
    Set lv_movimientos.selectedItem = Item
        pro_textos
        var_modifica_registro_movimiento = True
        txt_movimiento.Enabled = True
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_movimientos.SetFocus
      Call pro_avanzar(Me, lv_movimientos, Button)
      lv_movimientos.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_movimientos.ListItems(1).Selected = True
      lv_movimientos.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_movimientos = lv_movimientos.ListItems.Count
      lv_movimientos.ListItems(numero_items_movimientos).Selected = True
      lv_movimientos.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_movimientos()
   Dim ok As Boolean
   Set TB_MOVIMIENTOS = New TB_MOVIMIENTOS
   Set TB_BITACORA_MOVIMIENTOS = New TB_BITACORA_MOVIMIENTOS
   If txt_movimiento <> "" And txt_nombre_movimiento <> "" Then
      If var_hubo_cambios Then
         rs.Open "select * from tb_movimientos where vcha_mov_movimiento_id = '" + txt_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
         ok = TB_MOVIMIENTOS.Anadir(txt_movimiento, txt_nombre_movimiento, txt_afectacion, chk_hacer_referencia, txt_referencia, txt_folio, Me.chk_requiere_factura, Me.txt_documento, Me.chk_causa_devolucion, Me.txt_dependencia, Me.txt_clase, Me.chk_intercompañia, Me.chk_relectura, Me.chk_aceptar_mas, Me.txt_tipo_proveedor, Me.txt_tipo_cliente, Me.chk_promedia_costo, Me.chk_devolucion_factura, Me.txt_reporte, Me.chk_ajuste, Me.chk_ultimo_costo, Me.chk_reempaque, Me.chk_ajuste_reempaque, Me.chk_sobrante, Me.txt_agrupador)
         If ok Then
            bitacora = True
            If var_modifica_registro_movimiento = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_MOVIMIENTOS.Anadir(txt_movimiento, "VCHA_MOV_NOMBRE", var_operacion_bitacora, "", txt_nombre_movimiento, var_clave_usuario_global, fun_NombrePc, Date)
            Else
               var_operacion_bitacora = "M"
               If rs(0) <> txt_movimiento Then
                  bitacora = TB_BITACORA_MOVIMIENTOS.Anadir(txt_movimiento, "VCHA_MOV_MOVIMIENTO_ID", var_operacion_bitacora, rs(0), txt_movimiento, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(1) <> txt_nombre_movimiento Then
                  bitacora = TB_BITACORA_MOVIMIENTOS.Anadir(txt_movimiento, "VCHA_MOV_NOMBRE", var_operacion_bitacora, rs(1), txt_nombre_movimiento, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If Trim(rs(2)) <> Trim(txt_afectacion) Then
                  bitacora = TB_BITACORA_MOVIMIENTOS.Anadir(txt_movimiento, "VCHA_MOV_AFECTACION", var_operacion_bitacora, rs(2), txt_afectacion, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(3) <> chk_hacer_referencia Then
                  bitacora = TB_BITACORA_MOVIMIENTOS.Anadir(txt_movimiento, "VCHA_MOV_REFERENCIA", var_operacion_bitacora, rs(3), chk_hacer_referencia, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(4) <> txt_referencia Then
                  bitacora = TB_BITACORA_MOVIMIENTOS.Anadir(txt_movimiento, "VCHA_REF_REFERENCIA_ID", var_operacion_bitacora, rs(4), txt_referencia, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(5) <> txt_folio Then
                  bitacora = TB_BITACORA_MOVIMIENTOS.Anadir(txt_movimiento, "VCHA_MOV_FOLIO", var_operacion_bitacora, rs(5), txt_folio, var_clave_usuario_global, fun_NombrePc, Date)
               End If
            End If
            rs.Close
            pro_actualiza_ListView
            txt_movimiento.Enabled = False
            MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
            txt_registros = lv_movimientos.ListItems.Count
            var_modifica_registro_movimiento = True
         Else
            MsgBox "No se puede grabar registro: " + TB_MOVIMIENTOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
   Set TB_MOVIMIENTOS = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_movimientos()
Dim var_llave_usuarios As String

Set TB_MOVIMIENTOS = New TB_MOVIMIENTOS
Set TB_BITACORA_MOVIMIENTOS = New TB_BITACORA_MOVIMIENTOS
On Error GoTo salir:
   ok = True
   If txt_movimiento <> "" And txt_nombre_movimiento <> "" And var_modifica_registro_movimiento = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_MOVIMIENTOS.Eliminar(txt_movimiento)
      Else
         GoTo salir:
      End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_MOVIMIENTOS.Anadir(txt_movimiento, "VCHA_REF_NOMBRE", var_operacion_bitacora, txt_nombre_movimiento, "", var_clave_usuario_global, fun_NombrePc, Date)
         numero_items_movimientos = numero_items_movimientos - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_movimientos.ListItems.Remove (lv_movimientos.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_movimientos.ListItems.Count
         lv_movimientos.selectedItem.Selected = True
         pro_textos
      Else
         MsgBox "No se puede grabar registro: " + TB_MOVIMIENTOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
Set TB_MOVIMIENTOS = Nothing
End Sub


Sub pro_llena_listview1()

Dim list_item As ListItem
   numero_items_movimientos = 0
   rs.Open "select * from TB_movimientos", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
      Set list_item = lv_movimientos.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
      list_item.SubItems(1) = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
      list_item.SubItems(2) = IIf(IsNull(rs!CHAR_MOV_AFECTACION), "", rs!CHAR_MOV_AFECTACION)
      list_item.SubItems(3) = IIf(IsNull(rs!INTE_MOV_REFEREANCIA), 0, rs!INTE_MOV_REFEREANCIA)
      list_item.SubItems(4) = IIf(IsNull(rs!VCHA_MOV_TITULO_ORIGEN), "", rs!VCHA_MOV_TITULO_ORIGEN)
      list_item.SubItems(5) = IIf(IsNull(rs!INTE_MOV_FACTURA), 0, rs!INTE_MOV_FACTURA)
      list_item.SubItems(6) = IIf(IsNull(rs!INTE_MOV_FOLIO), 0, rs!INTE_MOV_FOLIO)
      list_item.SubItems(7) = IIf(IsNull(rs!char_mov_documento), "", rs!char_mov_documento)
      list_item.SubItems(8) = IIf(IsNull(rs!INTE_MOV_CAUSA_DEVOLUCION), 0, rs!INTE_MOV_CAUSA_DEVOLUCION)
      list_item.SubItems(9) = IIf(IsNull(rs!vcha_mov_movimiento_dependencia), "", rs!vcha_mov_movimiento_dependencia)
      list_item.SubItems(10) = IIf(IsNull(rs!vcha_mov_clase), "", rs!vcha_mov_clase)
      list_item.SubItems(11) = IIf(IsNull(rs!inte_mov_intercompañia), 0, rs!inte_mov_intercompañia)
      list_item.SubItems(12) = IIf(IsNull(rs!INTE_MOV_RELECTURA), 0, rs!INTE_MOV_RELECTURA)
      list_item.SubItems(13) = IIf(IsNull(rs!INTE_MOV_ACEPTAR_MAS), 0, rs!INTE_MOV_ACEPTAR_MAS)
      list_item.SubItems(14) = IIf(IsNull(rs!char_mov_tipo_proveedor), "", rs!char_mov_tipo_proveedor)
      list_item.SubItems(15) = IIf(IsNull(rs!CHAR_MOV_TIPO_CLIENTE), "", rs!CHAR_MOV_TIPO_CLIENTE)
      list_item.SubItems(16) = IIf(IsNull(rs!INTE_MOV_PROMEDIA_COSTO), 0, rs!INTE_MOV_PROMEDIA_COSTO)
      list_item.SubItems(17) = IIf(IsNull(rs!INTE_MOV_DEVOLUCION_FACTURA), 0, rs!INTE_MOV_DEVOLUCION_FACTURA)
      list_item.SubItems(18) = IIf(IsNull(rs!vcha_mov_reporte_imprimir), "", rs!vcha_mov_reporte_imprimir)
      
      list_item.SubItems(19) = IIf(IsNull(rs!INTE_MOV_AJUSTE), 0, rs!INTE_MOV_AJUSTE)
      list_item.SubItems(20) = IIf(IsNull(rs!INTE_MOV_ULTIMO_COSTO), 0, rs!INTE_MOV_ULTIMO_COSTO)
      list_item.SubItems(21) = IIf(IsNull(rs!INTE_MOV_REEMPAQUE), 0, rs!INTE_MOV_REEMPAQUE)
      list_item.SubItems(22) = IIf(IsNull(rs!INTE_MOV_AJUSTE_REEMPAQUE), 0, rs!INTE_MOV_AJUSTE_REEMPAQUE)
      list_item.SubItems(23) = IIf(IsNull(rs!INTE_MOV_SOBRANTE), 0, rs!INTE_MOV_SOBRANTE)
      list_item.SubItems(24) = IIf(IsNull(rs!VCHA_MOV_AGRUPADOR_CONCENTRADO), 0, rs!VCHA_MOV_AGRUPADOR_CONCENTRADO)
      rs.MoveNext:
      numero_items_movimientos = numero_items_movimientos + 1
   Wend
   rs.Close
End Sub


Sub pro_textos()
   'On Error GoTo err0:
   Dim var_n As Double
   var_n = lv_movimientos.ListItems.Count
   If var_n > 0 Then
      txt_movimiento = lv_movimientos.selectedItem
      txt_nombre_movimiento = lv_movimientos.selectedItem.SubItems(1)
      txt_afectacion = lv_movimientos.selectedItem.SubItems(2)
      chk_hacer_referencia = lv_movimientos.selectedItem.SubItems(3)
      txt_referencia = lv_movimientos.selectedItem.SubItems(4)
      Me.chk_requiere_factura = lv_movimientos.selectedItem.SubItems(5)
      txt_folio = lv_movimientos.selectedItem.SubItems(6)
      Combo1 = Obtener_llave(cnn, rs, "TB_referencias", "VCHA_ref_referencia_ID", txt_referencia, 1, "T")
      If Trim(txt_afectacion) = "+" Then
         cmbmovimientos = "POSITIVA"
      End If
      If Trim(txt_afectacion) = "-" Then
         cmbmovimientos = "NEGATIVA"
      End If
      If Trim(txt_afectacion) = "TS" Then
         cmbmovimientos = "SALIDA PARA TRASPASO"
      End If
      If Trim(txt_afectacion) = "TE" Then
         cmbmovimientos = "ENTRADA PARA TRASPASO"
      End If
      If Trim(txt_afectacion = "T") Then
         cmbmovimientos = "TRASPASOS"
      End If
      txt_documento = lv_movimientos.selectedItem.SubItems(7)
      If txt_documento = "N" Then
         cmb_documentos = "NOTA DE ENVIO"
      End If
      If txt_documento = "D" Then
         cmb_documentos = "DOCUMENTO"
      End If
      If txt_documento = "N" Then
         cmb_documentos = "FACTURA"
      End If
      If txt_documento = "V" Then
         cmb_documentos = "VISTAS"
      End If
      If Trim(txt_documento) = "" Then
         cmb_documentos = ""
      End If
      chk_causa_devolucion = lv_movimientos.selectedItem.SubItems(8)
      txt_dependencia = lv_movimientos.selectedItem.SubItems(8)
      If Trim(txt_dependencia) = "" Then
         cmb_dependencias = ""
      Else
         rs.Open "select * from tb_movimientos where vcha_mov_movimiento_id = '" + txt_dependencia + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            cmb_dependencias = rs!vcha_mov_nombre
         Else
            cmb_dependencias = ""
         End If
         rs.Close
      End If
      txt_clase = lv_movimientos.selectedItem.SubItems(10)
      chk_intercompañia = lv_movimientos.selectedItem.SubItems(11)
      Me.chk_relectura = lv_movimientos.selectedItem.SubItems(12)
      Me.chk_aceptar_mas = lv_movimientos.selectedItem.SubItems(13)
      Me.txt_tipo_proveedor = lv_movimientos.selectedItem.SubItems(14)
      Me.txt_tipo_cliente = lv_movimientos.selectedItem.SubItems(15)
      Me.chk_promedia_costo = lv_movimientos.selectedItem.SubItems(16)
      Me.chk_devolucion_factura = lv_movimientos.selectedItem.SubItems(17)
      Me.txt_reporte = lv_movimientos.selectedItem.SubItems(18)
      Me.chk_ajuste = lv_movimientos.selectedItem.SubItems(19)
      Me.chk_ultimo_costo = lv_movimientos.selectedItem.SubItems(20)
      Me.chk_reempaque = lv_movimientos.selectedItem.SubItems(21)
      Me.chk_ajuste_reempaque = lv_movimientos.selectedItem.SubItems(22)
      Me.chk_sobrante = lv_movimientos.selectedItem.SubItems(23)
      Me.txt_agrupador = lv_movimientos.selectedItem.SubItems(24)
   End If
   var_numero_renglones = lv_movimientos.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_movimientos.ColumnHeaders(2).Width = 3850
   Else
      lv_movimientos.ColumnHeaders(2).Width = 4099.71
   End If
   var_hubo_cambios = False
   var_modifica_registro_movimiento = True
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_movimiento = False Then
        Set list_item = lv_movimientos.ListItems.Add(, , txt_movimiento)
        list_item.SubItems(1) = txt_nombre_movimiento
        list_item.SubItems(2) = txt_afectacion
        list_item.SubItems(3) = chk_hacer_referencia
        list_item.SubItems(4) = txt_referencia
        list_item.SubItems(5) = Me.chk_requiere_factura
        list_item.SubItems(6) = txt_folio
        list_item.SubItems(7) = txt_documento
        list_item.SubItems(8) = Me.chk_causa_devolucion
        list_item.SubItems(9) = Me.txt_dependencia
        list_item.SubItems(10) = Me.txt_clase
        list_item.SubItems(11) = Me.chk_intercompañia
        list_item.SubItems(12) = Me.chk_relectura
        list_item.SubItems(13) = Me.chk_aceptar_mas
        list_item.SubItems(14) = Me.txt_tipo_proveedor
        list_item.SubItems(15) = Me.txt_tipo_cliente
        list_item.SubItems(16) = Me.chk_promedia_costo
        list_item.SubItems(17) = Me.chk_devolucion_factura
        list_item.SubItems(18) = Me.txt_reporte
        list_item.SubItems(19) = Me.chk_ajuste
        list_item.SubItems(20) = Me.chk_ultimo_costo
        list_item.SubItems(21) = Me.chk_reempaque
        list_item.SubItems(22) = Me.chk_ajuste_reempaque
        list_item.SubItems(23) = Me.chk_sobrante
        list_item.SubItems(24) = Me.txt_agrupador
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_movimientos = numero_items_movimientos + 1
    Else
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).Checked = False
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index) = txt_movimiento
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(1) = txt_nombre_movimiento
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(2) = txt_afectacion
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(3) = chk_hacer_referencia
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(4) = txt_referencia
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(5) = Me.chk_requiere_factura
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(6) = txt_folio
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(7) = txt_documento
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(8) = Me.chk_causa_devolucion
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(9) = Me.txt_dependencia
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(10) = Me.txt_clase
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(11) = Me.chk_intercompañia
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(12) = Me.chk_relectura
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(13) = Me.chk_aceptar_mas
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(14) = Me.txt_tipo_proveedor
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(15) = Me.txt_tipo_cliente
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(16) = Me.chk_promedia_costo
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(17) = Me.chk_devolucion_factura
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(18) = Me.txt_reporte
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(19) = Me.chk_ajuste
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(20) = Me.chk_ultimo_costo
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(21) = Me.chk_reempaque
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(22) = Me.chk_ajuste_reempaque
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(23) = Me.chk_sobrante
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).ListSubItems(24) = Me.txt_agrupador
        lv_movimientos.ListItems.Item(lv_movimientos.selectedItem.Index).Selected = True
    
    End If
    lv_movimientos.SetFocus
End Sub




Private Sub txt_afectacion_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_agrupador_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_movimientos, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_clase_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_clase_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_dependencia_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_dependencia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(txt_dependencia) <> "" Then
         rs.Open "SELECT * FROM TB_MOVIMIENTOS WHERE VCHA_MOV_MOVIMIENTO_ID = '" + txt_dependencia + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            cmb_dependencias = rs!vcha_mov_nombre
            rs.Close
         Else
            txt_dependencia = ""
            rs.Close
            cmb_dependencias.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txt_documento_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_documento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If txt_documento <> "D" And txt_documento <> "F" And txt_documento <> "N" And txt_documento <> "V" Then
         MsgBox "Clave de documento incorrecta", vbOKOnly, "ATENCION"
         txt_documento = ""
         cmb_documentos.SetFocus
      Else
         If txt_documento = "N" Then
            cmb_documentos = "NOTA DE ENVIO"
         End If
         If txt_documento = "D" Then
            cmb_documentos = "DOCUMENTO"
         End If
         If txt_documento = "N" Then
            cmb_documentos = "FACTURA"
         End If
         If txt_documento = "N" Then
            cmb_documentos = "VISTAS"
         End If
      End If
   End If
End Sub

Private Sub txt_movimientos_Change(Index As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_movimientos_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   If KeyAscii = 13 Then
      If Index = 2 Then
         If txt_afectacion <> "-" And txt_afectacion <> "+" And txt_afectacion <> "T" And txt_afectacion <> "TS" And txt_afectacion <> "TE" Then
            MsgBox "Clave de afectación incorrecta", vbOKOnly, "ATENCION"
            txt_afectacion = ""
            cmbmovimientos = ""
            cmbmovimientos.SetFocus
         Else
            If Trim(txt_afectacion) = "+" Then
               cmbmovimientos = "POSITIVA"
            End If
            If Trim(txt_afectacion) = "-" Then
               cmbmovimientos = "NEGATIVA"
            End If
            If Trim(txt_afectacion) = "TS" Then
               cmbmovimientos = "SALIDA PARA TRASPASO"
            End If
            If Trim(txt_afectacion) = "TE" Then
               cmbmovimientos = "ENTRADA PARA TRASPASO"
            End If
            If Trim(txt_afectacion = "T") Then
               cmbmovimientos = "TRASPASOS"
            End If
         End If
      End If
   End If
End Sub

Private Sub txt_folio_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_movimiento_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_movimiento_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_referencia_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_reporte_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_tipo_cliente_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_tipo_proveedor_Change()
   var_hubo_cambios = True
End Sub
