VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmfactura_embarques_nuevo_hogar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturacion Nuevo Hogar"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_embarque_envio 
      Height          =   855
      Left            =   1200
      TabIndex        =   26
      Top             =   585
      Width           =   2115
      Begin VB.TextBox txt_embarque_activo 
         Height          =   315
         Left            =   90
         TabIndex        =   27
         Top             =   420
         Width           =   1920
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         Caption         =   "Embarque"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   28
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.Frame frm_reimpresion 
      Height          =   810
      Left            =   960
      TabIndex        =   62
      Top             =   645
      Width           =   2115
      Begin VB.TextBox txt_factura 
         Height          =   315
         Left            =   75
         TabIndex        =   63
         Top             =   420
         Width           =   1920
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000D&
         Caption         =   "Factura"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   64
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.CommandButton cmd_refacturacion 
      Height          =   315
      Left            =   360
      Picture         =   "frmfactura_embarques_nuevo_hogar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Reimpresion de factura"
      Top             =   60
      Width           =   345
   End
   Begin VB.Frame frm_reimpresion_nueva 
      Height          =   810
      Left            =   645
      TabIndex        =   58
      Top             =   660
      Width           =   2115
      Begin VB.TextBox txt_factura_nueva 
         Height          =   315
         Left            =   75
         TabIndex        =   59
         Top             =   420
         Width           =   1920
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         Caption         =   "Factura"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   60
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Facturas Sugeridas "
      Height          =   1095
      Left            =   6270
      TabIndex        =   50
      Top             =   540
      Width           =   5205
      Begin VB.ComboBox cmb_series 
         Height          =   315
         Left            =   615
         TabIndex        =   8
         Top             =   285
         Width           =   795
      End
      Begin VB.TextBox txt_renglones 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4575
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   615
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txt_a 
         Height          =   315
         Left            =   4575
         TabIndex        =   51
         Top             =   285
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txt_de 
         Height          =   315
         Left            =   2970
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   285
         Width           =   1920
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   135
         TabIndex        =   56
         Top             =   345
         Width           =   405
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Total de Facturas a Imprimir:"
         Height          =   195
         Left            =   1440
         TabIndex        =   55
         Top             =   675
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   195
         Left            =   3510
         TabIndex        =   54
         Top             =   345
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Factura a Imprimir:"
         Height          =   195
         Left            =   1635
         TabIndex        =   53
         Top             =   345
         Width           =   1290
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Movimientos a Facturar"
      Height          =   5550
      Left            =   15
      TabIndex        =   38
      Top             =   1680
      Width           =   11445
      Begin VB.Frame frm_mensaje 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   1740
         Left            =   1545
         TabIndex        =   41
         Top             =   1665
         Width           =   8850
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "Espere un momento por favor."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   675
            Left            =   322
            TabIndex        =   43
            Top             =   945
            Width           =   7920
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "Procesando información"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   600
            Left            =   105
            TabIndex        =   42
            Top             =   255
            Width           =   8355
         End
      End
      Begin VB.TextBox txt_piezas 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   6015
         Width           =   1005
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10485
         TabIndex        =   39
         Top             =   6015
         Width           =   1065
      End
      Begin MSComctlLib.ListView lv_movimientos 
         Height          =   5190
         Left            =   75
         TabIndex        =   10
         Top             =   255
         Width           =   11340
         _ExtentX        =   20003
         _ExtentY        =   9155
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
         NumItems        =   21
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Orden de Surtido"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clase"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Núm. Mov."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Establecimiento"
            Object.Width           =   5115
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cliente"
            Object.Width           =   5115
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Piezas     "
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Importe     "
            Object.Width           =   1799
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Numero de Facturas"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "de"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "a"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "subimporte"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "descuento1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "descuento2"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "descuento3"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "IVA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Importe IVA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Plazo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Agrupador"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Almacen"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Clave Moneda"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Tipo Cambio"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Totales:"
         Height          =   195
         Left            =   8805
         TabIndex        =   44
         Top             =   6075
         Width           =   570
      End
   End
   Begin VB.Frame frm_embarque_correo_ft 
      Height          =   30
      Left            =   3045
      TabIndex        =   35
      Top             =   510
      Width           =   15
      Begin VB.TextBox txt_embarque_correo_ft 
         Height          =   315
         Left            =   90
         TabIndex        =   36
         Top             =   420
         Width           =   1920
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         Caption         =   "Embarque"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   37
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.CommandButton cmd_correo_facturacion_tiendas 
      Caption         =   "FT"
      Height          =   315
      Left            =   360
      TabIndex        =   34
      Top             =   60
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_embarques_cerrados 
      Height          =   315
      Left            =   360
      Picture         =   "frmfactura_embarques_nuevo_hogar.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Embarques cerrados no facturados"
      Top             =   60
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   315
      Left            =   375
      TabIndex        =   32
      Top             =   45
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   195
      Left            =   8325
      TabIndex        =   31
      Top             =   75
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   195
      Left            =   8205
      TabIndex        =   30
      Top             =   75
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   195
      Left            =   8115
      TabIndex        =   29
      Top             =   75
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Frame frm_envio_informacion 
      Height          =   15
      Left            =   2670
      TabIndex        =   23
      Top             =   510
      Width           =   15
      Begin MSComctlLib.ListView lv_envio_informacion 
         Height          =   2235
         Left            =   30
         TabIndex        =   24
         Top             =   435
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   3942
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
            Text            =   "Número"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Agente"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         Caption         =   "Envio de informacion"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   25
         Top             =   45
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmd_nota_envio 
      Height          =   315
      Left            =   720
      Picture         =   "frmfactura_embarques_nuevo_hogar.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Generar Nota de Envio y Correo"
      Top             =   60
      Width           =   345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   8265
      TabIndex        =   21
      Top             =   75
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Frame frm_embarque_relacion 
      Height          =   30
      Left            =   2685
      TabIndex        =   18
      Top             =   495
      Width           =   15
      Begin VB.TextBox txt_embarque_relacion 
         Height          =   315
         Left            =   105
         TabIndex        =   19
         Top             =   405
         Width           =   1920
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         Caption         =   " Embarque para relación"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   20
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.CommandButton cmd_relacion_facturas 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   375
      Picture         =   "frmfactura_embarques_nuevo_hogar.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Relación de Facturas"
      Top             =   60
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Frame frm_correo 
      Height          =   855
      Left            =   735
      TabIndex        =   14
      Top             =   555
      Width           =   2115
      Begin VB.TextBox txt_embarque 
         Height          =   315
         Left            =   90
         TabIndex        =   15
         Top             =   420
         Width           =   1920
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         Caption         =   "Embarque"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   16
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.CommandButton cmd_correo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1065
      Picture         =   "frmfactura_embarques_nuevo_hogar.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Enviar Información"
      Top             =   60
      Width           =   315
   End
   Begin VB.Frame frm_embarques_vivos 
      Height          =   15
      Left            =   2820
      TabIndex        =   2
      Top             =   570
      Width           =   45
      Begin MSComctlLib.ListView lv_embarques 
         Height          =   2235
         Left            =   30
         TabIndex        =   11
         Top             =   435
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   3942
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
            Text            =   "Número"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Agente"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         Caption         =   "Embarques por facturar"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   12
         Top             =   120
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11100
      Picture         =   "frmfactura_embarques_nuevo_hogar.frx":050A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   60
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmfactura_embarques_nuevo_hogar.frx":0B44
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   60
      Width           =   330
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9390
      Top             =   30
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
            Picture         =   "frmfactura_embarques_nuevo_hogar.frx":0C46
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   8805
      Top             =   45
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
            Picture         =   "frmfactura_embarques_nuevo_hogar.frx":0D58
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques_nuevo_hogar.frx":1632
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   8280
      Top             =   60
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
            Picture         =   "frmfactura_embarques_nuevo_hogar.frx":1F0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques_nuevo_hogar.frx":27E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques_nuevo_hogar.frx":30C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques_nuevo_hogar.frx":365C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques_nuevo_hogar.frx":3F38
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques_nuevo_hogar.frx":4812
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques_nuevo_hogar.frx":50EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques_nuevo_hogar.frx":51FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques_nuevo_hogar.frx":5310
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques_nuevo_hogar.frx":5422
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques_nuevo_hogar.frx":5534
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques_nuevo_hogar.frx":5646
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques_nuevo_hogar.frx":57C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   9945
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   10500
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   0
      TabIndex        =   57
      Top             =   360
      Width           =   11460
   End
   Begin VB.Frame Frame1 
      Caption         =   " Embarque "
      Height          =   1095
      Left            =   15
      TabIndex        =   45
      Top             =   540
      Width           =   6210
      Begin VB.TextBox txt_clave_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   885
         TabIndex        =   6
         Top             =   630
         Width           =   1230
      End
      Begin VB.TextBox txt_jaula 
         Height          =   315
         Left            =   4950
         TabIndex        =   5
         Top             =   285
         Width           =   1155
      End
      Begin VB.TextBox txt_fecha 
         Height          =   315
         Left            =   2760
         TabIndex        =   4
         Top             =   285
         Width           =   1620
      End
      Begin VB.TextBox txt_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2115
         TabIndex        =   7
         Top             =   630
         Width           =   3990
      End
      Begin VB.TextBox txt_numero_embarque 
         Height          =   315
         Left            =   885
         TabIndex        =   3
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Jaula:"
         Height          =   195
         Left            =   4515
         TabIndex        =   49
         Top             =   345
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   2235
         TabIndex        =   48
         Top             =   345
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   195
         TabIndex        =   47
         Top             =   675
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   195
         TabIndex        =   46
         Top             =   345
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmfactura_embarques_nuevo_hogar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_ruta As String
Dim var_tabla As ADODB.Connection
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_subtotal As Double
Dim var_descuento As Double
Dim var_total As Double
Dim var_piezas As Double
Dim var_total_piezas As Double
Dim var_total_importe As Double
Dim var_clave_mov As String
Dim var_numero_mov As Double
Dim var_numero_renglones As Double
Dim var_renglones_factura As Double
Dim var_factura_inicio As Double
Dim var_factura_de As Double
Dim var_factura_a As Double
Dim var_descuento_1 As Double
Dim var_descuento_2 As Double
Dim var_descuento_3 As Double
Dim var_imp_descuento_1 As Double
Dim var_imp_descuento_2 As Double
Dim var_imp_total_desc_1 As Double
Dim var_imp_total_desc_2 As Double
Dim var_imp_total_desc_3 As Double
Dim var_subimporte As Double
Dim var_importe_iva As Double
Dim var_imp_neto As Double
Dim var_precio_desc_1 As Double
Dim var_cantidad As Double
Dim var_precio As Double
Dim var_plazo As Integer
Dim var_rfc As String
Dim var_iva As Double
Dim var_agrupador As String
Dim var_moneda As String
Dim var_total_facturas As Double
Dim var_total_de As Double
Dim var_total_a As Double
Dim var_almacen As String
Dim var_estatus_embarque As String
Dim var_tipo_Cambio As Double
Dim var_clave_moneda As String
Dim var_serie As String

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Function fun_copia_archivo(Origen, Destino)
    Copy_File = CopyFile(Origen, Destino, 1)
End Function
Private Sub envio_tb_transito()
    var_cadena = " SELECT dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID, dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID, dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID, dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_SALIDAS.DTIM_SAL_FECHA, dbo.TB_SALIDAS.INTE_SAL_NUMERO, dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID, dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, dbo.TB_SALIDAS.FLOA_SAL_COSTO, dbo.TB_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO,  dbo.TB_SALIDAS.VCHA_SER_SERIE_ID, dbo.TB_SALIDAS.INTE_CAR_NUMERO, dbo.TB_SALIDAS.FLOA_SAL_PROMOCION_1, dbo.TB_SALIDAS.FLOA_SAL_PROMOCION_2, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID , dbo.TB_ENCABEZADO_CARTERA.VCHA_AGE_AGENTE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ENCABEZADO_CARTERA.VCHA_ESB_ESTABLECIMIENTO_ID FROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_SALIDAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND "
    var_cadena = var_cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON  dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_SALIDAS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID WHERE (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_numero_embarque + ") AND (dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')"
    rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
       VAR_AGENTE_TRANSITO = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
       If VAR_AGENTE_TRANSITO = "00100" Then
          
          rsaux11.Open "select * from tb_establecimientos_plantas where vcha_esb_establecimiento_id = '" + rs!vcha_ESB_ESTABLECIMIENTO_id + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
          If Not rsaux11.EOF Then
             txt_planta = IIf(IsNull(rsaux11!VCHA_UOR_UNIDAD_ID), "", rsaux11!VCHA_UOR_UNIDAD_ID)
          Else
             txt_planta = ""
          End If
          rsaux11.Close
          If txt_planta <> "" Then
             rsaux10.Open "select * from tb_plantas where vcha_uor_unidad_id = '" + txt_planta + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
             If Not rsaux10.EOF Then
                var_clave_planta_destino = IIf(IsNull(rsaux10!vcha_pla_planta_id), "", rsaux10!vcha_pla_planta_id)
                VAR_NOMBRE_PLANTA_DESTINO = IIf(IsNull(rsaux10!vcha_pla_descripc), "", rsaux10!vcha_pla_descripc)
             End If
             rsaux10.Close
             rsaux10.Open "select * from tb_plantas where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
             var_clave_planta_origen = IIf(IsNull(rsaux10!vcha_pla_planta_id), "", rsaux10!vcha_pla_planta_id)
             rsaux10.Close
             While Not rs.EOF
                   rsaux11.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_aRTICULO_ID = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                   If Not rsaux11.EOF Then
                      var_descripcion_articulo = IIf(IsNull(rsaux11!vcha_Art_nombre_español), "", rsaux11!vcha_Art_nombre_español)
                   Else
                      var_descripcion_articulo = ""
                   End If
                   rsaux11.Close
                   var_descuento_1 = IIf(IsNull(rs!FLOA_SAL_DESCUENTO_1), 0, rs!FLOA_SAL_DESCUENTO_1)
                   var_descuento_2 = IIf(IsNull(rs!FLOA_SAL_DESCUENTO_2), 0, rs!FLOA_SAL_DESCUENTO_2)
                   var_costo = rs!floa_Sal_precio * (1 - (var_descuento_1 / 100))
                   var_costo = var_costo * (1 - (var_descuento_2 / 100))
                   '1
                   rsaux10.Open "select * from tb_transito where vcha_tra_nota_envio = '" + var_clave_planta_origen + "_" + CStr(rs!inte_Car_numero) + "' and vcha_Art_articulo_id = '" + rs!vcha_Art_Articulo_id + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
                   If rsaux10.EOF Then
                      var_cadena = "insert into tb_transito (vcha_tra_nota_envio, vcha_Art_Articulo_id,                                                              vcha_Art_descripcion,           floa_Tra_cantidad_Enviada,                            floa_tra_costo, vcha_tra_planta_origen, vcha_tra_planta_destino, floa_tra_Cantidad_recibida, vcha_tra_Calidad,VCHA_TRA_STATUS,VCHA_MOV_MOVIMIENTO_ID, VCHA_EMP_EMPRESA_ID, VCHA_SER_SERIE_ID) "
                      var_cadena = var_cadena + "   values  ('" + var_clave_planta_origen + "_" + CStr(rs!inte_Car_numero) + "', '" + rs!vcha_Art_Articulo_id + "','" + var_descripcion_articulo + "', " + CStr(rs!floa_Sal_Cantidad) + ", " + CStr(var_costo) + ",'" + var_clave_planta_origen + "','" + var_clave_planta_destino + "',0,'1','A','EI', '" + var_empresa + "','" + rs!vcha_Ser_Serie_id + "')"
                      rsaux11.Open var_cadena, cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
                   End If
                   rsaux10.Close
                   rs.MoveNext
             Wend
          End If
       End If
    End If
    rs.Close
End Sub



Private Sub cmb_series_Click()
   Dim list_item As ListItem
      txt_agente = ""
      txt_clave_agente = ""
      txt_fecha = ""
      txt_jaula = ""
      txt_de = ""
      txt_a = ""
      txt_renglones = ""
      txt_importe = ""
      txt_piezas = ""
      lv_movimientos.ListItems.Clear
      var_serie = cmb_series
      If Trim(txt_numero_embarque) <> "" Then
         rs.Open "select * from tb_encabezado_embarques where inte_emb_embarque = " + txt_numero_embarque, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_estatus_embarque = IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", rs!CHAR_EMB_ESTATUS)
            If Trim(rs!CHAR_EMB_ESTATUS) = "I" Then
               var_total_facturas = 0
               rsaux2.Open "Select * from tb_agentes where vcha_age_agente_id ='" + rs!VCHA_AGE_AGENTE_ID + "'", cnn, adOpenDynamic, adLockOptimistic
               txt_agente = rsaux2!VCHA_AGE_NOMBRE
               txt_clave_agente = rsaux2!VCHA_AGE_AGENTE_ID
               txt_jaula = rs!inte_jau_jaula_id
               txt_fecha = Date
               var_total_importe = 0
               var_total_piezas = 0
               var_numero_renglones = 0
               rsaux2.Close
               rs.Close
               lv_movimientos.ListItems.Clear
               rs.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
               var_factura_inicio = rs!inte_ser_factura
               var_total_de = var_factura_inicio
               var_total_a = var_factura_inicio
               rs.Close
               rs.Open "SELECT * FROM VW_EMBARQUES_1 WHERE INTE_EMB_EMBARQUE = " + txt_numero_embarque + " and vcha_mov_movimiento_id <> 'AV'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_total_facturas = 0
                  While Not rs.EOF
                     rsaux2.Open "Select * from vw_datos_factura where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        var_plazo = IIf(IsNull(rsaux2!inte_pla_dias), 0, rsaux2!inte_pla_dias)
                        var_agrupador = IIf(IsNull(rsaux2!VCHA_FAG_FAMILIA_AGRUPADOR_ID), "", rsaux2!VCHA_FAG_FAMILIA_AGRUPADOR_ID)
                        var_iva = IIf(IsNull(rsaux2!FLOA_TPE_IVA), 0, rsaux2!FLOA_TPE_IVA)
                     Else
                        var_plazo = 0
                        var_agrupador = ""
                        var_iva = 0
                     End If
                     rsaux2.Close
                     Set list_item = lv_movimientos.ListItems.Add(, , rs!INTE_ORS_ORDEN_SURTIDO)
                     list_item.SubItems(1) = IIf(IsNull(rs!VCHA_MOV_MOVIMIENTO_ID), "", rs!VCHA_MOV_MOVIMIENTO_ID)
                     var_clave_mov = IIf(IsNull(rs!VCHA_MOV_MOVIMIENTO_ID), "", rs!VCHA_MOV_MOVIMIENTO_ID)
                     list_item.SubItems(2) = IIf(IsNull(rs!INTE_SAL_NUMERO), 0, rs!INTE_SAL_NUMERO)
                     var_numero_mov = IIf(IsNull(rs!INTE_SAL_NUMERO), 0, rs!INTE_SAL_NUMERO)
                     list_item.SubItems(3) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
                     list_item.SubItems(4) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                     Call facturas
                     list_item.SubItems(5) = Format(var_piezas, "###,###,##0.00")
                     list_item.SubItems(6) = Format(var_total * (1 + (var_iva / 100)), "###,###,##0.00")
                     list_item.SubItems(8) = var_total_de
                     list_item.SubItems(7) = (var_total_a + 1) - var_total_de
                     var_total_facturas = var_total_facturas + ((var_total_a + 1) - var_total_de)
                     list_item.SubItems(9) = var_total_a
                     var_total_de = var_total_de + ((var_total_a + 1) - var_total_de)
                     list_item.SubItems(10) = var_subimporte
                     list_item.SubItems(11) = var_imp_total_desc_1
                     list_item.SubItems(12) = var_imp_total_desc_2
                     list_item.SubItems(13) = 0
                     list_item.SubItems(14) = var_iva
                     list_item.SubItems(15) = var_total * (var_iva / 100)
                     list_item.SubItems(16) = var_plazo
                     list_item.SubItems(17) = var_agrupador
                     list_item.SubItems(18) = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
                     list_item.SubItems(19) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                     var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                     list_item.SubItems(20) = IIf(IsNull(rs!floa_emo_tipo_cambio), 1, rs!floa_emo_tipo_cambio)
                     var_tipo_Cambio = IIf(IsNull(rs!floa_emo_tipo_cambio), 1, rs!floa_emo_tipo_cambio)
                     var_factura_inicio = var_factura_inicio + (0 + list_item.SubItems(7))
                     var_factura_inicio = var_factura_inicio + Val(txt_renglones)
                     var_total_importe = var_total_importe + (var_total + (var_total * var_iva / 100))
                     var_total_piezas = var_total_piezas + var_piezas
                     rs.MoveNext
                  Wend
                  txt_piezas = Format(var_total_piezas, "###,###,##0.00")
                  txt_importe = Format(var_total_importe, "###,###,##0.00")
                  txt_renglones = lv_movimientos.selectedItem.SubItems(7)
                  txt_de = lv_movimientos.selectedItem.SubItems(8)
                  txt_a = lv_movimientos.selectedItem.SubItems(9)
                  var_almacen = lv_movimientos.selectedItem.SubItems(18)
                  var_clave_movimiento = lv_movimientos.selectedItem.SubItems(1)
                  var_numero_mov = lv_movimientos.selectedItem.SubItems(2)
                  lv_movimientos.SetFocus
               Else
                  MsgBox "El embarque no tiene movimientos asignados", vbOKOnly, "ATENCION"
               End If
            Else
               If Trim(rs!CHAR_EMB_ESTATUS) = "F" Then
                  MsgBox "El embarque ya fue facturado", vbOKOnly, "ATENCION"
               Else
                  MsgBox "El embarque no a sido cerrado aun", vbOKOnly, "ATENCION"
               End If
            End If
            rs.Close
         Else
            rs.Close
            MsgBox "El número de embarque no existe", vbOKOnly, "ATENCION"
            txt_agente = ""
            txt_clave_agente = ""
            lv_movimientos.ListItems.Clear
         End If
      End If
End Sub

Private Sub cmd_aceptar_Click()
Dim si As Integer
   If Trim(txt_numero_embarque) <> "" Then
      If opt_agente = True Then
         si = MsgBox("¿Deseas imprimir la relación de facturas", vbYesNo, "ATENCION")
         If si = 6 Then
         End If
      End If
      If opt_embarque = True Then
      End If
   End If
End Sub

Private Sub cmd_cancelar_Click()
   frm_embarque_relacion.Visible = False
End Sub

Private Sub cmd_correo_Click()
   frm_correo.Visible = True
   txt_embarque.SetFocus
End Sub

Private Sub cmd_correo_facturacion_tiendas_Click()
   frm_embarque_correo_ft.Visible = True
   txt_embarque_correo_ft = ""
   txt_embarque_correo_ft.SetFocus
End Sub

Private Sub cmd_embarques_cerrados_Click()
      rsaux2.Open "select * from VW_EMBARQUES_CERRADOS_NO_FACTURADOS", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux2.EOF Then
         'Me.Enabled = False
         var_activa_forma_detalle_cajas = Me.Name
         frmembarques_cerrados_no_facturados.Show 1
      Else
         MsgBox "No existen embarques cerrados sin facturar", vbOKOnly, "ATENCION"
      End If
      rsaux2.Close
End Sub

Private Sub cmd_imprimir_Click()
   Dim var_cliente_coppel As String
   Dim var_numero_movimientos As Double
   Dim var_numero_factura_inicio As Double
   Dim var_i As Double
   Dim var_j As Double
   Dim var_k As Double
   Dim var_cliente As String
   Dim var_expedicion As String
   Dim var_domicilio As String
   Dim var_ciudad As String
   Dim var_agente As String
   Dim var_linea As String
   Dim var_cantidad As String
   Dim var_descuento_1 As Double
   Dim var_descuento_2 As Double
   Dim var_descuento_3 As Double
   Dim var_precio As Double
   Dim var_precio_str As String
   Dim var_importe As String
   Dim var_subimporte As String
   Dim var_cantidad_letra As String
   Dim var_iva As String
   Dim var_rfc As String
   Dim var_dia As String
   Dim var_mes As String
   Dim var_año As String
   Dim var_porcentaje As Double
   Dim var_Archivo As String
   Dim var_importe_descuento_1 As Double
   Dim var_importe_descuento_2 As Double
   Dim var_importe_descuento_3 As Double
   Dim var_importe_descuento_1_2 As Double
   Dim var_importe_descuento_2_2 As Double
   Dim var_importe_descuento_3_2 As Double
   Dim var_importe_descuento_1_str As String
   Dim var_importe_descuento_2_str As String
   Dim var_importe_descuento_3_str As String
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   Dim var_marca_promocion As String
   Dim var_contador_promociones As Double
   Dim var_cantidad_total As Double
   Dim var_cantidad_total_str As String
   Dim var_factura_envio As Double
   Dim var_consecutivo As Double
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
   Dim var_cliente_sigo As String
   Dim var_pedido_credito As Double
   Dim var_numero_orden_surtido As Double
   Dim var_x As Double
   Dim var_correo_ft As String
   Dim var_si_correo_ft As Integer
   Dim var_leyenda_sorteo As String
   Dim var_si_sorteo As Integer
   Dim var_si_sorteo_pregunta As Integer
   Dim var_importe_pedido_ft As Double
   Dim var_importe_facturado_ft As Double
   cnn.CommandTimeout = 360
   var_leyenda_sorteo = ""
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
   Dim var_posible_cerrar_movimiento As Integer
   var_posible_cerrar_movimiento = 1
   
   Dim dl As Long                                 ' Valor devuelto por la función API
   Dim sAttributes As String                  ' Aributos
   Dim sDriver As String                       ' Nombre del controlador
   Dim sDescription As String                ' Descripción del DSN
   Dim sDsnName As String                  ' Nombre del DSN

   Const ODBC_ADD_SYS_DSN As Long = 4         ' Se creará un DSN de sistema
   Const vbAPINull As Long = 0&                         ' Puntero NULL

   ' se elimina
   Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema
   sDsnName = "DSN=sqlsistema"
   sDriver = "SQL Server"
   dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)

   'se crea
   sDsnName = "sqlsistema"
   sDescription = "sqlsistema"
   sDriver = "SQL Server"
   sAttributes = "DSN=" & sDsnName & Chr(0)
   sAttributes = sAttributes & "Server=" + parametros(0) & Chr$(0)
   sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
   sAttributes = sAttributes & "Database=" + var_bd_movimientos & Chr(0)
   strAttributes = strAttributes & "UID=sa" & Chr$(0)
   strAttributes = strAttributes & "PWD=elia" & Chr$(0)
   dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
   
   If Trim(txt_numero_embarque) <> "" Then
      Me.txt_embarque_correo_ft = txt_numero_embarque
      If var_estatus_embarque = "F" Then
         MsgBox "El embarque ya fue facturado con anterioridad", vbOKOnly, "ATENCION"
      Else
         
         rs.Open "SELECT * FROM TB_PRINCIPAL", cnn, adOpenDynamic, adLockOptimistic
         var_si_sorteo = IIf(IsNull(rs!inte_pri_activar_sorteo), 0, rs!inte_pri_activar_sorteo)
         rs.Close
         If var_si_sorteo = 1 Then
            rs.Open "select * from tb_detalle_embarques where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_numero_embarque, cnn, adOpenDynamic, adLockOptimistic
            If rs!VCHA_MOV_MOVIMIENTO_ID = "FT" Then
               rsaux.Open "select * from tb_sorteo_folios", cnn, adOpenDynamic, adLockOptimistic
               var_si = MsgBox("¿Se va a asignar el boleto del sorteo número " + CStr(rsaux!inte_sor_folio_actual) + "?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  var_si_sorteo_pregunta = 0
               Else
                  var_si_sorteo_pregunta = 1
               End If
               rsaux.Close
            Else
               var_si_sorteo_pregunta = 0
            End If
            rs.Close
         Else
            var_si_sorteo_pregunta = 0
         End If
         If var_si_sorteo_pregunta = 0 Then
         'Sirve para validar que no vaya mercancia con cantidad en NULL
         Cadena = "SELECT     dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID, "
         Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO, dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID,"
         Cadena = Cadena + " dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID , dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD"
         Cadena = Cadena + " FROM         dbo.TB_DETALLE_EMBARQUES INNER JOIN"
         Cadena = Cadena + " dbo.TB_SALIDAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND"
         Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND"
         Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO AND"
         Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID"
         Cadena = Cadena + " WHERE     (dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD IS NULL) AND (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + txt_numero_embarque + ") AND"
         Cadena = Cadena + " (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')"
         rsaux4.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux4.EOF Then
            rsaux4.Close
            MsgBox "El movimiento tiene cantidad en NULL", vbOKOnly, "ATENCION"
         Else
         rsaux4.Close
         si = MsgBox("¿Deseas imprimir las facturas correspondientes al movimiento?", vbYesNo, "ATENCION")
         If si = 6 Then
            si = MsgBox("Confirmar la impresión del movimiento", vbYesNo, "ATENCION")
            If si = 6 Then
               lv_movimientos.ListItems(1).Selected = True
               var_numero_factura_inicio = lv_movimientos.selectedItem.SubItems(8)
               'MsgBox "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'"
               rs.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
               var_factura_inicio = rs!inte_ser_factura
               rs.Close
               If var_numero_factura_inicio <> var_factura_inicio Then
                  MsgBox "La numeración de facturas a cambiado, vuelva a cargar el numero de embarque", vbOKOnly, "ATENCION"
               Else
                  si = 6
                  If si = 6 Then
                     Me.frm_mensaje.Visible = True
                     Me.Refresh
                     fecha_inicio = CStr(Now)
                     Set TB_ENC_EMBARQUE_M = New TB_ENC_EMBARQUE_M
                     'MsgBox "execute factura_embarques '" + var_empresa + "', '" + var_unidad_organizacional + "', " + txt_numero_embarque + ", '', '','" + var_serie + "', 'FA'"
                     rs.Open "execute factura_embarques '" + var_empresa + "', '" + var_unidad_organizacional + "', " + txt_numero_embarque + ", '', '','" + var_serie + "', 'FA'", cnn, adOpenDynamic, adLockOptimistic
                     ok = TB_ENC_EMBARQUE_M.Anadir(var_empresa, var_unidad_organizacional, txt_numero_embarque, "F")
                     Call envio_tb_transito

                     rsaux5.Open "select * from tb_detalle_embarques where inte_emb_embarque = " + Me.txt_numero_embarque + " and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_si_correo_ft = 0
                     var_leyenda_sorteo = ""
                     
                     fecha_fin = CStr(Now)
                     var_estatus_embarque = "F"
                     'aqui se imprime la factura
                     cnn.BeginTrans
                     If rs.State = 1 Then
                        rs.Close
                     End If
                     rs.Open "select isnull(max(inte_tem_consecutivo),0) from tb_temp_factura_embarques", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_consecutivo = rs(0).Value
                     Else
                        var_consecutivo = 0
                     End If
                     rs.Close
                     var_consecutivo = var_consecutivo + 1
                     rs.Open "insert into tb_temp_factura_embarques (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                     cnn.CommitTrans
                     If var_empresa = "28" Then
                        Cadena = "EXEC SP_CREA_TABLA_FACTURAS_VIANNEY_CATALOG " + CStr(var_consecutivo) + ",'" + var_empresa + "'," + txt_numero_embarque
                     Else
                        Cadena = "EXEC SP_CREA_TABLA_FACTURAS_CHIQUIBLANCOS " + CStr(var_consecutivo) + ",'" + var_empresa + "'," + txt_numero_embarque
                     End If
                     rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     Me.frm_mensaje.Visible = False
                     rsaux3.Open "select distinct inte_car_numero, vcha_ser_Serie_id from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic

                     var_cadena = "SELECT dbo.TB_ENCABEZADO_EMBARQUES.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_CLI_CLAVE_ID FROM dbo.TB_ENCABEZADO_EMBARQUES INNER JOIN dbo.TB_DETALLE_EMBARQUES ON   dbo.TB_ENCABEZADO_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE = dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND "
                     var_cadena = var_cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO WHERE  (dbo.TB_ENCABEZADO_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_numero_embarque + ")"
                     'MsgBox var_cadena
                     rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     var_cliente_retencion = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                     rs.Close
                     If Not rsaux3.EOF Then
                        While Not rsaux3.EOF
                              If var_cliente_retencion = "C000008200" Then
                                 MsgBox "Se va a imprimir la factura " + Trim(Str(rsaux3!inte_Car_numero)) + ", prepare la impresora", vbOKOnly, "ATENCION"
                                 Set reporte = appl.OpenReport(App.Path + "\rep_factura_chiquiblancos_reTencion.rpt")
                                 reporte.RecordSelectionFormula = "{TB_TEMP_FACTURA_EMBARQUES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {TB_TEMP_FACTURA_EMBARQUES.inte_car_numero} = " + CStr(rsaux3!inte_Car_numero) + " and {TB_TEMP_FACTURA_EMBARQUES.vcha_ser_serie_id} = '" + Trim(rsaux3!vcha_Ser_Serie_id) + "' and {TB_TEMP_FACTURA_EMBARQUES.vcha_emp_empresa_id} = '" + var_empresa + "'"
                                 frmvistasprevias.cr.ReportSource = reporte
                                 For ntablas = 1 To reporte.Database.Tables.Count
                                     reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                 Next ntablas
                                 frmvistasprevias.cr.ViewReport
                                 frmvistasprevias.Caption = "Reporte de Movimientos"
                                 frmvistasprevias.Show 1
                                 Set reporte = Nothing
                              Else
                                 MsgBox "Se va a imprimir la factura " + Trim(Str(rsaux3!inte_Car_numero)) + ", prepare la impresora", vbOKOnly, "ATENCION"
                                 If var_empresa = "28" Then
                                    Set reporte = appl.OpenReport(App.Path + "\rep_factura_vianney_catalog.rpt")
                                 Else
                                    Set reporte = appl.OpenReport(App.Path + "\rep_factura_chiquiblancos.rpt")
                                 End If
                                 reporte.RecordSelectionFormula = "{TB_TEMP_FACTURA_EMBARQUES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {TB_TEMP_FACTURA_EMBARQUES.inte_car_numero} = " + CStr(rsaux3!inte_Car_numero) + " and {TB_TEMP_FACTURA_EMBARQUES.vcha_ser_serie_id} = '" + Trim(rsaux3!vcha_Ser_Serie_id) + "' and {TB_TEMP_FACTURA_EMBARQUES.vcha_emp_empresa_id} = '" + var_empresa + "'"
                                 frmvistasprevias.cr.ReportSource = reporte
                                 For ntablas = 1 To reporte.Database.Tables.Count
                                     reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                 Next ntablas
                                 frmvistasprevias.cr.ViewReport
                                 frmvistasprevias.Caption = "Reporte de Movimientos"
                                 frmvistasprevias.Show 1
                                 Set reporte = Nothing
                              End If
                              'Set reporte = appl.OpenReport(App.Path + "\rep_factura_chiquiblancos.rpt")
                              'reporte.RecordSelectionFormula = "{TB_TEMP_FACTURA_EMBARQUES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {TB_TEMP_FACTURA_EMBARQUES.inte_car_numero} = " + CStr(rsaux3!inte_car_numero) + " and {TB_TEMP_FACTURA_EMBARQUES.vcha_ser_serie_id} = '" + Trim(rsaux3!vcha_ser_serie_id) + "' and {TB_TEMP_FACTURA_EMBARQUES.vcha_emp_empresa_id} = '" + var_empresa + "'"
                              'For ntablas = 1 To reporte.Database.Tables.Count
                              '    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), "chiquiblancos_sid", parametros(4), parametros(5)
                              'Next ntablas
                              'reporte.PrintOut False
                              'Set reporte = Nothing
                              
                              'Set reporte = appl.OpenReport(App.Path + "\rep_factura_chiquiblancos.rpt")
                              'reporte.RecordSelectionFormula = "{TB_TEMP_FACTURA_EMBARQUES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {TB_TEMP_FACTURA_EMBARQUES.inte_car_numero} = " + CStr(rsaux3!inte_car_numero) + " and {TB_TEMP_FACTURA_EMBARQUES.vcha_ser_serie_id} = '" + Trim(rsaux3!vcha_ser_serie_id) + "' and {TB_TEMP_FACTURA_EMBARQUES.vcha_emp_empresa_id} = '" + var_empresa + "'"
                              'For ntablas = 1 To reporte.Database.Tables.Count
                              '    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), "chiquiblancos_sid", parametros(4), parametros(5)
                              'Next ntablas
                              'reporte.PrintOut False
                              'Set reporte = Nothing
                              rsaux3.MoveNext
                        Wend
                     End If
                     rsaux3.Close
                     rsaux3.Open "delete from TB_TEMP_FACTURA_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                     MsgBox "Se a terminado el proceso de facturación", vbOKOnly, "ATENCION"
                  Else
                     MsgBox "Se a cancelado la impresión de las facturas", vbOKOnly, "ATENCION"
                  End If
               End If
            Else
               MsgBox "Se a cancelado la impresión de las facturas", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Se a cancelado la impresión de las facturas", vbOKOnly, "ATENCION"
         End If
         End If
         Else
         'pregunta si sorteo
         MsgBox "Se a cancelado la impresión de las facturas", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "No se a seleccionado un embarque", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nota_envio_Click()
   Me.frm_embarque_envio.Visible = True
   Me.txt_embarque_activo.SetFocus
End Sub

Private Sub cmd_refacturacion_Click()
   Me.frm_reimpresion.Visible = True
   Me.txt_factura = ""
   Me.txt_factura.SetFocus
End Sub

Private Sub cmd_relacion_facturas_Click()
   txt_embarque_relacion = ""
   frm_embarque_relacion.Visible = True
   txt_embarque_relacion.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
   Set TB_ENTRADAS_INSERTA = New TB_ENTRADAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_LIBERA_APARTADOS = New TB_LIBERA_APARTADOS
   Set TB_SALIDA_VISTAS_I = New TB_SALIDA_VISTAS_I
   Set TB_ARCH_COMPARACION_I = New TB_ARCH_COMPARACION_I
   Set TB_ENC_EMBARQUE_M = New TB_ENC_EMBARQUE_M
   Dim var_referencia_vi As String
   Dim var_contador_renglones As Integer
   Dim var_cadena_cajas As String
   Dim var_posible As Boolean
   Dim var_copia As Boolean
   Dim var_eliminar As Boolean
   Dim var_nombre_archivo As String
   Dim var_numero_folio_anterior As Double
   Dim var_clave_moneda As String
   Dim var_moneda_local As Integer
   Dim var_tipo_Cambio As Double
   Dim var_posible_tipo_cambio As Boolean
   Dim var_clave_movimiento_anterior As String
   Dim var_catalogo_1 As String
   Dim var_catalogo_2 As String
   Dim var_fecha_surtido_catalogo As Date
   Dim var_importe_posible_surtido As Double
   Dim var_importe_surtir As Double
   Dim var_lista_precios_catalogo As String
   Dim var_precio_catalogo_1 As Double
   Dim var_precio_catalogo_2 As Double
   Dim var_importe_disponible As Double
   Dim var_importe_catalogos As Double
   Dim var_mes_catalogo As Integer
   Dim var_año_catalogo As Integer
   
   rs.Open "select * from tb_encabezado_embarques where inte_emb_embarque = " + Str(var_numero_embarque), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_embarque_cerrado = IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", Trim(rs!CHAR_EMB_ESTATUS))
   End If
   rs.Close
   si = MsgBox("¿Esta seguro que desea cerrar el embarque?", vbYesNo, "ATENCION")
   If si = 6 Then
      si = MsgBox("Confirmar el cerrado del embarque", vbOKCancel, "ATENCION")
      If si = 1 Then
                              var_fecha_inicio = CStr(Now)
         var_clave_movimiento_anterior = var_clave_movimiento
         If Trim(var_embarque_cerrado) = "" Then
            rsaux3.Open "select * from vw_embarques_cerrar where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_emb_embarque = " + txt_embarque + " and char_emb_estatus = ''", cnn, adOpenDynamic, adLockOptimistic
            var_tipo_Cambio = 0
            var_posible_tipo_cambio = True
            While Not rsaux3.EOF
               var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
               If var_moneda_local = 0 Then
                  var_tipo_Cambio = IIf(IsNull(rsaux3!mone_tca_importe), 0, rsaux3!mone_tca_importe)
                  If var_tipo_Cambio = 0 Then
                     var_posible_tipo_cambio = False
                  End If
               End If
               rsaux3.MoveNext
            Wend
            If var_posible_tipo_cambio = True Then
               var_numero_folio_anterior = var_numero_folio
               If rsaux3.RecordCount > 0 Then
                  rsaux3.MoveFirst
               End If
               While Not rsaux3.EOF
                  var_clave_movimiento = rsaux3!VCHA_MOV_MOVIMIENTO_ID
                  var_numero_folio = rsaux3!INTE_SAL_NUMERO
                  var_clave_moneda = rsaux3!vcha_mon_moneda_id
                  var_almacen_origen = rsaux3!VCHA_ALM_ALMACEN_ID
                  var_clave_titular = IIf(IsNull(rsaux3!vcha_tit_titular_id), "", rsaux3!vcha_tit_titular_id)
                  var_clave_cliente = IIf(IsNull(rsaux3!vcha_cli_clave_id), "", rsaux3!vcha_cli_clave_id)
                  var_almacen_OS = var_almacen_origen
                  var_estatus_movimiento = rsaux3!char_Emo_estatus
                  var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
                  If var_moneda_local = 0 Then
                     var_tipo_Cambio = IIf(IsNull(rsaux3!mone_tca_importe), 0, rsaux3!mone_tca_importe)
                  Else
                     var_tipo_Cambio = 1
                  End If
                  If var_numero_folio > 0 Then
                     If var_tipo_documento = "F" Then
                        If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
                        Else
                           If var_tipo_Cambio > 0 Then
                              cnn.BeginTrans
                              var_año_catalogo = Year(Date)
                              var_mes_catalogo = Month(Date)
                              rsaux4.Open "select * from TB_PORCENTAJES_FACTURACION_CATALOGOS where INTE_PFC_AÑO = " + CStr(var_año_catalogo) + " and INTE_PFC_MES = " + CStr(var_mes_catalogo) + " and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux4.EOF Then
                                 var_fecha_surtido_catalogo = rsaux4!DTIM_PFC_FECHA_SURTIDO
                                 var_i = 1
                                 While Not rsaux4.EOF
                                     If var_i = 1 Then
                                        var_catalogo_1 = rsaux4!vcha_Art_Articulo_id
                                     Else
                                        var_catalogo_2 = rsaux4!vcha_Art_Articulo_id
                                     End If
                                     var_i = var_i + 1
                                     rsaux4.MoveNext
                                 Wend
                              End If
                              rsaux4.Close
                              If var_fecha_surtido_catalogo <= Date Then
                                 rsaux4.Open "select * from tb_clientes where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux4.EOF Then
                                    var_lista_precios_cataloGos = IIf(IsNull(rsaux4!vcha_LIS_LISTA_iD), "", rsaux4!vcha_LIS_LISTA_iD)
                                 End If
                                 rsaux4.Close
                                 rsaux4.Open "select * from tb_detalle_lista_precios where VCHA_LIS_LISTA_PRECIOS_ID = '" + var_lista_precios_cataloGos + "' and vcha_art_articulo_id = '" + var_catalogo_1 + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux4.EOF Then
                                    var_precio_catalogo_1 = IIf(IsNull(rsaux4!floa_dli_precio), 0, rsaux4!floa_dli_precio)
                                 Else
                                    var_precio_catalogo_1 = 0
                                 End If
                                 rsaux4.Close
                                 rsaux4.Open "select * from tb_detalle_lista_precios where VCHA_LIS_LISTA_PRECIOS_ID = '" + var_lista_precios_cataloGos + "' and vcha_art_articulo_id = '" + var_catalogo_2 + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux4.EOF Then
                                    var_precio_catalogo_2 = IIf(IsNull(rsaux4!floa_dli_precio), 0, rsaux4!floa_dli_precio)
                                 Else
                                    var_precio_catalogo_2 = 0
                                 End If
                                 rsaux4.Close
                                 rsaux4.Open "SELECT * FROM TB_IMPORTES_FACTURACION_CATALOGOS_TITULAR WHERE INTE_FCA_MES = " + CStr(var_mes_catalogo) + " AND INTE_FCA_AÑO = " + CStr(var_año_catalogo) + " AND VCHA_TIT_TITULAR_ID = '" + var_clave_titular + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux4.EOF Then
                                    var_importe_disponible = IIf(IsNull(rsaux4!floa_fca_disponible), 0, rsaux4!floa_fca_disponible)
                                 Else
                                    var_importe_disponible = 0
                                 End If
                                 rsaux4.Close
                              Else
                                 var_catalogo_1 = ""
                                 var_catalogo_2 = ""
                                 var_precio_catalogo_1 = 0
                                 var_precio_catalogo_2 = 0
                                 var_importe_disponible = 0
                              End If
                              rsaux4.Open "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", " + CStr(var_tipo_Cambio) + ",'" + var_catalogo_1 + "','" + var_catalogo_2 + "'," + CStr(var_precio_catalogo_1) + ", " + CStr(var_precio_catalogo_2) + "," + CStr(var_importe_disponible) + ",'" + var_clave_titular + "'," + CStr(var_año_catalogo) + "," + CStr(var_mes_catalogo), cnn, adOpenDynamic, adLockOptimistic
                              
                              x = 0
                              If x = 1 Then
                              Cadena = "SELECT dbo.TB_TEMPORAL_SALIDAS.VCHA_EMP_EMPRESA_ID, dbo.TB_TEMPORAL_SALIDAS.VCHA_UOR_UNIDAD_ID, dbo.TB_TEMPORAL_SALIDAS.VCHA_ALM_ALMACEN_ID, dbo.TB_TEMPORAL_SALIDAS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_TEMPORAL_SALIDAS.INTE_SAL_NUMERO, dbo.TB_TEMPORAL_SALIDAS.VCHA_ART_ARTICULO_ID, dbo.TB_TEMPORAL_SALIDAS.FLOA_SAL_CANTIDAD, dbo.TB_TEMPORAL_SALIDAS.FLOA_SAL_COSTO, dbo.TB_TEMPORAL_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_TEMPORAL_SALIDAS.FLOA_SAL_DESCUENTO, dbo.TB_TEMPORAL_SALIDAS.FLOA_SAL_PROMOCION_1, dbo.TB_TEMPORAL_SALIDAS.FLOA_SAL_PROMOCION_2, char_ped_tipo from dbo.TB_TEMPORAL_SALIDAS LEFT OUTER JOIN dbo.TB_ARTICULOS ON dbo.TB_TEMPORAL_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID where dbo.TB_TEMPORAL_SALIDAS.vcha_emp_empresa_id = '" + var_empresa + "' and dbo.TB_TEMPORAL_SALIDAS.vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
                              Cadena = Cadena + " and dbo.TB_TEMPORAL_SALIDAS.vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero =  " + Str(var_numero_folio) + " ORDER BY dbo.TB_TEMPORAL_SALIDAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_LIN_LINEA_ID"
                              
                              rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                              var_renglones_factura = 0
                              While Not rs.EOF
                                 If var_contador_renglones = 0 Then
                                 End If
                                 ok = TB_LIBERA_APARTADOS.Anadir(var_almacen_OS, rs!vcha_Art_Articulo_id, 0 - rs!floa_Sal_Cantidad)
                                 var_inserta = False
                                 var_suma_cantidad = 0
                                 var_cantidad_llegar = IIf(IsNull(rs!floa_Sal_Cantidad), 0, rs!floa_Sal_Cantidad)
                                 var_cantidad = 0
                                 While var_suma_cantidad < var_cantidad_llegar
                                       rsaux2.Open "select * from tb_existencias where vcha_art_articulo_id =  '" + rs!vcha_Art_Articulo_id + "' and vcha_alm_almacen_id = '" + var_almacen_OS + "'", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux2.EOF Then
                                          If rsaux2!floa_exi_cantidad_2004 >= var_cantidad_llegar Then
                                             var_año = 2004
                                             var_suma_cantidad = var_cantidad_llegar
                                             var_cantidad = var_cantidad_llegar
                                             var_costo = rsaux2!FLOA_EXI_COSTO_2004
                                          Else
                                             var_cantidad_disponible = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                             If var_cantidad_disponible > 0 Then
                                                var_año = 2004
                                                var_suma_cantidad = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                                var_cantidad = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                                var_costo = rsaux2!FLOA_EXI_COSTO_2004
                                             Else
                                                var_año = 2005
                                                var_cantidad = rs!floa_Sal_Cantidad - var_suma_cantidad
                                                var_suma_cantidad = var_cantidad_llegar
                                                var_costo = rsaux2!floa_exi_costo_2005
                                             End If
                                          End If
                                       Else
                                          var_año = 2005
                                          var_suma_cantidad = var_cantidad_llegar
                                          var_cantidad = var_cantidad_llegar
                                          rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id =  '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux4.EOF Then
                                             var_costo = rsaux4!mone_Art_costo_estandar
                                          Else
                                             var_costo = 0
                                          End If
                                          rsaux4.Close
                                       End If
                                       rsaux2.Close
                                       'OTORGAMIENTO DE CATALOGOS
                                       If rs!vcha_Art_Articulo_id = var_catalogo_1 Then
                                          var_importe_catalogos = var_precio_catalogo_1 * var_cantidad
                                          If var_importe_catalogos <= var_importe_disponible Then
                                             rsaux2.Open "update TB_IMPORTES_FACTURACION_CATALOGOS_TITULAR set  floa_fca_disponible = floa_fca_disponible - " + CStr(var_importe_catalogos) + " where vcha_tit_titular_id = '" + var_clave_titular + "' and inte_fca_año = " + CStr(var_año_catalogo) + " and inte_fca_mes = " + CStr(var_mes_catalogo), cnn, adOpenDynamic, adLockOptimistic
                                             var_importe_disponible = var_importe_disponible - var_importe_catalogos
                                             
                                             Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID],[INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2],[char_ped_tipo], [inte_sal_año]) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_Articulo_id + "' , " + CStr(var_cantidad) + ", " + CStr(IIf(IsNull(var_costo), 0, var_costo)) + ", " + CStr(0 * var_tipo_Cambio) + ", 0, 0,  0,'" + rs!char_ped_tipo + "', " + CStr(var_año) + ")"
                                             rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                             If var_renglones_factura = var_contador_renglones Then
                                                var_contador_renglones = 0
                                             End If
                                          Else
                                             rsaux2.Open "update TB_IMPORTES_FACTURACION_CATALOGOS_TITULAR set floa_fca_disponible = floa_fca_disponible - floa_fca_disponible where vcha_tit_titular_id = '" + var_clave_titular + "' and inte_fca_año = " + CStr(var_año_catalogo) + " and inte_fca_mes = " + CStr(var_mes_catalogo), cnn, adOpenDynamic, adLockOptimistic
                                             var_cantidad_gratis = Round(CInt(var_importe_disponible) / var_precio_catalogo_1, 0)
                                             var_cantidad_cobrada = rs!floa_Sal_Cantidad - var_cantidad_gratis
                                             var_importe_disponible = 0
                                             If var_cantidad_gratis > 0 Then
                                                Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID],[INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2],[char_ped_tipo], [inte_sal_año]) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_Articulo_id + "' , " + CStr(var_cantidad_gratis) + ", " + CStr(IIf(IsNull(var_costo), 0, var_costo)) + ", " + CStr(0 * var_tipo_Cambio) + ", 0, 0,  0,'" + rs!char_ped_tipo + "', " + CStr(var_año) + ")"
                                                rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                             End If
                                             If var_renglones_factura = var_contador_renglones Then
                                                var_contador_renglones = 0
                                             End If
                                             Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID],[INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2],[char_ped_tipo], [inte_sal_año]) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_Articulo_id + "' , " + CStr(var_cantidad_cobrada) + ", " + CStr(IIf(IsNull(var_costo), 0, var_costo)) + ", " + CStr(rs!floa_Sal_precio * var_tipo_Cambio) + ", " + CStr(rs!FLOA_SAL_DESCUENTO) + ", " + CStr(rs!floa_sal_promocion_1) + ",  " + CStr(rs!FLOA_SAL_PROMOCION_2) + ",'" + rs!char_ped_tipo + "', " + CStr(var_año) + ")"
                                             rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                             If var_renglones_factura = var_contador_renglones Then
                                                var_contador_renglones = 0
                                             End If
                                          End If
                                       Else
                                          If rs!vcha_Art_Articulo_id = var_catalogo_2 Then
                                             var_importe_catalogos = var_precio_catalogo_2 * var_cantidad
                                             If var_importe_catalogos <= var_importe_disponible Then
                                                rsaux2.Open "update TB_IMPORTES_FACTURACION_CATALOGOS_TITULAR set floa_fca_disponible = floa_fca_disponible - " + CStr(var_importe_catalogos) + " where vcha_tit_titular_id = '" + var_clave_titular + "' and inte_fca_año = " + CStr(var_año_catalogo) + " and inte_fca_mes = " + CStr(var_mes_catalogo), cnn, adOpenDynamic, adLockOptimistic
                                                var_importe_disponible = var_importe_disponible - var_importe_catalogos
                                                Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID],[INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2],[char_ped_tipo], [inte_sal_año]) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_Articulo_id + "' , " + CStr(var_cantidad) + ", " + CStr(IIf(IsNull(var_costo), 0, var_costo)) + ", " + CStr(rs!floa_Sal_precio * var_tipo_Cambio) + ", " + CStr(rs!FLOA_SAL_DESCUENTO) + ", " + CStr(rs!floa_sal_promocion_1) + ",  " + CStr(rs!FLOA_SAL_PROMOCION_2) + ",'" + rs!char_ped_tipo + "', " + CStr(var_año) + ")"
                                                rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                                If var_renglones_factura = var_contador_renglones Then
                                                   var_contador_renglones = 0
                                                End If
                                             Else
                                                rsaux2.Open "update TB_IMPORTES_FACTURACION_CATALOGOS_TITULAR set floa_fca_disponible = floa_fca_disponible - floa_fca_disponible where vcha_tit_titular_id = '" + var_clave_titular + "' and inte_fca_año = " + CStr(var_año_catalogo) + " and inte_fca_mes = " + CStr(var_mes_catalogo), cnn, adOpenDynamic, adLockOptimistic
                                                var_cantidad_gratis = CInt(var_importe_disponible) / var_precio_catalogo_2
                                                var_cantidad_cobrada = rs!floa_Sal_Cantidad - var_cantidad_gratis
                                                var_importe_disponible = 0
                                                If var_cantidad_gratis > 0 Then
                                                   Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID],[INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2],[char_ped_tipo], [inte_sal_año]) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_Articulo_id + "' , " + CStr(var_cantidad_gratis) + ", " + CStr(IIf(IsNull(var_costo), 0, var_costo)) + ", " + CStr(0 * var_tipo_Cambio) + ", 0, 0,  0,'" + rs!char_ped_tipo + "', " + CStr(var_año) + ")"
                                                   rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                                End If
                                                If var_renglones_factura = var_contador_renglones Then
                                                   var_contador_renglones = 0
                                                End If
                                                Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID],[INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2],[char_ped_tipo], [inte_sal_año]) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_Articulo_id + "' , " + CStr(var_cantidad_cobrada) + ", " + CStr(IIf(IsNull(var_costo), 0, var_costo)) + ", " + CStr(rs!floa_Sal_precio * var_tipo_Cambio) + ", " + CStr(rs!FLOA_SAL_DESCUENTO) + ", " + CStr(rs!floa_sal_promocion_1) + ",  " + CStr(rs!FLOA_SAL_PROMOCION_2) + ",'" + rs!char_ped_tipo + "', " + CStr(var_año) + ")"
                                                rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                                If var_renglones_factura = var_contador_renglones Then
                                                   var_contador_renglones = 0
                                                End If
                                             End If
                                          Else
                                             Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID],[INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2],[char_ped_tipo], [inte_sal_año]) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_Articulo_id + "' , " + CStr(var_cantidad) + ", " + CStr(IIf(IsNull(var_costo), 0, var_costo)) + ", " + CStr(rs!floa_Sal_precio * var_tipo_Cambio) + ", " + CStr(rs!FLOA_SAL_DESCUENTO) + ", " + CStr(rs!floa_sal_promocion_1) + ",  " + CStr(rs!FLOA_SAL_PROMOCION_2) + ",'" + rs!char_ped_tipo + "', " + CStr(var_año) + ")"
                                             rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                             If var_renglones_factura = var_contador_renglones Then
                                                var_contador_renglones = 0
                                             End If
                                          End If
                                       End If
                                 Wend
                                 rs.MoveNext
                              Wend
                              rs.Close
                              var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, "", Now, var_tipo_Cambio)
                              var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, "I", Now, var_tipo_Cambio)
                              var_estatus_movimiento = "I"
                              txt_codigo.Enabled = False
                              txt_foco.Enabled = False
                              End If
                              cnn.CommitTrans
                           End If
                        End If
                     End If
                     
                     If var_tipo_documento = "N" Then
                        If var_estatus_movimiento = "I" Then
                           var_nombre_archivo = ""
                           If Len(Trim(Str(var_numero_folio))) = 1 Then
                              var_nombre_archivo = "00000" + Trim(Str(var_numero_folio))
                           End If
                           If Len(Trim(Str(var_numero_folio))) = 2 Then
                              var_nombre_archivo = "0000" + Trim(Str(var_numero_folio))
                           End If
                           If Len(Trim(Str(var_numero_folio))) = 3 Then
                              var_nombre_archivo = "000" + Trim(Str(var_numero_folio))
                           End If
                           If Len(Trim(Str(var_numero_folio))) = 4 Then
                              var_nombre_archivo = "00" + Trim(Str(var_numero_folio))
                           End If
                           If Len(Trim(Str(var_numero_folio))) = 5 Then
                              var_nombre_archivo = "0" + Trim(Str(var_numero_folio))
                           End If
                           If Len(Trim(Str(var_numero_folio))) = 6 Then
                              var_nombre_archivo = Trim(Str(var_numero_folio))
                           End If
                           If Dir(App.Path & "\nota_env.dbf") <> "" Then
                              Set var_tabla = CreateObject("ADODB.connection")
                              var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + App.Path + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
                              rsaux2.Open "delete from nota_env", var_tabla, adOpenDynamic, adLockOptimistic
                              var_eliminar = DeleteFile(App.Path & "\temp_" + Trim(var_nombre_archivo) + ".dbf")
                              var_eliminar = DeleteFile(App.Path & "\" + Trim(var_nombre_archivo) + ".dbf")
                              var_copia = CopyFile(App.Path & "\nota_env.dbf", App.Path & "\temp_" + Trim(var_nombre_archivo) + ".dbf", 1)
                              Cadena = "select * from VW_ORDEN_SURTIDO_MOV where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_emo_numero = " + Str(var_numero_folio)
                              rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                              var_numero_pedido_cliente = 0
                              If Not rs.EOF Then
                                 var_numero_pedido_cliente = IIf(IsNull(rs!INTE_PED_REFERENCIA), 0, rs!INTE_PED_REFERENCIA)
                              Else
                                 var_numero_pedido_cliente = 0
                              End If
                              rs.Close
                              Cadena = "select * from tb_salidas where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio)
                              rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                              While Not rs.EOF
                                    Cadena = "insert into " + App.Path + "\temp_" + Trim(var_nombre_archivo) + ".dbf (cvenota, cvecliente, clapr, canp1, canp2, canp3, canp4, canp5, canp6, prepr, cvepedido, anocosto) values ('" + Trim(Str(var_numero_folio)) + "', '" + var_clave_cliente + "', '" + Mid(Trim(rs!vcha_Art_Articulo_id), 7, 5) + "', " + Trim(CStr(rs!floa_Sal_Cantidad)) + ", 0, 0, 0, 0, 0, " + Trim(CStr(Round(rs!floa_Sal_costo, 4))) + ", " + CStr(var_numero_pedido_cliente) + ", '" + Trim(CStr(rs!INTE_sAL_AÑO)) + "')"
                                    rsaux2.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
                                    rs.MoveNext
                              Wend
                              rs.Close
                              var_tabla.Close
                              var_copia = CopyFile(App.Path & "\temp_" + Trim(var_nombre_archivo) + ".dbf", App.Path & "\" + Trim(var_nombre_archivo) + ".dbf", 1)
                              If Trim(var_correo_electronico) <> "" Then
                                 If MAPISession1.SessionID = 0 Then
                                    MAPISession1.SignOn
                                 End If
                                 MAPIMessages1.SessionID = MAPISession1.SessionID
                                 MAPIMessages1.Compose
                                 MAPIMessages1.RecipDisplayName = var_correo_electronico
                                 MAPIMessages1.RecipAddress = var_correo_electronico
                                 MAPIMessages1.AddressResolveUI = True
                                 MAPIMessages1.ResolveName
                                 MAPIMessages1.MsgSubject = "Nota de envio " + Str(var_numero_folio)
                                 MAPIMessages1.MsgNoteText = "Se adjunta nota de envio número " + Str(var_numero_folio)
                                 MAPIMessages1.AttachmentPathName = App.Path + "\" + Trim(var_nombre_archivo) + ".dbf"
                                 MAPIMessages1.Send True
                                 If MAPISession1.SessionID > 0 Then
                                    MAPISession1.SignOff
                                 End If
                              Else
                                 MsgBox "El cliente no cuenta con una cuenta de correo electronico", vbOKOnly, "ATENCION"
                              End If
                           Else
                              MsgBox "No se encuentra el archivo " + App.Path + "\nota_env.dbf, consulte con el administrador del sistema", vbOKOnly, "ATENCION"
                           End If
                           Set reporte = appl.OpenReport(App.Path + "\rep_notas_envio.rpt")
                           reporte.RecordSelectionFormula = "{VW_orden_surtido_mov.inte_emo_numero} = " + Str(var_numero_folio) + " and {VW_ORDEN_SURTIDO_MOV.FLOA_SAL_CANTIDAD} > 0 and {VW_ORDEN_SURTIDO_MOV.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "'"
                           frmvistasprevias.cr.ReportSource = reporte
                           For ntablas = 1 To reporte.Database.Tables.Count
                               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                           Next ntablas
                           frmvistasprevias.cr.ViewReport
                           frmvistasprevias.Caption = "Reporte de Movimientos"
                           frmvistasprevias.Show 1
                           Set reporte = Nothing
                           rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                        Else
                           If var_tipo_Cambio > 0 Then
                              cnn.BeginTrans
                              Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio)
                              rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                              While Not rs.EOF
                                    ok = TB_LIBERA_APARTADOS.Anadir(var_almacen_OS, rs!vcha_Art_Articulo_id, 0 - rs!floa_Sal_Cantidad)
                                    var_inserta = False
                                    var_suma_cantidad = 0
                                    var_cantidad_llegar = IIf(IsNull(rs!floa_Sal_Cantidad), 0, rs!floa_Sal_Cantidad)
                                    var_cantidad = 0
                                    While var_suma_cantidad < var_cantidad_llegar
                                          rsaux2.Open "select * from tb_existencias where vcha_art_articulo_id =  '" + rs!vcha_Art_Articulo_id + "' and vcha_alm_almacen_id = '" + var_almacen_OS + "'", cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux2.EOF Then
                                             If rsaux2!floa_exi_cantidad_2004 >= var_cantidad_llegar Then
                                                var_año = 2004
                                                var_suma_cantidad = var_cantidad_llegar
                                                var_cantidad = var_cantidad_llegar
                                                var_costo = rsaux2!FLOA_EXI_COSTO_2004
                                             Else
                                                var_cantidad_disponible = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                                If var_cantidad_disponible > 0 Then
                                                   var_año = 2004
                                                   var_suma_cantidad = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                                   var_cantidad = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                                   var_costo = rsaux2!FLOA_EXI_COSTO_2004
                                                Else
                                                   var_año = 2005
                                                   var_cantidad = rs!floa_Sal_Cantidad - var_suma_cantidad
                                                   var_suma_cantidad = var_cantidad_llegar
                                                   var_costo = rsaux2!floa_exi_costo_2005
                                                End If
                                             End If
                                          Else
                                             var_año = 2005
                                             var_suma_cantidad = var_cantidad_llegar
                                             var_cantidad = var_cantidad_llegar
                                             rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id =  '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                             If Not rsaux4.EOF Then
                                                var_costo = rsaux4!mone_Art_costo_estandar
                                             Else
                                                var_costo = 0
                                             End If
                                             rsaux4.Close
                                          End If
                                          rsaux2.Close
                                          Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2], [char_ped_tipo],[inte_sal_año]) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_Articulo_id + "' , " + Str(var_cantidad) + ", " + CStr(var_costo) + ", " + CStr(rs!floa_Sal_precio * var_tipo_Cambio) + ", " + CStr(rs!FLOA_SAL_DESCUENTO) + ", " + CStr(rs!floa_sal_promocion_1) + ",  " + CStr(rs!FLOA_SAL_PROMOCION_2) + ", '" + rs!char_ped_tipo + "', " + CStr(var_año) + ")"
                                          rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                    Wend
                                    rs.MoveNext
                              Wend
                              rs.Close
                              var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, "", Now, var_tipo_Cambio)
                              var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, "I", Now, var_tipo_Cambio)
                              var_estatus_movimiento = "I"
                              var_nombre_archivo = ""
                              cnn.CommitTrans
                              If Len(Trim(Str(var_numero_folio))) = 1 Then
                                 var_nombre_archivo = "00000" + Trim(Str(var_numero_folio))
                              End If
                              If Len(Trim(Str(var_numero_folio))) = 2 Then
                                 var_nombre_archivo = "0000" + Trim(Str(var_numero_folio))
                              End If
                              If Len(Trim(Str(var_numero_folio))) = 3 Then
                                 var_nombre_archivo = "000" + Trim(Str(var_numero_folio))
                              End If
                              If Len(Trim(Str(var_numero_folio))) = 4 Then
                                 var_nombre_archivo = "00" + Trim(Str(var_numero_folio))
                              End If
                              If Len(Trim(Str(var_numero_folio))) = 5 Then
                                 var_nombre_archivo = "0" + Trim(Str(var_numero_folio))
                              End If
                              If Len(Trim(Str(var_numero_folio))) = 6 Then
                                 var_nombre_archivo = Trim(Str(var_numero_folio))
                              End If
                              If Dir(App.Path & "\nota_env.dbf") <> "" Then
                                 Set var_tabla = CreateObject("ADODB.connection")
                                 var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + App.Path + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
                                 rsaux2.Open "delete from nota_env", var_tabla, adOpenDynamic, adLockOptimistic
                                 var_eliminar = DeleteFile(App.Path & "\temp_" + Trim(var_nombre_archivo) + ".dbf")
                                 var_eliminar = DeleteFile(App.Path & "\" + Trim(var_nombre_archivo) + ".dbf")
                                 var_copia = CopyFile(App.Path & "\nota_env.dbf", App.Path & "\temp_" + Trim(var_nombre_archivo) + ".dbf", 1)
                                 Cadena = "select * from VW_ORDEN_SURTIDO_MOV where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_emo_numero = " + Str(var_numero_folio)
                                 rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                 var_numero_pedido_cliente = 0
                                 If Not rs.EOF Then
                                    var_numero_pedido_cliente = IIf(IsNull(rs!INTE_PED_REFERENCIA), 0, rs!INTE_PED_REFERENCIA)
                                 Else
                                    var_numero_pedido_cliente = 0
                                 End If
                                 rs.Close
                                 Cadena = "select * from tb_salidas where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio)
                                 rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                 While Not rs.EOF
                                       Cadena = "insert into " + App.Path + "\temp_" + Trim(var_nombre_archivo) + ".dbf (cvenota, cvecliente, clapr, canp1, canp2, canp3, canp4, canp5, canp6, prepr, cvepedido,ANOCOSTO ) values ('" + Trim(Str(var_numero_folio)) + "', '" + var_clave_cliente + "', '" + Mid(Trim(rs!vcha_Art_Articulo_id), 7, 5) + "', " + Trim(CStr(rs!floa_Sal_Cantidad)) + ", 0, 0, 0, 0, 0, " + Trim(CStr(Round(rs!floa_Sal_costo, 4))) + ", " + CStr(var_numero_pedido_cliente) + ",'" + CStr(rs!INTE_sAL_AÑO) + "')"
                                       rsaux2.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
                                       rs.MoveNext
                                 Wend
                                 var_copia = CopyFile(App.Path & "\temp_" + Trim(var_nombre_archivo) + ".dbf", App.Path & "\" + Trim(var_nombre_archivo) + ".dbf", 1)
                                 var_tabla.Close
                                 If Trim(var_correo_electronico) <> "" Then
                                    If MAPISession1.SessionID = 0 Then
                                       MAPISession1.SignOn
                                    End If
                                    MAPIMessages1.SessionID = MAPISession1.SessionID
                                    MAPIMessages1.Compose
                                    MAPIMessages1.RecipDisplayName = var_correo_electronico
                                    MAPIMessages1.RecipAddress = var_correo_electronico
                                    MAPIMessages1.AddressResolveUI = True
                                    MAPIMessages1.ResolveName
                                    MAPIMessages1.MsgSubject = "Nota de envio " + Str(var_numero_folio)
                                    MAPIMessages1.MsgNoteText = "Se adjunta nota de envio número " + Str(var_numero_folio)
                                    MAPIMessages1.AttachmentPathName = App.Path + "\" + Trim(var_nombre_archivo) + ".dbf"
                                    MAPIMessages1.Send True
                                    If MAPISession1.SessionID > 0 Then
                                       MAPISession1.SignOff
                                    End If
                                 Else
                                    MsgBox "El cliente no cuenta con una cuenta de correo", vbOKOnly, "ATENCION"
                                 End If
                              Else
                                 MsgBox "No se encuentra el archivo " + App.Path + "\nota_env.dbf, consulte con el administrador del sistema", vbOKOnly, "ATENCION"
                              End If
                              Set reporte = appl.OpenReport(App.Path + "\rep_notas_envio.rpt")
                              reporte.RecordSelectionFormula = "{VW_orden_surtido_mov.inte_emo_numero} = " + Str(var_numero_folio) + " and {VW_ORDEN_SURTIDO_MOV.FLOA_SAL_CANTIDAD} > 0 and {VW_ORDEN_SURTIDO_MOV.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "'"
                              frmvistasprevias.cr.ReportSource = reporte
                              For ntablas = 1 To reporte.Database.Tables.Count
                                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                              Next ntablas
                              frmvistasprevias.cr.ViewReport
                              frmvistasprevias.Caption = "Reporte de Movimientos"
                              frmvistasprevias.Show 1
                              Set reporte = Nothing
                              rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                              txt_codigo.Enabled = False
                              rs.Close
                           End If
                        End If
                     End If
                     If var_tipo_documento = "V" Then
                        rs.Open "select * from tb_movimientos where vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_movimiento_dependencia = IIf(IsNull(rs!vcha_mov_movimiento_dependencia), "", rs!vcha_mov_movimiento_dependencia)
                        End If
                        rs.Close
                        If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
                           Set reporte = appl.OpenReport(App.Path + "\rep_salida_vistas.rpt")
                           reporte.RecordSelectionFormula = "{VW_orden_surtido_mov.inte_emo_numero} = " + Str(var_numero_folio) + " and {VW_ORDEN_SURTIDO_MOV.FLOA_SAL_CANTIDAD} > 0 and {VW_ORDEN_SURTIDO_MOV.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "'"
                           frmvistasprevias.cr.ReportSource = reporte
                           For ntablas = 1 To reporte.Database.Tables.Count
                               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                           Next ntablas
                           frmvistasprevias.cr.ViewReport
                           frmvistasprevias.Caption = "Reporte de Movimientos"
                           frmvistasprevias.Show 1
                           Set reporte = Nothing
                           rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                        Else
                           If var_tipo_Cambio > 0 Then
                              cnn.BeginTrans
                              Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio)
                              rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                              var_referencia_vi = ""
                              If Len(Trim(Str(var_numero_folio))) = 1 Then
                                 var_referencia_vi = "000000" + Trim(Str(var_numero_folio))
                              End If
                              If Len(Trim(Str(var_numero_folio))) = 2 Then
                                 var_referencia_vi = "00000" + Trim(Str(var_numero_folio))
                              End If
                              If Len(Trim(Str(var_numero_folio))) = 3 Then
                                 var_referencia_vi = "0000" + Trim(Str(var_numero_folio))
                              End If
                              If Len(Trim(Str(var_numero_folio))) = 4 Then
                                 var_referencia_vi = "000" + Trim(Str(var_numero_folio))
                              End If
                              If Len(Trim(Str(var_numero_folio))) = 5 Then
                                 var_referencia_vi = "00" + Trim(Str(var_numero_folio))
                              End If
                              If Len(Trim(Str(var_numero_folio))) = 6 Then
                                 var_referencia_vi = "0" + Trim(Str(var_numero_folio))
                              End If
                              If Len(Trim(Str(var_numero_folio))) = 7 Then
                                 var_referencia_vi = Trim(Str(var_numero_folio))
                              End If
                              var_referencia_vi = Trim(var_movimiento_dependencia) + "A" + var_referencia_vi
                              If Not rs.EOF Then
                                 'clVCHA_EMP_EMPRESA_ID,clVCHA_UOR_UNIDAD_ID,clvcha_alm_almacen_id,clvcha_mov_movimiento_id,clINTE_ENT_NUMERO As Double, clVCHA_ART_ARTICULO_ID As String, clFLOA_ENT_CANTIDAD As Double, clFLOA_ENT_COSTO As Double, clFLOA_ENT_PRECIO As Double, clFLOA_ENT_DESCUENTO As Double, clVCHA_ENT_ALMACEN_ORIGEN As String
                                 var_inserta = TB_ENTRADAS_INSERTA.Anadir(rs(0).Value, rs(1).Value, var_almacen_Destino, rs(3).Value, rs(4).Value, rs(5).Value, rs(6).Value, rs(7).Value, rs(8).Value, 0, var_almacen_origen)
                                 var_inserta = False
                                 While Not rs.EOF
                                    ok = TB_LIBERA_APARTADOS.Anadir(var_almacen_OS, rs!vcha_Art_Articulo_id, 0 - rs!floa_Sal_Cantidad)
                                    var_inserta = False
                                    Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2], [CHAR_PED_TIPO]) values ('" + rs(0).Value + "', '" + rs(1).Value + "', '" + rs(2).Value + "', '" + rs(3).Value + "', " + CStr(rs(4).Value) + ", '" + rs(5).Value + "' , " + Str(rs(6).Value) + ", " + CStr(rs(7).Value) + ", " + CStr(rs(8).Value * var_tipo_Cambio) + ", " + CStr(rs(9).Value) + ", " + CStr(rs(10).Value) + ",  " + CStr(rs(11).Value) + ", '" + rs!char_ped_tipo + "')"
                                    rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                    'var_inserta = TB_SALIDAS_INSERTA.Anadir(rs(0).Value, rs(1).Value, rs(2).Value, rs(3).Value, rs(4).Value, rs(5).Value, rs(6).Value, rs(7).Value, rs(8).Value * var_tipo_cambio, rs(9).Value)
                                    var_inserta = False
                                    var_inserta = TB_SALIDA_VISTAS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_numero_folio, var_clave_movimiento, rs!vcha_Art_Articulo_id, rs!floa_Sal_Cantidad, rs!floa_Sal_costo, rs!floa_Sal_precio)
                                    var_inserta = False
                                    var_inserta = TB_ARCH_COMPARACION_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_movimiento_dependencia, var_numero_folio, Date, "G", var_clave_agente, rs!vcha_Art_Articulo_id, rs!floa_Sal_costo, rs!floa_Sal_Cantidad, 0, "", var_referencia_vi)
                                    rs.MoveNext
                                 Wend
                              End If
                              rs.Close
                              var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, "", Now, var_tipo_Cambio)
                              var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, "I", Now, var_tipo_Cambio)
                              var_estatus_movimiento = "I"
                              cnn.CommitTrans
                              Set reporte = appl.OpenReport(App.Path + "\rep_salida_vistas.rpt")
                              reporte.RecordSelectionFormula = "{VW_orden_surtido_mov.inte_emo_numero} = " + Str(var_numero_folio) + " and {VW_ORDEN_SURTIDO_MOV.FLOA_SAL_CANTIDAD} > 0 and {VW_ORDEN_SURTIDO_MOV.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "'"
                              frmvistasprevias.cr.ReportSource = reporte
                              For ntablas = 1 To reporte.Database.Tables.Count
                                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                              Next ntablas
                              frmvistasprevias.cr.ViewReport
                              frmvistasprevias.Caption = "Reporte de Movimientos"
                              frmvistasprevias.Show 1
                              Set reporte = Nothing
                              txt_codigo.Enabled = False
                              txt_foco.Enabled = False
                              rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                           End If
                        End If
                     End If
                  End If
                  rsaux3.MoveNext
               Wend
               rsaux3.Close
                              var_fecha_fin = CStr(Now)
                              MsgBox var_fecha_inicio + " " + CStr(var_fecha_fin), vbOKOnly, ""
               MsgBox "Se a cerrado el embarque", vbOKOnly, "ATENCION"
               ok = False
               ok = TB_ENC_EMBARQUE_M.Anadir(var_empresa, var_unidad_organizacional, var_numero_embarque, "I")
               var_estatus_movimiento = "I"
               var_numero_folio = var_numero_folio_anterior
               var_embarque_cerrado = "I"
            Else
               rsaux3.Close
               MsgBox "No es posible cerrar el embarque ya que no se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El embaruqe ya habia sido cerrado con anterioridad", vbOKOnly, "ATENCION"
         End If
         var_clave_movimiento = var_clave_movimiento_anterior
      Else
         MsgBox "El cerrado del embarque a sido cancelado", vbOKOnly, "ATENCION"
      End If
   End If
   frm_sellos.Visible = False
   Exit Sub
   frm_sellos.Visible = False
archivo_ocupado:
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If

   Exit Sub
   frm_sellos.Visible = False
End Sub

Private Sub Command2_Click()
   Dim var_numero_movimientos As Double
   Dim var_numero_factura_inicio As Double
   Dim var_i As Double
   Dim var_j As Double
   Dim var_k As Double
   Dim var_cliente As String
   Dim var_expedicion As String
   Dim var_domicilio As String
   Dim var_ciudad As String
   Dim var_agente As String
   Dim var_linea As String
   Dim var_cantidad As String
   Dim var_descuento_1 As Double
   Dim var_descuento_2 As Double
   Dim var_descuento_3 As Double
   Dim var_precio As Double
   Dim var_precio_str As String
   Dim var_importe As String
   Dim var_subimporte As String
   Dim var_cantidad_letra As String
   Dim var_iva As String
   Dim var_rfc As String
   Dim var_dia As String
   Dim var_mes As String
   Dim var_año As String
   Dim var_porcentaje As Double
   Dim var_Archivo As String
   Dim var_importe_descuento_1 As Double
   Dim var_importe_descuento_2 As Double
   Dim var_importe_descuento_3 As Double
   Dim var_importe_descuento_1_2 As Double
   Dim var_importe_descuento_2_2 As Double
   Dim var_importe_descuento_3_2 As Double
   Dim var_importe_descuento_1_str As String
   Dim var_importe_descuento_2_str As String
   Dim var_importe_descuento_3_str As String
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   Dim var_marca_promocion As String
   Dim var_contador_promociones As Double
   Dim var_cantidad_total As Double
   Dim var_cantidad_total_str As String
   Dim var_factura_envio As Double
   Dim var_consecutivo As Double
   Dim var_x As Double
    
   If Trim(txt_numero_embarque) <> "" Then
      If var_estatus_embarque = "F" Then
         MsgBox "El embarque ya fue facturado con anterioridad", vbOKOnly, "ATENCION"
      Else
         'Sirve para validar que no vaya mercancia con cantidad en NULL
         Cadena = "SELECT     dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID, "
         Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO, dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID,"
         Cadena = Cadena + " dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID , dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD"
         Cadena = Cadena + " FROM         dbo.TB_DETALLE_EMBARQUES INNER JOIN"
         Cadena = Cadena + " dbo.TB_SALIDAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND"
         Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND"
         Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO AND"
         Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID"
         Cadena = Cadena + " WHERE     (dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD IS NULL) AND (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + txt_numero_embarque + ") AND"
         Cadena = Cadena + " (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')"
         rsaux4.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux4.EOF Then
            rsaux4.Close
            MsgBox "El movimiento tiene cantidad en NULL", vbOKOnly, "ATENCION"
         Else
         rsaux4.Close
         si = MsgBox("¿Deseas imprimir las facturas correspondientes al movimiento?", vbYesNo, "ATENCION")
         If si = 6 Then
            si = MsgBox("Confirmar la impresión del movimiento", vbYesNo, "ATENCION")
            If si = 6 Then
               lv_movimientos.ListItems(1).Selected = True
               var_numero_factura_inicio = lv_movimientos.selectedItem.SubItems(8)
               rs.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
               var_factura_inicio = rs!inte_ser_factura
               rs.Close
               If var_numero_factura_inicio <> var_factura_inicio Then
                  MsgBox "La numeración de facturas a cambiado, vuelva a cargar el numero de embarque", vbOKOnly, "ATENCION"
               Else
                  MsgBox "Se va a imprimir la factura " + Trim(Str(var_factura_inicio)), vbOKOnly, "ATENCION"
                  si = MsgBox("¿La impresora esta lista?", vbYesNo, "ATENCION")
                  If si = 6 Then
                     Me.frm_mensaje.Visible = True
                     Me.Refresh
                     fecha_inicio = CStr(Now)
                     Set TB_ENC_EMBARQUE_M = New TB_ENC_EMBARQUE_M
                     rs.Open "execute factura_embarques '" + var_empresa + "', '" + var_unidad_organizacional + "', " + txt_numero_embarque + ", '', '','" + var_serie + "', 'FA'", cnn, adOpenDynamic, adLockOptimistic
                     ok = TB_ENC_EMBARQUE_M.Anadir(var_empresa, var_unidad_organizacional, txt_numero_embarque, "F")
                     fecha_fin = CStr(Now)
                     var_estatus_embarque = "F"
                     'aqui se imprime la factura
                     
                     cnn.BeginTrans
                     rs.Open "select isnull(max(inte_tem_consecutivo),0) from tb_temp_factura_embarques", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_consecutivo = rs(0).Value
                     Else
                        var_consecutivo = 0
                     End If
                     rs.Close
                     var_consecutivo = var_consecutivo + 1
                     rs.Open "insert into tb_temp_factura_embarques (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                     cnn.CommitTrans
                     
                     Cadena = "INSERT INTO [vianney].[dbo].[TB_TEMP_FACTURA_EMBARQUES] ([INTE_TEM_CONSECUTIVO], [VCHA_AGR_AGRUPADOR_ID], [VCHA_SAL_DESCRIPCION_FACTURA], [IMPORTE], [CANTIDAD], [INTE_CAR_PLAZO], [FLOA_CAR_PORCENTAJE_IVA], [FLOA_CAR_PORCENTAJE_IMPUESTO_1], [FLOA_CAR_PORCENTAJE_IMPUESTO_2], [FLOA_CAR_PORCENTAJE_DESCUENTO_1], [FLOA_CAR_PORCENTAJE_DESCUENTO_2], [FLOA_CAR_PORCENTAJE_DESCUENTO_3], [FLOA_CAR_IMPORTE_TOTAL], [FLOA_CAR_IMPORTE_IVA], [FLOA_CAR_IMPORTE_IMPUESTO_1], [FLOA_CAR_IMPORTE_IMPUESTO_2], [FLOA_CAR_IMPORTE_DESCUENTO_1], [FLOA_CAR_IMPORTE_DESCUENTO_2],"
                     Cadena = Cadena + " [FLOA_CAR_IMPORTE_DESCUENTO_3], [FLOA_CAR_SUBIMPORTE], [FLOA_CAR_IMPORTE_NETO], [VCHA_CAR_IMPORTE_LETRA], [VCHA_SER_SERIE_ID], [VCHA_CAR_DOCUMENTO], [INTE_CAR_NUMERO], [DTIM_CAR_FECHA], [VCHA_EMP_EMPRESA_ID], [INTE_EMB_EMBARQUE], [VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_CLI_CLAVE_ID], [VCHA_CLI_NOMBRE], [VCHA_CLI_REPRESENTANTE], [VCHA_AGE_AGENTE_ID], [VCHA_RUT_RUTA_ID], [VCHA_CLI_CURP], [VCHA_CLI_RFC], [VCHA_MON_MONEDA_ID], [VCHA_PLA_PLAZO_ID], [VCHA_TCL_TIPO_CLIENTE_ID], [VCHA_LIS_LISTA_ID], [VCHA_CAN_CANAL_VENTA_ID], [VCHA_TRA_TRANSPORTE_ID], [VCHA_FAG_FAMILIA_AGRUPADOR_ID], [INTE_CLI_AGRUPADOR],"
                     Cadena = Cadena + " [INTE_CLI_ESTATUS], [VCHA_TIT_TITULAR_ID], [CHAR_PRI_PRIORIDAD_ID], [VCHA_CLI_EMAIL], [VCHA_PAI_PAIS_ID], [VCHA_PAI_NOMBRE], [VCHA_EST_ESTADO_ID], [VCHA_EST_NOMBRE], [VCHA_CIU_CIUDAD_ID], [VCHA_CIU_NOMBRE], [VCHA_CLI_COLONIA], [VCHA_CLI_DIRECCION], [VCHA_CLI_CP], [FLOA_GRE_DESCUENTO_1], [FLOA_GRE_DESCUENTO_2], [FLOA_GRE_DESCUENTO_3],  [VCHA_GRE_GRUPO_REAL_ID], [VCHA_GRE_NOMBRE], [VCHA_GAC_GRUPO_ACTUAL_ID], [VCHA_TIT_NOMBRE], [FLOA_TIT_LIMITE_CREDITO], [INTE_PLA_DIAS], [FLOA_GAC_DESCUENTO_1], [FLOA_GAC_DESCUENTO_2], [FLOA_GAC_DESCUENTO_3], [VCHA_CAN_NOMBRE], [INTE_CAN_BUSQUEDA_FACTURA_GRUPO], [FLOA_TPE_IVA], [VCHA_GAC_NOMBRE], "
                     Cadena = Cadena + " [VCHA_MON_NOMBRE], [VCHA_MON_NOMBRE_PLURAL], [VCHA_AGE_NOMBRE], [VCHA_ESB_ESTABLECIMIENTO_ID], [FLOA_CAR_TIPO_CAMBIO], [INTE_ORS_ORDEN_SURTIDO], [INTE_PED_NUMERO], [FLOA_SAL_PROMOCION_1],  [FLOA_SAL_PROMOCION_2], [VCHA_CAR_TIPO_DOCUMENTO], [VCHA_UOR_UNIDAD_ID], [INTE_JAU_JAULA_ID], [VCHA_VEH_VEHICULO_ID], [DTIM_EMB_FECHA_INICIO], [DTIM_EMB_FECHA_FINAL], [CHAR_EMB_ESTATUS], [VCHA_CHO_CHOFER_ID], [FLOA_EMB_CUBICAJE], [CHAR_CAR_TIPO_FACTURACION], [VCHA_CAR_CLASE_ID], [CHAR_CAR_AFECTACION], [VCHA_ALM_ALMACEN_ID], [Expr1], [INTE_EMO_NUMERO], [Expr2], [Expr3], [Expr4], [Expr5], [Expr6], [Expr7], [VCHA_AUD_USUARIO], [VCHA_AUD_MAQUINA], [VCHA_AUD_FECHA], [FLOA_CAR_SALDO], "
                     Cadena = Cadena + " [DTIM_CAR_FECHA_VENCIMIENTO], [DTIM_CAR_FECHA_ENTREGA],[Expr8], [CHAR_CAR_ESTATUS], [DTIM_CAR_FECHA_CANCELACION], [VCHA_CAR_USUARIO_CANCELACION],  [VCHA_CAR_MAQUINA_CANCELACION], [INTE_CLI_ENVIO_FACTURA], [FLOA_SAL_PRECIO_PROMEDIO],  [INTE_CAR_FACTURA_CEROS], [FLOA_CAR_COSTO], [INTE_SAL_CONSECUTIVO_FACTURA])"
                     Cadena = Cadena + " select " + CStr(var_consecutivo) + ", VCHA_AGR_AGRUPADOR_ID, VCHA_SAL_DESCRIPCION_FACTURA, IMPORTE, CANTIDAD, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3, FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_SER_SERIE_ID, VCHA_CAR_DOCUMENTO, INTE_CAR_NUMERO, DTIM_CAR_FECHA, VCHA_EMP_EMPRESA_ID, INTE_EMB_EMBARQUE, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, "
                     Cadena = Cadena + " VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_CLI_REPRESENTANTE, VCHA_AGE_AGENTE_ID, VCHA_RUT_RUTA_ID, VCHA_CLI_CURP, VCHA_CLI_RFC, VCHA_MON_MONEDA_ID, VCHA_PLA_PLAZO_ID, VCHA_TCL_TIPO_CLIENTE_ID, VCHA_LIS_LISTA_ID, VCHA_CAN_CANAL_VENTA_ID, VCHA_TRA_TRANSPORTE_ID, VCHA_FAG_FAMILIA_AGRUPADOR_ID, INTE_CLI_AGRUPADOR, INTE_CLI_ESTATUS, VCHA_TIT_TITULAR_ID, CHAR_PRI_PRIORIDAD_ID, VCHA_CLI_EMAIL, VCHA_PAI_PAIS_ID, VCHA_PAI_NOMBRE, VCHA_EST_ESTADO_ID, VCHA_EST_NOMBRE, VCHA_CIU_CIUDAD_ID, VCHA_CIU_NOMBRE, VCHA_CLI_COLONIA, VCHA_CLI_DIRECCION, VCHA_CLI_CP, FLOA_GRE_DESCUENTO_1, FLOA_GRE_DESCUENTO_2, FLOA_GRE_DESCUENTO_3, VCHA_GRE_GRUPO_REAL_ID, VCHA_GRE_NOMBRE, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_TIT_NOMBRE, FLOA_TIT_LIMITE_CREDITO, INTE_PLA_DIAS, "
                     Cadena = Cadena + " FLOA_GAC_DESCUENTO_1, FLOA_GAC_DESCUENTO_2, FLOA_GAC_DESCUENTO_3, VCHA_CAN_NOMBRE, INTE_CAN_BUSQUEDA_FACTURA_GRUPO, FLOA_TPE_IVA, VCHA_GAC_NOMBRE, VCHA_MON_NOMBRE, VCHA_MON_NOMBRE_PLURAL, VCHA_AGE_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, FLOA_CAR_TIPO_CAMBIO, INTE_ORS_ORDEN_SURTIDO, INTE_PED_NUMERO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, VCHA_CAR_TIPO_DOCUMENTO, VCHA_UOR_UNIDAD_ID, INTE_JAU_JAULA_ID, VCHA_VEH_VEHICULO_ID, DTIM_EMB_FECHA_INICIO, DTIM_EMB_FECHA_FINAL, CHAR_EMB_ESTATUS, VCHA_CHO_CHOFER_ID, FLOA_EMB_CUBICAJE, CHAR_CAR_TIPO_FACTURACION, VCHA_CAR_CLASE_ID, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, Expr1, INTE_EMO_NUMERO, Expr2, Expr3, Expr4, Expr5, Expr6, Expr7, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, "
                     Cadena = Cadena + " FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, Expr8, CHAR_CAR_ESTATUS, DTIM_CAR_FECHA_CANCELACION, VCHA_CAR_USUARIO_CANCELACION, VCHA_CAR_MAQUINA_CANCELACION, INTE_CLI_ENVIO_FACTURA, FLOA_SAL_PRECIO_PROMEDIO, INTE_CAR_FACTURA_CEROS, FLOA_CAR_COSTO, INTE_SAL_CONSECUTIVO_FACTURA from vw_facturas_embarque where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque
                     
                     rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     Me.frm_mensaje.Visible = False
                     rsaux3.Open "select distinct inte_car_numero from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux3.EOF Then
                        var_Archivo = App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat"
                        Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat") For Output As #2
                        While Not rsaux3.EOF
                           If rs.State = 1 Then
                              rs.Close
                           End If
                           If var_empresa <> "03" Then
                              rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY INTE_SAL_CONSECUTIVO_FACTURA", cnn, adOpenDynamic, adLockOptimistic
                           Else
                              rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY vcha_sal_descripcion_factura", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           If Not rs.EOF Then
                             'AQUI EMPIEZA LA FACTURA
                              Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt") For Output As #1
                              'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                              'Print #1, Chr(27) + Chr(67) + Chr(44) + Chr(13)
                              'Print #1, ""
                              Print #1, Chr(15) + Chr(27) + Chr(64)
                              Print #1, Spc(105); Str(rsaux3!inte_Car_numero)
                              Print #1, ""
                              Print #1, ""
                              Print #1, Spc(105); Str(rs!INTE_CAR_PLAZO) + " DIAS DE VENCIMIENTO" + "                  " + Format(rs!dtim_Car_fecha, "Short Date")
                              Print #1, ""
                              'Print #1, Spc(92); Str(rs!inte_car_PLAZO) + " DIAS DE VENCIMIENTO"
                              var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                              For var_j = 1 + Len(Trim(var_cliente)) To 83
                                  var_cliente = var_cliente + " "
                              Next var_j
                              If var_unidad_organizacional = "21" Then
                                 var_cliente = var_cliente + "               MEXICO, D.F."
                              Else
                                 var_cliente = var_cliente + "               AGUASCALIENTES, AGS."
                              End If
                              Print #1, Spc(10); var_cliente
                              var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                              For var_j = 1 + Len(Trim(var_domicilio)) To 83
                                  var_domicilio = var_domicilio + " "
                              Next var_j
                              var_agente = ""
                              var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                              For var_j = 1 + Len(Trim(var_agente)) To 8
                                  var_agente = var_agente + " "
                              Next var_j
                              rsaux4.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux4.EOF Then
                                 var_agente = var_agente + IIf(IsNull(rsaux4!VCHA_AGE_NOMBRE), "", rsaux4!VCHA_AGE_NOMBRE)
                              Else
                                 var_agente = var_agente + ""
                              End If
                              rsaux4.Close
                              var_domicilio = var_domicilio
                              'Print #1, Spc(111); var_agente
                              Print #1, Spc(10); var_domicilio
                              var_ciudad = ""
                              var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                              For var_j = 1 + Len(Trim(var_ciudad)) To 37
                                 var_ciudad = var_ciudad + " "
                              Next var_j
                              
                              var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                              var_ciudad = var_ciudad
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              var_ciudad = var_ciudad + var_rfc
                              
                              For var_j = 1 + Len(Trim(var_estado)) To 46
                                 var_estado = var_estado + " "
                              Next var_j
                              

                              For var_j = 1 + Len(Trim(var_ciudad)) To 14
                                 var_ciudad = var_ciudad + " "
                              Next var_j
                               
                              var_ciudad = var_ciudad + "                                                      " + var_agente
                              
                              VAR_EMBARQUE = "EMB.: " + txt_numero_embarque
                              var_ordern_surtido = x
                              Print #1, Spc(10); var_ciudad
                              var_rfc = "RFC:  " + var_rfc
                              var_rfc = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                              For var_j = 1 + Len(Trim(var_rfc)) To 89
                                 var_rfc = var_rfc + " "
                              Next var_j
                              var_rfc = var_rfc + "               PED.: " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))) + " "
                              var_rfc = var_rfc + " O.S.: " + Trim(Str(IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), 0, rs!INTE_ORS_ORDEN_SURTIDO))) + " " + VAR_EMBARQUE
                              Print #1, var_rfc
                              'Print #1, Spc(10); IIf(IsNull(rs!vcha_esb_establecimiento_id), "", rs!vcha_esb_establecimiento_id)
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
                                    var_i = 0
                                    While Len((var_linea)) < 115
                                          var_linea = var_linea + " "
                                    Wend
                                    var_linea = var_linea + " "
                                    var_linea = var_linea + var_marca_promocion
                                    var_cantidad = Format(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                                    var_cantidad_total = var_cantidad_total + IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                                    If Len(Trim(var_cantidad)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_cantidad)) To 14
                                          var_cantidad = " " + var_cantidad
                                       Next var_j
                                    End If
                                    var_precio = IIf(IsNull(rs!Importe), 0, rs!Importe)
                                    var_descuento_1 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
                                    var_descuento_2 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2)
                                    var_descuento_3 = IIf(IsNull(rs!floa_car_porcentaje_descuento_3), 0, rs!floa_car_porcentaje_descuento_3)
                                    var_porcentaje = (100 - var_descuento_1) / 100
                                    var_precio = var_precio * var_porcentaje
                                    var_importe_descuento_1_2 = (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                    var_importe_descuento_1 = var_importe_descuento_1 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                    var_precio = var_precio * ((100 - var_descuento_2) / 100)
                                    var_importe_descuento_2 = var_importe_descuento_2 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - (var_importe_descuento_1_2 + var_precio))
                                    var_precio = var_precio * ((100 - var_descuento_3) / 100)
                                    var_precio = var_precio / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                                    'var_precio_str = Format(var_precio / IIf(IsNull(rs!cantidad), 0, rs!cantidad), "###,###,##0.00")
                                    var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                    If Len(Trim(var_rfc)) > 0 Then
                                       var_precio_str = Format(IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                                    Else
                                       var_precio_str = Format((IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) * (1 + (rs!floa_car_porcentaje_iva / 100)), "###,###,##0.00")
                                    End If
                                    If Len(Trim(var_precio_str)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_precio_str)) To 14
                                           var_precio_str = " " + var_precio_str
                                       Next var_j
                                    End If
                                    var_linea = var_linea + var_cantidad + var_precio_str
                                    If Len(Trim(var_rfc)) > 0 Then
                                       var_importe = Format((IIf(IsNull(rs!Importe), 0, rs!Importe)), "###,###,##0.00")
                                       If Len(Trim(var_importe)) < 14 Then
                                           For var_j = 1 + Len(Trim(var_importe)) To 14
                                              var_importe = " " + var_importe
                                           Next var_j
                                       End If
                                    Else
                                       var_importe = Format((IIf(IsNull(rs!Importe), 0, rs!Importe) * (1 + (rs!floa_car_porcentaje_iva / 100))), "###,###,##0.00")
                                       If Len(Trim(var_importe)) < 14 Then
                                           For var_j = 1 + Len(Trim(var_importe)) To 14
                                              var_importe = " " + var_importe
                                           Next var_j
                                       End If
                                    End If
                                    var_linea = var_linea + var_importe
                                     
                                    Print #1, var_linea
                                    rs.MoveNext
                                 Else
                                    Print #1, ""
                                 End If
                              Next var_k
                              Print #1, ""
                              'Print #1, ""
                              rs.MoveFirst
                              
                              var_cantidad_total_str = Format(var_cantidad_total, "###,###,##0.00")
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              If Len(Trim(var_rfc)) > 0 Then
                                 var_cantidad_letra = rs!vcha_car_importe_letra
                                 var_importe_descuento_1_str = Format(IIf(IsNull(rs!floa_car_importe_descuento_1), 0, rs!floa_car_importe_descuento_1) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio), "###,###,##0.00")
                                 If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                         var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                    Next var_j
                                 End If
                                 var_importe_descuento_2_str = Format(IIf(IsNull(rs!floa_car_importe_descuento_2), 0, rs!floa_car_importe_descuento_2) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio), "###,###,##0.00")
                                 If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                        var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                    Next var_j
                                 End If
                              Else
                                 var_cantidad_letra = rs!vcha_car_importe_letra
                                 var_importe_descuento_1_str = Format((IIf(IsNull(rs!floa_car_importe_descuento_1), 0, rs!floa_car_importe_descuento_1)) * (1 + (rs!floa_car_porcentaje_iva / 100) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                 If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                         var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                    Next var_j
                                 End If
                                 var_importe_descuento_2_str = Format((IIf(IsNull(rs!floa_car_importe_descuento_2), 0, rs!floa_car_importe_descuento_2)) * (1 + (rs!floa_car_porcentaje_iva / 100) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                 If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                        var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                    Next var_j
                                 End If
                              End If
                              var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                              If Len(Trim(var_linea)) < 145 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 145
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              Print #1, var_linea + var_importe_descuento_1_str
                              If var_empresa = "18" Then
                                 var_linea = ""
                              Else
                                 If Trim(var_cliente_coppel) = "C000002947" Then
                                    var_linea = ""
                                 Else
                                    var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%"
                                 End If
                              End If
                              If Len(Trim(var_linea)) < 145 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 145
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              var_linea = var_linea + var_importe_descuento_2_str
                              Print #1, var_linea
                              If var_contador_promociones > 0 Then
                                 If var_cliente_sigo = "C000001636" Then
                                    Print #1, "Descuento adicional del 2%"
                                 Else
                                    Print #1, var_cadena_promocion_171209
                                 End If
                              Else
                                 If var_cliente_sigo = "C000001636" Then
                                    Print #1, "Descuento adicional del 2%"
                                 Else
                                    Print #1, ""
                                 End If
                              End If
                              
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                              
                              If Len(Trim(var_linea)) < 117 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 117
                                     var_x = var_j Mod 2
                                     If var_x >= 1 Then
                                        var_linea = " " + var_linea
                                     Else
                                        var_linea = var_linea + " "
                                     End If
                                 Next var_j
                              End If
                              
                              If Len(Trim(var_rfc)) = 0 Then
                                 var_subimporte = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                 If Len(Trim(var_subimporte)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                        var_subimporte = " " + var_subimporte
                                    Next var_j
                                 End If
                                 var_iva = "-"
                                 For var_j = 1 + Len(Trim(var_iva)) To 11
                                     var_iva = " " + var_iva
                                  Next var_j
                              Else
                                 var_subimporte = Format(Round(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                 If Len(Trim(var_subimporte)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                        var_subimporte = " " + var_subimporte
                                    Next var_j
                                 End If
                                 var_iva = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                 If Len(Trim(var_iva)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_iva)) To 14
                                        var_iva = " " + var_iva
                                    Next var_j
                                 End If
                              End If
                              
                              If Len(Trim(var_subimporte)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                     var_subimporte = " " + var_subimporte
                                 Next var_j
                              End If
                              var_espacios = 131 - Len(var_cantidad_total_str)
                              var_cantidad_total_str = Trim(var_cantidad_total_str)
                              If Len(Trim(var_cantidad_total_str)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_cantidad_total_str)) To 14
                                     var_cantidad_total_str = " " + var_cantidad_total_str
                                 Next var_j
                              End If
                              var_subimporte = Trim(var_subimporte)
                              If Len(Trim(var_subimporte)) < 24 Then
                                 For var_j = 1 + Len(Trim(var_subimporte)) To 24
                                     var_subimporte = " " + var_subimporte
                                 Next var_j
                              End If
                              
                              var_cantidad_total_str = var_linea + var_cantidad_total_str + "    " + var_subimporte
                              'Print #1, Spc(var_espacios); var_cantidad_total_str; Spc(8); var_subimporte
                              Print #1, var_cantidad_total_str
                              var_linea = "                                                                          ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                        " + var_iva
                              Print #1, var_linea
                              var_dia = Day(rs!dtim_Car_fecha)
                              var_mes = Month(rs!dtim_Car_fecha)
                              var_año = Year(rs!dtim_Car_fecha)
                              
                              var_linea = "                                                             " + CStr(var_dia) + "     " + CStr(var_mes)
                              
                              If Len(var_linea) < 145 Then
                                 For var_j = 1 + Len(var_linea) To 145
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              
                              var_importe = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                              
                              If Len(Trim(var_importe)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_importe)) To 14
                                     var_importe = " " + var_importe
                                 Next var_j
                              End If
                              
                              'var_linea = "                                                                   ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                               " + var_iva
                              'var_linea = "                                                                                                                                                 " + var_importe
                              
                              var_linea = var_linea + var_importe
                              Print #1, var_linea
                              
                              var_linea = var_importe
                              If Len(Trim(var_linea)) < 20 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 20
                                     var_linea = " " + var_linea
                                 Next var_j
                              End If
                              var_linea = var_linea + " " + var_cantidad_letra
                              Print #1, Spc(2); CStr(var_año); var_linea
                              
                              var_linea = ""
                              Print #1, ""
                              Print #1, ""
                              Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                              Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!VCHA_CLI_COLONIA), "", rs!VCHA_CLI_COLONIA))
                              Print #1, Spc(5); Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre))
                              If var_empresa <> "03" Then
                                 Print #1, ""
                                 Print #1, ""
                              Else
                                 Print #1, ""
                                 Print #1, ""
                              End If
                              Print #1, ""
                              Print #1, ""
                              Close #1
                              If Trim(var_empresa) = "02" Then
                                 Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt lpt1"
                              Else
                                 Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt lpt1"
                              End If
                              'AQUI TERMINA LA FACTURA
                           End If
                           rs.Close
                           rsaux3.MoveNext
                        Wend
                        Close #2
                        x = Shell(var_Archivo, vbHide)
                     End If
                     rsaux3.Close
                     'Aqui se termina de imprimir la factura
                     rsaux3.Open "delete from TB_TEMP_FACTURA_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                     MsgBox "Se a terminado el proceso de facturación", vbOKOnly, "ATENCION"
                  Else
                     MsgBox "Se a cancelado la impresión de las facturas", vbOKOnly, "ATENCION"
                  End If
               End If
            Else
               MsgBox "Se a cancelado la impresión de las facturas", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Se a cancelado la impresión de las facturas", vbOKOnly, "ATENCION"
         End If
         End If
      End If
   Else
      MsgBox "No se a seleccionado un embarque", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command3_Click()
            rsaux3.Open "select * from vw_embarques_cerrar where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque_activo + " and char_emb_estatus = 'I'", cnn, adOpenDynamic, adLockOptimistic
            var_tipo_Cambio = 0
            var_posible_tipo_cambio = True
            While Not rsaux3.EOF
               var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
               If var_moneda_local = 0 Then
                  var_tipo_Cambio = IIf(IsNull(rsaux3!mone_tca_importe), 0, rsaux3!mone_tca_importe)
                  If var_tipo_Cambio = 0 Then
                     var_posible_tipo_cambio = False
                  End If
               End If
               rsaux3.MoveNext
            Wend
            If var_posible_tipo_cambio = True Then
               var_numero_folio_anterior = var_numero_folio
               If rsaux3.RecordCount > 0 Then
                  rsaux3.MoveFirst
               End If
               While Not rsaux3.EOF
                  var_clave_movimiento = rsaux3!VCHA_MOV_MOVIMIENTO_ID
                  var_numero_folio = rsaux3!INTE_SAL_NUMERO
                  var_clave_moneda = rsaux3!vcha_mon_moneda_id
                  var_almacen_origen = rsaux3!VCHA_ALM_ALMACEN_ID
                  var_clave_titular = IIf(IsNull(rsaux3!vcha_tit_titular_id), "", rsaux3!vcha_tit_titular_id)
                  var_clave_cliente = IIf(IsNull(rsaux3!vcha_cli_clave_id), "", rsaux3!vcha_cli_clave_id)
                  var_almacen_OS = var_almacen_origen
                  var_estatus_movimiento = rsaux3!char_Emo_estatus
                  var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
                  If var_moneda_local = 0 Then
                     var_tipo_Cambio = IIf(IsNull(rsaux3!mone_tca_importe), 0, rsaux3!mone_tca_importe)
                  Else
                     var_tipo_Cambio = 1
                  End If
                  If var_numero_folio > 0 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_notas_envio.rpt")
                     reporte.RecordSelectionFormula = "{VW_orden_surtido_mov.inte_emo_numero} = " + Str(var_numero_folio) + " and {VW_ORDEN_SURTIDO_MOV.FLOA_SAL_CANTIDAD} > 0 and {VW_ORDEN_SURTIDO_MOV.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "'"
                     frmvistasprevias.cr.ReportSource = reporte
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     frmvistasprevias.cr.ViewReport
                     frmvistasprevias.Caption = "Reporte de Movimientos"
                     frmvistasprevias.Show 1
                     Set reporte = Nothing
                     var_nombre_archivo = ""
                     If Len(Trim(Str(var_numero_folio))) = 1 Then
                        var_nombre_archivo = "00000" + Trim(Str(var_numero_folio))
                     End If
                     If Len(Trim(Str(var_numero_folio))) = 2 Then
                        var_nombre_archivo = "0000" + Trim(Str(var_numero_folio))
                     End If
                     If Len(Trim(Str(var_numero_folio))) = 3 Then
                        var_nombre_archivo = "000" + Trim(Str(var_numero_folio))
                     End If
                     If Len(Trim(Str(var_numero_folio))) = 4 Then
                        var_nombre_archivo = "00" + Trim(Str(var_numero_folio))
                     End If
                     If Len(Trim(Str(var_numero_folio))) = 5 Then
                        var_nombre_archivo = "0" + Trim(Str(var_numero_folio))
                     End If
                     If Len(Trim(Str(var_numero_folio))) = 6 Then
                        var_nombre_archivo = Trim(Str(var_numero_folio))
                     End If
                     If Dir(App.Path & "\nota_env.dbf") <> "" Then
                        Set var_tabla = CreateObject("ADODB.connection")
                        var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + App.Path + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
                        rsaux2.Open "delete from nota_env", var_tabla, adOpenDynamic, adLockOptimistic
                        var_eliminar = DeleteFile(App.Path & "\temp_" + Trim(var_nombre_archivo) + ".dbf")
                        var_eliminar = DeleteFile(App.Path & "\" + Trim(var_nombre_archivo) + ".dbf")
                        var_copia = CopyFile(App.Path & "\nota_env.dbf", App.Path & "\temp_" + Trim(var_nombre_archivo) + ".dbf", 1)
                        Cadena = "select * from VW_ORDEN_SURTIDO_MOV where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_emo_numero = " + Str(var_numero_folio)
                        var_si = MsgBox("          ¿Enviar Correo?", vbYesNo, "ATENCION")
                        If var_si = 6 Then
                           rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                           var_numero_pedido_cliente = 0
                           If Not rs.EOF Then
                              var_numero_pedido_cliente = IIf(IsNull(rs!INTE_PED_REFERENCIA), 0, rs!INTE_PED_REFERENCIA)
                           Else
                              var_numero_pedido_cliente = 0
                           End If
                           rs.Close
                           Cadena = "select * from tb_salidas where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio)
                           rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                           While Not rs.EOF
                                 Cadena = "insert into " + App.Path + "\temp_" + Trim(var_nombre_archivo) + ".dbf (cvenota, cvecliente, clapr, canp1, canp2, canp3, canp4, canp5, canp6, prepr, cvepedido, anocosto) values ('" + Trim(Str(var_numero_folio)) + "', '" + var_clave_cliente + "', '" + Mid(Trim(rs!vcha_Art_Articulo_id), 7, 5) + "', " + Trim(CStr(IIf(IsNull(rs!floa_Sal_Cantidad), 0, rs!floa_Sal_Cantidad))) + ", 0, 0, 0, 0, 0, " + Trim(CStr(Round(rs!floa_Sal_costo, 4))) + ", " + CStr(var_numero_pedido_cliente) + ", '" + Trim(CStr(rs!INTE_sAL_AÑO)) + "')"
                                 rsaux2.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
                                  rs.MoveNext
                           Wend
                           rs.Close
                           var_tabla.Close
                           var_copia = CopyFile(App.Path & "\temp_" + Trim(var_nombre_archivo) + ".dbf", App.Path & "\" + Trim(var_nombre_archivo) + ".dbf", 1)
                           var_correo_electronico = ""
                           rsaux4.Open "select * from tb_clientes where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux4.EOF Then
                              var_correo_electronico = rsaux4!vcha_cli_email
                           End If
                           rsaux4.Close
                           If Trim(var_correo_electronico) <> "" Then
                              If MAPISession1.SessionID = 0 Then
                                 MAPISession1.SignOn
                              End If
                              MAPIMessages1.SessionID = MAPISession1.SessionID
                              MAPIMessages1.Compose
                              MAPIMessages1.RecipDisplayName = var_correo_electronico
                              MAPIMessages1.RecipAddress = var_correo_electronico
                              MAPIMessages1.AddressResolveUI = True
                              MAPIMessages1.ResolveName
                              MAPIMessages1.MsgSubject = "Nota de envio " + Str(var_numero_folio)
                              MAPIMessages1.MsgNoteText = "Se adjunta nota de envio número " + Str(var_numero_folio)
                              MAPIMessages1.AttachmentPathName = App.Path + "\" + Trim(var_nombre_archivo) + ".dbf"
                              MAPIMessages1.Send True
                              If MAPISession1.SessionID > 0 Then
                                 MAPISession1.SignOff
                              End If
                           Else
                              MsgBox "El cliente no cuenta con una cuenta de correo electronico", vbOKOnly, "ATENCION"
                           End If
                        End If
                     Else
                        MsgBox "No se encuentra el archivo " + App.Path + "\nota_env.dbf, consulte con el administrador del sistema", vbOKOnly, "ATENCION"
                     End If
                     rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux3.MoveNext
               Wend
               rsaux3.Close
            Else
               MsgBox "No se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
         End If

End Sub

Private Sub Command4_Click()
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Dim var_numero_movimientos As Double
      Dim var_numero_factura_inicio As Double
      Dim var_i As Double
      Dim var_j As Double
      Dim var_k As Double
      Dim var_cliente As String
      Dim var_expedicion As String
      Dim var_domicilio As String
      Dim var_ciudad As String
      Dim var_agente As String
      Dim var_linea As String
      Dim var_cantidad As String
      Dim var_descuento_1 As Double
      Dim var_descuento_2 As Double
      Dim var_descuento_3 As Double
      Dim var_precio As Double
      Dim var_precio_str As String
      Dim var_importe As String
      Dim var_subimporte As String
      Dim var_cantidad_letra As String
      Dim var_iva As String
      Dim var_rfc As String
      Dim var_dia As String
      Dim var_mes As String
      Dim var_año As String
      Dim var_porcentaje As Double
      Dim var_Archivo As String
      Dim var_importe_descuento_1 As Double
      Dim var_importe_descuento_2 As Double
      Dim var_importe_descuento_3 As Double
      Dim var_importe_descuento_1_2 As Double
      Dim var_importe_descuento_2_2 As Double
      Dim var_importe_descuento_3_2 As Double
      Dim var_importe_descuento_1_str As String
      Dim var_importe_descuento_2_str As String
      Dim var_importe_descuento_3_str As String
      Dim var_promocion_1 As Double
      Dim var_promocion_2 As Double
      Dim var_marca_promocion As String
      Dim var_contador_promociones As Double
      Dim var_cantidad_total As Double
      Dim var_cantidad_total_str As String
      Dim var_factura_envio As Double
      Dim var_x As Double
      Dim var_correo_electronico As String
      Dim var_numero_archivo As Double
      Dim var_nombre_archivo As String
      If Dir(var_ruta & "\tfactura_enivar.dbf") <> "" Then
         If IsNumeric(txt_embarque) Then
            If Trim(var_ruta) <> "" Then
               If var_tabla.State = 1 Then
                  var_tabla.Close
               End If
               var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
               rs.Open "select distinct vcha_cli_clave_id,VCHA_CLI_EMAIL, VCHA_CLI_NOMBRE from vw_facturas_embarque where vcha_emp_empresa_id = '" + var_empresa + "' and inte_cli_envio_factura = 1 and inte_emb_embarque = " + txt_embarque, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  While Not rs.EOF
                        'On Error GoTo salir:
                        var_cliente = rs!vcha_cli_clave_id
                        var_correo_electronico = IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
                        If Trim(var_correo_electronico) <> "" Then
                           rsaux.Open "select max(inte_car_numero) from vw_facturas_embarque where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + var_cliente + "' and inte_emb_embarque = " + txt_embarque, cnn, adOpenDynamic, adLockOptimistic
                           var_numero_archivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                           rsaux.Close
                           If Len(Trim(Str(var_numero_archivo))) = 1 Then
                              var_nombre_archivo = "00000" + Trim(Str(var_numero_archivo))
                           End If
                           If Len(Trim(Str(var_numero_archivo))) = 2 Then
                              var_nombre_archivo = "0000" + Trim(Str(var_numero_archivo))
                           End If
                           If Len(Trim(Str(var_numero_archivo))) = 3 Then
                               var_nombre_archivo = "000" + Trim(Str(var_numero_archivo))
                           End If
                           If Len(Trim(Str(var_numero_archivo))) = 4 Then
                              var_nombre_archivo = "00" + Trim(Str(var_numero_archivo))
                           End If
                           If Len(Trim(Str(var_numero_archivo))) = 5 Then
                              var_nombre_archivo = "0" + Trim(Str(var_numero_archivo))
                           End If
                           If Len(Trim(Str(var_numero_archivo))) = 6 Then
                              var_nombre_archivo = Trim(Str(var_numero_archivo))
                           End If
                           If Dir(var_ruta & "\t_" + var_nombre_archivo + ".dbf") <> "" Then
                              Kill var_ruta & "\t_" + Trim(var_nombre_archivo) + ".dbf"
                           End If
                           If Dir(var_ruta & "\" + var_nombre_archivo + ".dbf") <> "" Then
                              Kill var_ruta & "\" + Trim(var_nombre_archivo) + ".dbf"
                           End If
                           var_copia = CopyFile(var_ruta & "\tfactura_enivar.dbf", var_ruta + "\t_" + var_nombre_archivo + ".dbf", 1)
                           rsaux2.Open "select * from VW_FACTURAS_ENVIO_ARCHIVO where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + var_cliente + "' and inte_emb_embarque = " + txt_embarque + " order by inte_car_numero", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux2.EOF Then
                              While Not rsaux2.EOF
                                    var_precio = IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio)
                                    var_iva = IIf(IsNull(rsaux2!floa_car_porcentaje_iva), 0, rsaux2!floa_car_porcentaje_iva)
                                    var_descuento_1 = IIf(IsNull(rsaux2!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rsaux2!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
                                    var_descuento_2 = IIf(IsNull(rsaux2!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rsaux2!FLOA_CAR_PORCENTAJE_DESCUENTO_2)
                                    var_descuento_3 = IIf(IsNull(rsaux2!floa_car_porcentaje_descuento_3), 0, rsaux2!floa_car_porcentaje_descuento_3)
                                    var_porcentaje = (100 - var_descuento_1) / 100
                                    var_precio = var_precio * var_porcentaje
                                    var_importe_descuento_1_2 = (IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio) - var_precio)
                                    var_importe_descuento_1 = var_importe_descuento_1 + (IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio) - var_precio)
                                    var_precio = var_precio * ((100 - var_descuento_2) / 100)
                                    var_importe_descuento_2 = var_importe_descuento_2 + (IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio) - (var_importe_descuento_1_2 + var_precio))
                                    var_precio = var_precio * ((100 - var_descuento_3) / 100)
                                    var_precio = var_precio / IIf(IsNull(rsaux2!floa_car_tipo_cambio), 1, rsaux2!floa_car_tipo_cambio)
                                    var_precio = var_precio * (1 + (var_iva / 100))
                                    rsaux4.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + IIf(IsNull(rsaux2!vcha_Art_Articulo_id), "", rsaux2!vcha_Art_Articulo_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                                    rsaux.Open "insert into " + var_ruta + "\t_" + var_nombre_archivo + " (folio, cvetienda, codigo, cant1,cant2, cant3, cant4, cant5, cant6, costo) values (" + CStr(IIf(IsNull(rsaux2!inte_Car_numero), 0, rsaux2!inte_Car_numero)) + ", '" + var_cliente + "', '" + IIf(IsNull(rsaux4!VCHA_aRT_CODIGO_EXTERNO), "", rsaux4!VCHA_aRT_CODIGO_EXTERNO) + "', " + CStr(IIf(IsNull(rsaux2!floa_Sal_Cantidad), 0, rsaux2!floa_Sal_Cantidad)) + ",0,0,0,0,0," + CStr(var_precio) + ")", var_tabla, adOpenDynamic, adLockOptimistic
                                    rsaux4.Close
                                    rsaux2.MoveNext
                              Wend
                           End If
                           rsaux2.Close
                           var_copia = CopyFile(var_ruta + "\t_" + var_nombre_archivo + ".dbf", var_ruta + "\" + var_nombre_archivo + ".dbf", 1)
                           If MAPISession1.SessionID = 0 Then
                              MAPISession1.SignOn
                           End If
                           MAPIMessages1.SessionID = MAPISession1.SessionID
                           MAPIMessages1.Compose
                           MAPIMessages1.RecipDisplayName = var_correo_electronico
                           MAPIMessages1.RecipAddress = var_correo_electronico
                           MAPIMessages1.AddressResolveUI = True
                           MAPIMessages1.ResolveName
                           MAPIMessages1.MsgSubject = "Archivo de articulos enviados"
                           MAPIMessages1.MsgNoteText = "Se adjunta archivos de articulos enviados"
                           MAPIMessages1.AttachmentPathName = var_ruta + "\" + var_nombre_archivo + ".dbf"
                           MAPIMessages1.Send False
                           If MAPISession1.SessionID > 0 Then
                              MAPISession1.SignOff
                           End If
                       Else
                          MsgBox "El cliente " + rs!VCHA_CLI_NOMBRE + " no tiene una cuenta de correo electronico", vbOKOnly, "ATENCION"
                       End If
                       rs.MoveNext
                  Wend
               Else
                  MsgBox "los clientes del embarque no estan habilitados para enviarles la información de facturas", vbOKOnly, "ATENCION"
               End If
               rs.Close
            End If
         Else
            MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
         End If
         frm_correo.Visible = False
      Else
         MsgBox "No se encuentra el archivo " + var_ruta & "\tfactura_enivar.dbf" + ", favor de verificarlo con el administrador del sistema", vbOKOnly, "ATENCION"
         Me.frm_correo.Visible = False
      End If
   End If
   Exit Sub
salir:
   MsgBox "El archivo que desea enviar es probable que esta siendo usado, salga de la aplicación y vuelvalo a intentar", vbOKOnly, "ATENCION"
   If var_tabla.State = 1 Then
      var_tabla.Close
   End If
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   Me.frm_correo.Visible = False

End Sub

Private Sub Command5_Click()
   Dim var_numero_movimientos As Double
   Dim var_numero_factura_inicio As Double
   Dim var_i As Double
   Dim var_j As Double
   Dim var_k As Double
   Dim var_cliente As String
   Dim var_expedicion As String
   Dim var_domicilio As String
   Dim var_ciudad As String
   Dim var_agente As String
   Dim var_linea As String
   Dim var_cantidad As String
   Dim var_descuento_1 As Double
   Dim var_descuento_2 As Double
   Dim var_descuento_3 As Double
   Dim var_precio As Double
   Dim var_precio_str As String
   Dim var_importe As String
   Dim var_subimporte As String
   Dim var_cantidad_letra As String
   Dim var_iva As String
   Dim var_rfc As String
   Dim var_dia As String
   Dim var_mes As String
   Dim var_año As String
   Dim var_porcentaje As Double
   Dim var_Archivo As String
   Dim var_importe_descuento_1 As Double
   Dim var_importe_descuento_2 As Double
   Dim var_importe_descuento_3 As Double
   Dim var_importe_descuento_1_2 As Double
   Dim var_importe_descuento_2_2 As Double
   Dim var_importe_descuento_3_2 As Double
   Dim var_importe_descuento_1_str As String
   Dim var_importe_descuento_2_str As String
   Dim var_importe_descuento_3_str As String
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   Dim var_marca_promocion As String
   Dim var_contador_promociones As Double
   Dim var_cantidad_total As Double
   Dim var_cantidad_total_str As String
   Dim var_factura_envio As Double
   Dim var_consecutivo As Double
   Dim var_x As Double
   cnn.CommandTimeout = 360
    
   If Trim(txt_numero_embarque) <> "" Then
      If var_estatus_embarque = "F" Then
         MsgBox "El embarque ya fue facturado con anterioridad", vbOKOnly, "ATENCION"
      Else
         'Sirve para validar que no vaya mercancia con cantidad en NULL
         Cadena = "SELECT     dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID, "
         Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO, dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID,"
         Cadena = Cadena + " dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID , dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD"
         Cadena = Cadena + " FROM         dbo.TB_DETALLE_EMBARQUES INNER JOIN"
         Cadena = Cadena + " dbo.TB_SALIDAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND"
         Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND"
         Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO AND"
         Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID"
         Cadena = Cadena + " WHERE     (dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD IS NULL) AND (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + txt_numero_embarque + ") AND"
         Cadena = Cadena + " (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')"
         rsaux4.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux4.EOF Then
            rsaux4.Close
            MsgBox "El movimiento tiene cantidad en NULL", vbOKOnly, "ATENCION"
         Else
         rsaux4.Close
         si = MsgBox("¿Deseas imprimir las facturas correspondientes al movimiento?", vbYesNo, "ATENCION")
         If si = 6 Then
            si = MsgBox("Confirmar la impresión del movimiento", vbYesNo, "ATENCION")
            If si = 6 Then
               lv_movimientos.ListItems(1).Selected = True
               var_numero_factura_inicio = lv_movimientos.selectedItem.SubItems(8)
               rs.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
               var_factura_inicio = rs!inte_ser_factura
               rs.Close
               If var_numero_factura_inicio <> var_factura_inicio Then
                  MsgBox "La numeración de facturas a cambiado, vuelva a cargar el numero de embarque", vbOKOnly, "ATENCION"
               Else
                  MsgBox "Se va a imprimir la factura " + Trim(Str(var_factura_inicio)), vbOKOnly, "ATENCION"
                  si = MsgBox("¿La impresora esta lista?", vbYesNo, "ATENCION")
                  If si = 6 Then
                     Me.frm_mensaje.Visible = True
                     Me.Refresh
                     fecha_inicio = CStr(Now)
                     Set TB_ENC_EMBARQUE_M = New TB_ENC_EMBARQUE_M
                     ok = TB_ENC_EMBARQUE_M.Anadir(var_empresa, var_unidad_organizacional, txt_numero_embarque, "F")
                     fecha_fin = CStr(Now)
                     var_estatus_embarque = "F"
                     'aqui se imprime la factura
                     
                     cnn.BeginTrans
                     rs.Open "select isnull(max(inte_tem_consecutivo),0) from tb_temp_factura_embarques", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_consecutivo = rs(0).Value
                     Else
                        var_consecutivo = 0
                     End If
                     rs.Close
                     var_consecutivo = var_consecutivo + 1
                     rs.Open "insert into tb_temp_factura_embarques (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                     cnn.CommitTrans
                     
                     Cadena = "INSERT INTO [vianney].[dbo].[TB_TEMP_FACTURA_EMBARQUES] ([INTE_TEM_CONSECUTIVO], [VCHA_AGR_AGRUPADOR_ID], [VCHA_SAL_DESCRIPCION_FACTURA], [IMPORTE], [CANTIDAD], [INTE_CAR_PLAZO], [FLOA_CAR_PORCENTAJE_IVA], [FLOA_CAR_PORCENTAJE_IMPUESTO_1], [FLOA_CAR_PORCENTAJE_IMPUESTO_2], [FLOA_CAR_PORCENTAJE_DESCUENTO_1], [FLOA_CAR_PORCENTAJE_DESCUENTO_2], [FLOA_CAR_PORCENTAJE_DESCUENTO_3], [FLOA_CAR_IMPORTE_TOTAL], [FLOA_CAR_IMPORTE_IVA], [FLOA_CAR_IMPORTE_IMPUESTO_1], [FLOA_CAR_IMPORTE_IMPUESTO_2], [FLOA_CAR_IMPORTE_DESCUENTO_1], [FLOA_CAR_IMPORTE_DESCUENTO_2],"
                     Cadena = Cadena + " [FLOA_CAR_IMPORTE_DESCUENTO_3], [FLOA_CAR_SUBIMPORTE], [FLOA_CAR_IMPORTE_NETO], [VCHA_CAR_IMPORTE_LETRA], [VCHA_SER_SERIE_ID], [VCHA_CAR_DOCUMENTO], [INTE_CAR_NUMERO], [DTIM_CAR_FECHA], [VCHA_EMP_EMPRESA_ID], [INTE_EMB_EMBARQUE], [VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_CLI_CLAVE_ID], [VCHA_CLI_NOMBRE], [VCHA_CLI_REPRESENTANTE], [VCHA_AGE_AGENTE_ID], [VCHA_RUT_RUTA_ID], [VCHA_CLI_CURP], [VCHA_CLI_RFC], [VCHA_MON_MONEDA_ID], [VCHA_PLA_PLAZO_ID], [VCHA_TCL_TIPO_CLIENTE_ID], [VCHA_LIS_LISTA_ID], [VCHA_CAN_CANAL_VENTA_ID], [VCHA_TRA_TRANSPORTE_ID], [VCHA_FAG_FAMILIA_AGRUPADOR_ID], [INTE_CLI_AGRUPADOR],"
                     Cadena = Cadena + " [INTE_CLI_ESTATUS], [VCHA_TIT_TITULAR_ID], [CHAR_PRI_PRIORIDAD_ID], [VCHA_CLI_EMAIL], [VCHA_PAI_PAIS_ID], [VCHA_PAI_NOMBRE], [VCHA_EST_ESTADO_ID], [VCHA_EST_NOMBRE], [VCHA_CIU_CIUDAD_ID], [VCHA_CIU_NOMBRE], [VCHA_CLI_COLONIA], [VCHA_CLI_DIRECCION], [VCHA_CLI_CP], [FLOA_GRE_DESCUENTO_1], [FLOA_GRE_DESCUENTO_2], [FLOA_GRE_DESCUENTO_3],  [VCHA_GRE_GRUPO_REAL_ID], [VCHA_GRE_NOMBRE], [VCHA_GAC_GRUPO_ACTUAL_ID], [VCHA_TIT_NOMBRE], [FLOA_TIT_LIMITE_CREDITO], [INTE_PLA_DIAS], [FLOA_GAC_DESCUENTO_1], [FLOA_GAC_DESCUENTO_2], [FLOA_GAC_DESCUENTO_3], [VCHA_CAN_NOMBRE], [INTE_CAN_BUSQUEDA_FACTURA_GRUPO], [FLOA_TPE_IVA], [VCHA_GAC_NOMBRE], "
                     Cadena = Cadena + " [VCHA_MON_NOMBRE], [VCHA_MON_NOMBRE_PLURAL], [VCHA_AGE_NOMBRE], [VCHA_ESB_ESTABLECIMIENTO_ID], [FLOA_CAR_TIPO_CAMBIO], [INTE_ORS_ORDEN_SURTIDO], [INTE_PED_NUMERO], [FLOA_SAL_PROMOCION_1],  [FLOA_SAL_PROMOCION_2], [VCHA_CAR_TIPO_DOCUMENTO], [VCHA_UOR_UNIDAD_ID], [INTE_JAU_JAULA_ID], [VCHA_VEH_VEHICULO_ID], [DTIM_EMB_FECHA_INICIO], [DTIM_EMB_FECHA_FINAL], [CHAR_EMB_ESTATUS], [VCHA_CHO_CHOFER_ID], [FLOA_EMB_CUBICAJE], [CHAR_CAR_TIPO_FACTURACION], [VCHA_CAR_CLASE_ID], [CHAR_CAR_AFECTACION], [VCHA_ALM_ALMACEN_ID], [Expr1], [INTE_EMO_NUMERO], [Expr2], [Expr3], [Expr4], [Expr5], [Expr6], [Expr7], [VCHA_AUD_USUARIO], [VCHA_AUD_MAQUINA], [VCHA_AUD_FECHA], [FLOA_CAR_SALDO], "
                     Cadena = Cadena + " [DTIM_CAR_FECHA_VENCIMIENTO], [DTIM_CAR_FECHA_ENTREGA],[Expr8], [CHAR_CAR_ESTATUS], [DTIM_CAR_FECHA_CANCELACION], [VCHA_CAR_USUARIO_CANCELACION],  [VCHA_CAR_MAQUINA_CANCELACION], [INTE_CLI_ENVIO_FACTURA], [FLOA_SAL_PRECIO_PROMEDIO],  [INTE_CAR_FACTURA_CEROS], [FLOA_CAR_COSTO], [INTE_SAL_CONSECUTIVO_FACTURA])"
                     Cadena = Cadena + " select " + CStr(var_consecutivo) + ", VCHA_AGR_AGRUPADOR_ID, VCHA_SAL_DESCRIPCION_FACTURA, IMPORTE, CANTIDAD, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3, FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_SER_SERIE_ID, VCHA_CAR_DOCUMENTO, INTE_CAR_NUMERO, DTIM_CAR_FECHA, VCHA_EMP_EMPRESA_ID, INTE_EMB_EMBARQUE, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, "
                     Cadena = Cadena + " VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_CLI_REPRESENTANTE, VCHA_AGE_AGENTE_ID, VCHA_RUT_RUTA_ID, VCHA_CLI_CURP, VCHA_CLI_RFC, VCHA_MON_MONEDA_ID, VCHA_PLA_PLAZO_ID, VCHA_TCL_TIPO_CLIENTE_ID, VCHA_LIS_LISTA_ID, VCHA_CAN_CANAL_VENTA_ID, VCHA_TRA_TRANSPORTE_ID, VCHA_FAG_FAMILIA_AGRUPADOR_ID, INTE_CLI_AGRUPADOR, INTE_CLI_ESTATUS, VCHA_TIT_TITULAR_ID, CHAR_PRI_PRIORIDAD_ID, VCHA_CLI_EMAIL, VCHA_PAI_PAIS_ID, VCHA_PAI_NOMBRE, VCHA_EST_ESTADO_ID, VCHA_EST_NOMBRE, VCHA_CIU_CIUDAD_ID, VCHA_CIU_NOMBRE, VCHA_CLI_COLONIA, VCHA_CLI_DIRECCION, VCHA_CLI_CP, FLOA_GRE_DESCUENTO_1, FLOA_GRE_DESCUENTO_2, FLOA_GRE_DESCUENTO_3, VCHA_GRE_GRUPO_REAL_ID, VCHA_GRE_NOMBRE, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_TIT_NOMBRE, FLOA_TIT_LIMITE_CREDITO, INTE_PLA_DIAS, "
                     Cadena = Cadena + " FLOA_GAC_DESCUENTO_1, FLOA_GAC_DESCUENTO_2, FLOA_GAC_DESCUENTO_3, VCHA_CAN_NOMBRE, INTE_CAN_BUSQUEDA_FACTURA_GRUPO, FLOA_TPE_IVA, VCHA_GAC_NOMBRE, VCHA_MON_NOMBRE, VCHA_MON_NOMBRE_PLURAL, VCHA_AGE_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, FLOA_CAR_TIPO_CAMBIO, INTE_ORS_ORDEN_SURTIDO, INTE_PED_NUMERO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, VCHA_CAR_TIPO_DOCUMENTO, VCHA_UOR_UNIDAD_ID, INTE_JAU_JAULA_ID, VCHA_VEH_VEHICULO_ID, DTIM_EMB_FECHA_INICIO, DTIM_EMB_FECHA_FINAL, CHAR_EMB_ESTATUS, VCHA_CHO_CHOFER_ID, FLOA_EMB_CUBICAJE, CHAR_CAR_TIPO_FACTURACION, VCHA_CAR_CLASE_ID, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, Expr1, INTE_EMO_NUMERO, Expr2, Expr3, Expr4, Expr5, Expr6, Expr7, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, "
                     Cadena = Cadena + " FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, Expr8, CHAR_CAR_ESTATUS, DTIM_CAR_FECHA_CANCELACION, VCHA_CAR_USUARIO_CANCELACION, VCHA_CAR_MAQUINA_CANCELACION, INTE_CLI_ENVIO_FACTURA, FLOA_SAL_PRECIO_PROMEDIO, INTE_CAR_FACTURA_CEROS, FLOA_CAR_COSTO, INTE_SAL_CONSECUTIVO_FACTURA from vw_facturas_embarque where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque
                     
                     rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     Me.frm_mensaje.Visible = False
                     If var_empresa = "02" Then
                        rsaux3.Open "select distinct inte_car_numero from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           var_Archivo = App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat"
                           Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat") For Output As #2
                           While Not rsaux3.EOF
                              If rs.State = 1 Then
                                 rs.Close
                              End If
                              If var_empresa <> "03" Then
                                 rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY INTE_SAL_CONSECUTIVO_FACTURA", cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY vcha_sal_descripcion_factura", cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If Not rs.EOF Then
                                'AQUI EMPIEZA LA FACTURA
                                 Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt") For Output As #1
                                  'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                                 'Print #1, Chr(27) + Chr(67) + Chr(44) + Chr(13)
                                 'Print #1, ""
                                 Print #1, Chr(15) + Chr(27) + Chr(64)
                                 Print #1, Spc(105); Str(rsaux3!inte_Car_numero)
                                 Print #1, ""
                                 Print #1, ""
                                    Print #1, Spc(105); Str(rs!INTE_CAR_PLAZO) + " DIAS DE VENCIMIENTO" + "                  " + Format(rs!dtim_Car_fecha, "Short Date")
                                 Print #1, ""
                                 'Print #1, Spc(92); Str(rs!inte_car_PLAZO) + " DIAS DE VENCIMIENTO"
                                 var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                                 For var_j = 1 + Len(Trim(var_cliente)) To 83
                                     var_cliente = var_cliente + " "
                                 Next var_j
                                 If var_unidad_organizacional = "21" Then
                                    var_cliente = var_cliente + "               MEXICO, D.F."
                                 Else
                                    var_cliente = var_cliente + "               AGUASCALIENTES, AGS."
                                 End If
                                 Print #1, Spc(10); var_cliente
                                 var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                                 For var_j = 1 + Len(Trim(var_domicilio)) To 83
                                     var_domicilio = var_domicilio + " "
                                 Next var_j
                                 var_agente = ""
                                 var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                                 For var_j = 1 + Len(Trim(var_agente)) To 8
                                     var_agente = var_agente + " "
                                 Next var_j
                                 rsaux4.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux4.EOF Then
                                    var_agente = var_agente + IIf(IsNull(rsaux4!VCHA_AGE_NOMBRE), "", rsaux4!VCHA_AGE_NOMBRE)
                                 Else
                                    var_agente = var_agente + ""
                                 End If
                                 rsaux4.Close
                                 var_domicilio = var_domicilio
                                 'Print #1, Spc(111); var_agente
                                 Print #1, Spc(10); var_domicilio
                                 var_ciudad = ""
                                 var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                                 For var_j = 1 + Len(Trim(var_ciudad)) To 37
                                    var_ciudad = var_ciudad + " "
                                 Next var_j
                                 
                                 var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                                 var_ciudad = var_ciudad
                                 var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                 var_ciudad = var_ciudad + var_rfc
                                 
                                 For var_j = 1 + Len(Trim(var_estado)) To 46
                                    var_estado = var_estado + " "
                                 Next var_j
                                 
   
                                 For var_j = 1 + Len(Trim(var_ciudad)) To 14
                                    var_ciudad = var_ciudad + " "
                                 Next var_j
                                  
                                 var_ciudad = var_ciudad + "                                                      " + var_agente
                                 
                                 VAR_EMBARQUE = "EMB.: " + txt_numero_embarque
                                 var_ordern_surtido = x
                                 Print #1, Spc(10); var_ciudad
                                 var_rfc = "RFC:  " + var_rfc
                                 var_rfc = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                                 For var_j = 1 + Len(Trim(var_rfc)) To 89
                                    var_rfc = var_rfc + " "
                                 Next var_j
                                 var_rfc = var_rfc + "               PED.: " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))) + " "
                                 var_rfc = var_rfc + " O.S.: " + Trim(Str(IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), 0, rs!INTE_ORS_ORDEN_SURTIDO))) + " " + VAR_EMBARQUE
                                 Print #1, var_rfc
                                 'Print #1, Spc(10); IIf(IsNull(rs!vcha_esb_establecimiento_id), "", rs!vcha_esb_establecimiento_id)
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
                                       var_i = 0
                                       
                                       
                                       ''' imprimir cantidad en la orilla
                                       var_cantidad_nueva = Format(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                                       If Len(Trim(var_cantidad_nueva)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_cantidad_nueva)) To 14
                                             var_cantidad_nueva = " " + var_cantidad_nueva
                                          Next var_j
                                       End If
                                       While Len((var_linea)) < 60
                                             var_linea = var_linea + " "
                                       Wend
                                       var_linea = var_linea + var_cantidad_nueva
                                       
                                       ''' imprimir cantidad en la orilla
                                       
                                       
                                       
                                       
                                       
                                       While Len((var_linea)) < 115
                                             var_linea = var_linea + " "
                                       Wend
                                       var_linea = var_linea + " "
                                       var_linea = var_linea + var_marca_promocion
                                       var_cantidad = Format(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                                       var_cantidad_total = var_cantidad_total + IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                                       If Len(Trim(var_cantidad)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_cantidad)) To 14
                                             var_cantidad = " " + var_cantidad
                                          Next var_j
                                       End If
                                       var_precio = IIf(IsNull(rs!Importe), 0, rs!Importe)
                                       var_descuento_1 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
                                       var_descuento_2 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2)
                                       var_descuento_3 = IIf(IsNull(rs!floa_car_porcentaje_descuento_3), 0, rs!floa_car_porcentaje_descuento_3)
                                       var_porcentaje = (100 - var_descuento_1) / 100
                                       var_precio = var_precio * var_porcentaje
                                       var_importe_descuento_1_2 = (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                       var_importe_descuento_1 = var_importe_descuento_1 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                       var_precio = var_precio * ((100 - var_descuento_2) / 100)
                                       var_importe_descuento_2 = var_importe_descuento_2 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - (var_importe_descuento_1_2 + var_precio))
                                       var_precio = var_precio * ((100 - var_descuento_3) / 100)
                                       var_precio = var_precio / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                                       'var_precio_str = Format(var_precio / IIf(IsNull(rs!cantidad), 0, rs!cantidad), "###,###,##0.00")
                                       var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                       If Len(Trim(var_rfc)) > 0 Then
                                          var_precio_str = Format(IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                                       Else
                                          var_precio_str = Format((IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) * (1 + (rs!floa_car_porcentaje_iva / 100)), "###,###,##0.00")
                                       End If
                                       If Len(Trim(var_precio_str)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_precio_str)) To 14
                                              var_precio_str = " " + var_precio_str
                                          Next var_j
                                       End If
                                       var_linea = var_linea + var_cantidad + var_precio_str
                                       If Len(Trim(var_rfc)) > 0 Then
                                          var_importe = Format((IIf(IsNull(rs!Importe), 0, rs!Importe)), "###,###,##0.00")
                                          If Len(Trim(var_importe)) < 14 Then
                                              For var_j = 1 + Len(Trim(var_importe)) To 14
                                                 var_importe = " " + var_importe
                                              Next var_j
                                          End If
                                       Else
                                          var_importe = Format((IIf(IsNull(rs!Importe), 0, rs!Importe) * (1 + (rs!floa_car_porcentaje_iva / 100))), "###,###,##0.00")
                                          If Len(Trim(var_importe)) < 14 Then
                                              For var_j = 1 + Len(Trim(var_importe)) To 14
                                                 var_importe = " " + var_importe
                                              Next var_j
                                          End If
                                       End If
                                       var_linea = var_linea + var_importe
                                        
                                       Print #1, var_linea
                                       rs.MoveNext
                                    Else
                                       Print #1, ""
                                    End If
                                 Next var_k
                                 Print #1, ""
                                 'Print #1, ""
                                 rs.MoveFirst
                                 
                                 var_cantidad_total_str = Format(var_cantidad_total, "###,###,##0.00")
                                 var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                 If Len(Trim(var_rfc)) > 0 Then
                                    var_cantidad_letra = rs!vcha_car_importe_letra
                                    var_importe_descuento_1_str = Format(IIf(IsNull(rs!floa_car_importe_descuento_1), 0, rs!floa_car_importe_descuento_1) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio), "###,###,##0.00")
                                    If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                            var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                       Next var_j
                                    End If
                                    var_importe_descuento_2_str = Format(IIf(IsNull(rs!floa_car_importe_descuento_2), 0, rs!floa_car_importe_descuento_2) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio), "###,###,##0.00")
                                    If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                           var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                       Next var_j
                                    End If
                                 Else
                                    var_cantidad_letra = rs!vcha_car_importe_letra
                                    var_importe_descuento_1_str = Format((IIf(IsNull(rs!floa_car_importe_descuento_1), 0, rs!floa_car_importe_descuento_1)) * (1 + (rs!floa_car_porcentaje_iva / 100) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                    If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                            var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                       Next var_j
                                    End If
                                    var_importe_descuento_2_str = Format((IIf(IsNull(rs!floa_car_importe_descuento_2), 0, rs!floa_car_importe_descuento_2)) * (1 + (rs!floa_car_porcentaje_iva / 100) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                    If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                           var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                       Next var_j
                                    End If
                                 End If
                                 var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                                 If Len(Trim(var_linea)) < 145 Then
                                    For var_j = 1 + Len(Trim(var_linea)) To 145
                                        var_linea = var_linea + " "
                                    Next var_j
                                 End If
                                 Print #1, var_linea + var_importe_descuento_1_str
                                 If var_empresa = "18" Then
                                    var_linea = ""
                                 Else
                                    If Trim(var_cliente_coppel) = "C000002947" Then
                                       var_linea = ""
                                    Else
                                       var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%"
                                    End If
                                 End If
                                 If Len(Trim(var_linea)) < 145 Then
                                    For var_j = 1 + Len(Trim(var_linea)) To 145
                                        var_linea = var_linea + " "
                                    Next var_j
                                 End If
                                 var_linea = var_linea + var_importe_descuento_2_str
                                 Print #1, var_linea
                                 If var_contador_promociones > 0 Then
                                    If var_cliente_sigo = "C000001636" Then
                                       Print #1, "Descuento adicional del 2%"
                                    Else
                                       Print #1, var_cadena_promocion_171209
                                    End If
                                 Else
                                    If var_cliente_sigo = "C000001636" Then
                                       Print #1, "Descuento adicional del 2%"
                                    Else
                                       Print #1, ""
                                    End If
                                 End If
                                 
                                 var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                 var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                                 
                                 If Len(Trim(var_linea)) < 117 Then
                                    For var_j = 1 + Len(Trim(var_linea)) To 117
                                        var_x = var_j Mod 2
                                        If var_x >= 1 Then
                                           var_linea = " " + var_linea
                                        Else
                                           var_linea = var_linea + " "
                                        End If
                                    Next var_j
                                 End If
                                 
                                 If Len(Trim(var_rfc)) = 0 Then
                                    var_subimporte = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                    If Len(Trim(var_subimporte)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                           var_subimporte = " " + var_subimporte
                                       Next var_j
                                    End If
                                    var_iva = "-"
                                    For var_j = 1 + Len(Trim(var_iva)) To 11
                                        var_iva = " " + var_iva
                                     Next var_j
                                 Else
                                    var_subimporte = Format(Round(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                    If Len(Trim(var_subimporte)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                           var_subimporte = " " + var_subimporte
                                       Next var_j
                                    End If
                                    var_iva = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                    If Len(Trim(var_iva)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_iva)) To 14
                                           var_iva = " " + var_iva
                                       Next var_j
                                    End If
                                 End If
                                 
                                 If Len(Trim(var_subimporte)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                        var_subimporte = " " + var_subimporte
                                    Next var_j
                                 End If
                                 var_espacios = 131 - Len(var_cantidad_total_str)
                                 var_cantidad_total_str = Trim(var_cantidad_total_str)
                                 If Len(Trim(var_cantidad_total_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_cantidad_total_str)) To 14
                                        var_cantidad_total_str = " " + var_cantidad_total_str
                                    Next var_j
                                 End If
                                 var_subimporte = Trim(var_subimporte)
                                 If Len(Trim(var_subimporte)) < 24 Then
                                    For var_j = 1 + Len(Trim(var_subimporte)) To 24
                                        var_subimporte = " " + var_subimporte
                                    Next var_j
                                 End If
                              
                                 var_cantidad_total_str = var_linea + var_cantidad_total_str + "    " + var_subimporte
                                 'Print #1, Spc(var_espacios); var_cantidad_total_str; Spc(8); var_subimporte
                                 Print #1, var_cantidad_total_str
                                 var_linea = "                                                                          ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                        " + var_iva
                                 Print #1, var_linea
                                 var_dia = Day(rs!dtim_Car_fecha)
                                 var_mes = Month(rs!dtim_Car_fecha)
                                 var_año = Year(rs!dtim_Car_fecha)
                                 
                                 var_linea = "                                                             " + CStr(var_dia) + "     " + CStr(var_mes)
                                 
                                 If Len(var_linea) < 145 Then
                                    For var_j = 1 + Len(var_linea) To 145
                                        var_linea = var_linea + " "
                                    Next var_j
                                 End If
                                 
                                 var_importe = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                 
                                 If Len(Trim(var_importe)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe)) To 14
                                        var_importe = " " + var_importe
                                    Next var_j
                                 End If
                              
                                 'var_linea = "                                                                   ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                               " + var_iva
                                 'var_linea = "                                                                                                                                                 " + var_importe
                                 
                                 var_linea = var_linea + var_importe
                                 Print #1, var_linea
                                 
                                 var_linea = var_importe
                                 If Len(Trim(var_linea)) < 20 Then
                                    For var_j = 1 + Len(Trim(var_linea)) To 20
                                        var_linea = " " + var_linea
                                    Next var_j
                                 End If
                                 var_linea = var_linea + " " + var_cantidad_letra
                                 Print #1, Spc(2); CStr(var_año); var_linea
                                 
                                 var_linea = ""
                                 Print #1, ""
                                 Print #1, ""
                                 Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                                 Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!VCHA_CLI_COLONIA), "", rs!VCHA_CLI_COLONIA))
                                 Print #1, Spc(5); Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre))
                                 If var_empresa <> "03" Then
                                    Print #1, ""
                                    Print #1, ""
                                 Else
                                    Print #1, ""
                                    Print #1, ""
                                 End If
                                 Print #1, ""
                                 Print #1, ""
                                 Close #1
                                 If Trim(var_empresa) = "02" Then
                                    Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt lpt1"
                                 Else
                                    Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt lpt1"
                                 End If
                                 'AQUI TERMINA LA FACTURA
                              End If
                              rs.Close
                              rsaux3.MoveNext
                           Wend
                           Close #2
                           x = Shell(var_Archivo, vbHide)
                        End If
                        rsaux3.Close
                        'Aqui se termina de imprimir la factura
                     End If
                     If var_empresa = "03" Then
                     
                     rsaux3.Open "select distinct inte_car_numero from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux3.EOF Then
                        var_Archivo = App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat"
                        Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat") For Output As #2
                        While Not rsaux3.EOF
                           If rs.State = 1 Then
                              rs.Close
                           End If
                           If var_empresa <> "03" Then
                              rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY INTE_SAL_CONSECUTIVO_FACTURA", cnn, adOpenDynamic, adLockOptimistic
                           Else
                              rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY vcha_sal_descripcion_factura", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           If Not rs.EOF Then
                             'AQUI EMPIEZA LA FACTURA
                              Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt") For Output As #1
                              'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                              'Print #1, Chr(27) + Chr(67) + Chr(44) + Chr(13)
                              'Print #1, ""
                              Print #1, Chr(15) + Chr(27) + Chr(64)
                              Print #1, Spc(105); Str(rsaux3!inte_Car_numero)
                              Print #1, ""
                              Print #1, ""
                              Print #1, Spc(105); Str(rs!INTE_CAR_PLAZO) + " DIAS DE VENCIMIENTO" + "                  " + Format(rs!dtim_Car_fecha, "Short Date")
                              Print #1, ""
                              'Print #1, Spc(92); Str(rs!inte_car_PLAZO) + " DIAS DE VENCIMIENTO"
                              var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                              For var_j = 1 + Len(Trim(var_cliente)) To 83
                                  var_cliente = var_cliente + " "
                              Next var_j
                              If var_unidad_organizacional = "21" Then
                                 var_cliente = var_cliente + "               MEXICO, D.F."
                              Else
                                 var_cliente = var_cliente + "               AGUASCALIENTES, AGS."
                              End If
                              Print #1, Spc(10); var_cliente
                              var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                              For var_j = 1 + Len(Trim(var_domicilio)) To 83
                                  var_domicilio = var_domicilio + " "
                              Next var_j
                              var_agente = ""
                              var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                              For var_j = 1 + Len(Trim(var_agente)) To 8
                                  var_agente = var_agente + " "
                              Next var_j
                              rsaux4.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux4.EOF Then
                                 var_agente = var_agente + IIf(IsNull(rsaux4!VCHA_AGE_NOMBRE), "", rsaux4!VCHA_AGE_NOMBRE)
                              Else
                                 var_agente = var_agente + ""
                              End If
                              rsaux4.Close
                              var_domicilio = var_domicilio
                              'Print #1, Spc(111); var_agente
                              Print #1, Spc(10); var_domicilio
                              var_ciudad = ""
                              var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                              For var_j = 1 + Len(Trim(var_ciudad)) To 37
                                 var_ciudad = var_ciudad + " "
                              Next var_j
                              
                              var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre) + " " + IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
                              var_ciudad = var_ciudad
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              var_ciudad = var_ciudad + var_rfc
                              
                              For var_j = 1 + Len(Trim(var_estado)) To 46
                                 var_estado = var_estado + " "
                              Next var_j
                              

                              For var_j = 1 + Len(Trim(var_ciudad)) To 14
                                 var_ciudad = var_ciudad + " "
                              Next var_j
                               
                              var_ciudad = var_ciudad + "                                                      " + var_agente
                              
                              VAR_EMBARQUE = "EMB.: " + txt_numero_embarque
                              var_ordern_surtido = x
                              Print #1, Spc(10); var_ciudad
                              var_rfc = "RFC:  " + var_rfc
                              var_rfc = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre) + ", " + IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
                              For var_j = 1 + Len(Trim(var_rfc)) To 89
                                 var_rfc = var_rfc + " "
                              Next var_j
                              var_rfc = var_rfc + "               PED.: " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))) + " "
                              var_rfc = var_rfc + " O.S.: " + Trim(Str(IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), 0, rs!INTE_ORS_ORDEN_SURTIDO))) + " " + VAR_EMBARQUE
                              Print #1, var_rfc
                              'Print #1, Spc(10); IIf(IsNull(rs!vcha_esb_establecimiento_id), "", rs!vcha_esb_establecimiento_id)
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
                                    var_i = 0
                                    While Len((var_linea)) < 115
                                          var_linea = var_linea + " "
                                    Wend
                                    var_linea = var_linea + " "
                                    var_linea = var_linea + var_marca_promocion
                                    var_cantidad = Format(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                                    var_cantidad_total = var_cantidad_total + IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                                    If Len(Trim(var_cantidad)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_cantidad)) To 14
                                          var_cantidad = " " + var_cantidad
                                       Next var_j
                                    End If
                                    var_precio = IIf(IsNull(rs!Importe), 0, rs!Importe)
                                    var_descuento_1 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
                                    var_descuento_2 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2)
                                    var_descuento_3 = IIf(IsNull(rs!floa_car_porcentaje_descuento_3), 0, rs!floa_car_porcentaje_descuento_3)
                                    
                                    var_porcentaje = (100 - var_descuento_1) / 100
                                    var_precio = var_precio * var_porcentaje
                                    var_importe_descuento_1_2 = (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                    var_importe_descuento_1 = var_importe_descuento_1 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                    var_precio = var_precio * ((100 - var_descuento_2) / 100)
                                    var_importe_descuento_2 = var_importe_descuento_2 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - (var_importe_descuento_1_2 + var_precio))
                                    var_precio = var_precio * ((100 - var_descuento_3) / 100)
                                    var_precio = var_precio / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                                    'var_precio_str = Format(var_precio / IIf(IsNull(rs!cantidad), 0, rs!cantidad), "###,###,##0.00")
                                    var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                    
                                    If Len(Trim(var_rfc)) > 0 Then
                                       var_importe_precio = IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                                       var_importe_precio = var_importe_precio * ((100 - var_descuento_1) / 100)
                                       var_importe_precio = var_importe_precio * ((100 - var_descuento_2) / 100)
                                       var_importe_precio = var_importe_precio * ((100 - var_descuento_3) / 100)
                                       var_precio_str = Format(var_importe_precio, "###,###,##0.00")
                                    Else
                                       var_importe_precio = (IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) * (1 + (rs!floa_car_porcentaje_iva / 100))
                                       var_importe_precio = var_importe_precio * ((100 - var_descuento_1) / 100)
                                       var_importe_precio = var_importe_precio * ((100 - var_descuento_2) / 100)
                                       var_importe_precio = var_importe_precio * ((100 - var_descuento_3) / 100)
                                       var_precio_str = Format(var_importe_precio, "###,###,##0.00")
                                    End If
                                    
                                    If Len(Trim(var_precio_str)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_precio_str)) To 14
                                           var_precio_str = " " + var_precio_str
                                       Next var_j
                                    End If
                                    var_linea = var_linea + var_cantidad + var_precio_str
                                    If Len(Trim(var_rfc)) > 0 Then
                                       
                                       var_importe_G = IIf(IsNull(rs!Importe), 0, rs!Importe)
                                       var_importe_G = var_importe_G * ((100 - var_descuento_1) / 100)
                                       var_importe_G = var_importe_G * ((100 - var_descuento_2) / 100)
                                       var_importe_G = var_importe_G * ((100 - var_descuento_3) / 100)
                                       var_importe = Format(var_importe_G, "###,###,##0.00")
                                       
                                       If Len(Trim(var_importe)) < 14 Then
                                           For var_j = 1 + Len(Trim(var_importe)) To 14
                                              var_importe = " " + var_importe
                                           Next var_j
                                       End If
                                    Else
                                       var_importe_G = IIf(IsNull(rs!Importe), 0, rs!Importe) * (1 + (rs!floa_car_porcentaje_iva / 100))
                                       
                                       var_importe_G = var_importe_G * ((100 - var_descuento_1) / 100)
                                       var_importe_G = var_importe_G * ((100 - var_descuento_2) / 100)
                                       var_importe_G = var_importe_G * ((100 - var_descuento_3) / 100)
                                       var_importe = Format(var_importe_G, "###,###,##0.00")
                                       
                                       If Len(Trim(var_importe)) < 14 Then
                                           For var_j = 1 + Len(Trim(var_importe)) To 14
                                              var_importe = " " + var_importe
                                           Next var_j
                                       End If
                                    End If
                                    var_linea = var_linea + var_importe
                                     
                                    Print #1, var_linea
                                    rs.MoveNext
                                 Else
                                    Print #1, ""
                                 End If
                              Next var_k
                              'Print #1, ""
                              rs.MoveFirst
                              
                              var_cantidad_total_str = Format(var_cantidad_total, "###,###,##0.00")
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              If Len(Trim(var_rfc)) > 0 Then
                                 var_cantidad_letra = rs!vcha_car_importe_letra
                                 'var_importe_descuento_1_str = Format(IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_1), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_1) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio), "###,###,##0.00")
                                 var_importe_descuento_1_str = Format(0, "###,###,##0.00")
                                 If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                         var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                    Next var_j
                                 End If
                                 'var_importe_descuento_2_str = Format(IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_2), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_2) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio), "###,###,##0.00")
                                 var_importe_descuento_2_str = Format(0, "###,###,##0.00")
                                 If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                        var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                    Next var_j
                                 End If
                              Else
                                 var_cantidad_letra = rs!vcha_car_importe_letra
                                 'var_importe_descuento_1_str = Format((IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_1), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_1)) * (1 + (rs!floa_car_porcentaje_iva / 100) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                 var_importe_descuento_1_str = Format(0, "###,###,##0.00")
                                 
                                 If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                         var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                    Next var_j
                                 End If
                                 'var_importe_descuento_2_str = Format((IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_2), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_2)) * (1 + (rs!floa_car_porcentaje_iva / 100) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                 var_importe_descuento_2_str = Format(0, "###,###,##0.00")
                                 If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                        var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                    Next var_j
                                 End If
                              End If
                              var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                              If Len(Trim(var_linea)) < 145 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 145
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              Print #1, var_linea + var_importe_descuento_1_str
                              If var_empresa = "18" Then
                                 var_empresa = ""
                              Else
                                 If Trim(var_cliente_coppel) = "C000002947" Then
                                    var_linea = ""
                                 Else
                                    var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%"
                                 End If
                              End If
                              If Len(Trim(var_linea)) < 145 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 145
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              var_linea = var_linea + var_importe_descuento_2_str
                              Print #1, var_linea
                              If var_contador_promociones > 0 Then
                                 If var_cliente_sigo = "C000001636" Then
                                    Print #1, "Descuento adicional del 2%"
                                 Else
                                    Print #1, var_cadena_promocion_171209
                                 End If
                              Else
                                 If var_cliente_sigo = "C000001636" Then
                                    Print #1, "Descuento adicional del 2%"
                                 Else
                                    Print #1, ""
                                 End If
                              End If
                              
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                              
                              If Len(Trim(var_linea)) < 117 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 117
                                     var_x = var_j Mod 2
                                     If var_x >= 1 Then
                                        var_linea = " " + var_linea
                                     Else
                                        var_linea = var_linea + " "
                                     End If
                                 Next var_j
                              End If
                              
                              If Len(Trim(var_rfc)) = 0 Then
                                 var_subimporte = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                 If Len(Trim(var_subimporte)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                        var_subimporte = " " + var_subimporte
                                    Next var_j
                                 End If
                                 var_iva = "-"
                                 For var_j = 1 + Len(Trim(var_iva)) To 11
                                     var_iva = " " + var_iva
                                  Next var_j
                              Else
                                 var_subimporte = Format(Round(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                 If Len(Trim(var_subimporte)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                        var_subimporte = " " + var_subimporte
                                    Next var_j
                                 End If
                                 var_iva = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                 If Len(Trim(var_iva)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_iva)) To 14
                                        var_iva = " " + var_iva
                                    Next var_j
                                 End If
                              End If
                              
                              If Len(Trim(var_subimporte)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                     var_subimporte = " " + var_subimporte
                                 Next var_j
                              End If
                              var_espacios = 131 - Len(var_cantidad_total_str)
                              var_cantidad_total_str = Trim(var_cantidad_total_str)
                              If Len(Trim(var_cantidad_total_str)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_cantidad_total_str)) To 14
                                     var_cantidad_total_str = " " + var_cantidad_total_str
                                 Next var_j
                              End If
                              var_subimporte = Trim(var_subimporte)
                              If Len(Trim(var_subimporte)) < 24 Then
                                 For var_j = 1 + Len(Trim(var_subimporte)) To 24
                                     var_subimporte = " " + var_subimporte
                                 Next var_j
                              End If
                              
                              var_cantidad_total_str = var_linea + var_cantidad_total_str + "    " + var_subimporte
                              'Print #1, Spc(var_espacios); var_cantidad_total_str; Spc(8); var_subimporte
                              Print #1, var_cantidad_total_str
                              var_linea = "                                                                          ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                        " + var_iva
                              Print #1, var_linea
                              var_dia = Day(rs!dtim_Car_fecha)
                              var_mes = Month(rs!dtim_Car_fecha)
                              var_año = Year(rs!dtim_Car_fecha)
                              
                              var_linea = "                                                             " + CStr(var_dia) + "     " + CStr(var_mes)
                              
                              If Len(var_linea) < 145 Then
                                 For var_j = 1 + Len(var_linea) To 145
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              
                              var_importe = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                              
                              If Len(Trim(var_importe)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_importe)) To 14
                                     var_importe = " " + var_importe
                                 Next var_j
                              End If
                              
                              'var_linea = "                                                                   ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                               " + var_iva
                              'var_linea = "                                                                                                                                                 " + var_importe
                              
                              var_linea = var_linea + var_importe
                              Print #1, var_linea
                              
                              var_linea = var_importe
                              If Len(Trim(var_linea)) < 20 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 20
                                     var_linea = " " + var_linea
                                 Next var_j
                              End If
                              var_linea = var_linea + " " + var_cantidad_letra
                              Print #1, Spc(2); CStr(var_año); var_linea
                              
                              Print #1, ""
                              var_linea = ""
                              Print #1, ""
                              Print #1, ""
                              Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                              Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!VCHA_CLI_COLONIA), "", rs!VCHA_CLI_COLONIA))
                              Print #1, Spc(5); Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre))
                              If var_empresa <> "03" Then
                                 Print #1, ""
                                 Print #1, ""
                              Else
                                 Print #1, ""
                                 Print #1, ""
                              End If
                              Print #1, ""
                              Print #1, ""
                              Close #1
                              If Trim(var_empresa) = "02" Then
                                 Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt lpt1"
                              Else
                                 Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt lpt1"
                              End If
                              'AQUI TERMINA LA FACTURA
                           End If
                           rs.Close
                           rsaux3.MoveNext
                        Wend
                        Close #2
                        x = Shell(var_Archivo, vbHide)
                     End If
                     rsaux3.Close
                     'Aqui se termina de imprimir la factura
                     
                     
                     End If
                     rsaux3.Open "delete from TB_TEMP_FACTURA_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                     MsgBox "Se a terminado el proceso de facturación", vbOKOnly, "ATENCION"
                     var_activa_forma_informacion_pedido_sugerido = Me.Name
                     frminformacion_pedido_sugerido_rutas.Show
                     Me.Enabled = False
                  Else
                     MsgBox "Se a cancelado la impresión de las facturas", vbOKOnly, "ATENCION"
                  End If
               End If
            Else
               MsgBox "Se a cancelado la impresión de las facturas", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Se a cancelado la impresión de las facturas", vbOKOnly, "ATENCION"
         End If
         End If
      End If
   Else
      MsgBox "No se a seleccionado un embarque", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
   If Shift = 4 And KeyCode = 69 Then
      cmd_correo_Click
   End If
   If Shift = 4 And KeyCode = 82 Then
      cmd_relacion_facturas_Click
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If Me.frm_correo.Visible = True Then
         Me.frm_correo.Visible = False
      Else
         If Me.frm_embarque_envio.Visible = True Then
            Me.frm_embarque_envio.Visible = False
         Else
            If Me.frm_embarque_relacion.Visible = True Then
               Me.frm_embarque_relacion.Visible = False
            Else
               If Me.frm_embarques_vivos.Visible = True Then
                  Me.frm_embarques_vivos.Visible = False
               Else
                  If Me.frm_envio_informacion.Visible = True Then
                     Me.frm_envio_informacion.Visible = False
                  Else
                     Unload Me
                  End If
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub Form_Load()
  
   'If var_empresa <> "18" And var_empresa <> "16" Then
   '   If cnn_clientes_tiendas.State = 0 Then
   '      cnn_clientes_tiendas.Open var_conexion_pedidos_tiendas
   '      cnn_clientes_tiendas.CursorLocation = adUseClient
   '   End If
   'End If
   Me.frm_reimpresion.Visible = False
   Me.frm_reimpresion_nueva.Visible = False
   Me.frm_mensaje.Visible = False
   var_cadena_seguridad = ""
   frm_embarque_relacion.Visible = False
   frm_embarque_correo_ft.Visible = False
   Me.frm_embarque_envio.Visible = False
   frm_correo.Visible = False
   frm_embarques_vivos.Visible = False
   Me.frm_envio_informacion.Visible = False
   Dim var_contador_serie As Double
   Set var_tabla = CreateObject("ADODB.connection")
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open "select * from tb_principal where vcha_emp_Empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   var_ruta = IIf(IsNull(rs!VCHA_PRI_RUTA_ENVIOS_FACTURAS), "", rs!VCHA_PRI_RUTA_ENVIOS_FACTURAS)
   var_renglones_factura = rs!INTE_PRI_RENGLONES_FACTURA
   If var_empresa = "31" Then
      var_renglones_factura = 7
   End If
   rs.Close
   Top = 0
   Left = 0
   rs.Open "select vcha_ser_serie_id from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_contador_serie = 0
      While Not rs.EOF
         var_contador_serie = var_contador_serie + 1
         rs.MoveNext
      Wend
      rs.MoveFirst
      txt_numero_embarque.Enabled = True
      Call RecsetToCombo(cmb_series.hwnd, rs, 0)
      If var_contador_serie > 1 Then
         cmb_series.Enabled = True
      Else
         cmb_series.Enabled = False
      End If
      rs.MoveFirst
      cmb_series = rs!vcha_Ser_Serie_id
      var_serie = rs!vcha_Ser_Serie_id
      If var_empresa <> "16" Then
         rsaux2.Open "select * from VW_EMBARQUES_CERRADOS_NO_FACTURADOS where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux2.EOF Then
            'Me.Enabled = False
            var_activa_forma_detalle_cajas = Me.Name
            If var_empresa <> "31" Then
               frmembarques_cerrados_no_facturados.Show 1
            End If
         End If
         rsaux2.Close
      End If
   Else
      MsgBox "No se a indicado una serie para esta Unidad organizacional", vbOKOnly, "ATENCION"
      txt_numero_embarque.Enabled = False
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_factura_embarques)
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
End Sub

Private Sub lv_embarques_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_embarques, ColumnHeader)
End Sub

Private Sub lv_embarques_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txt_numero_embarque = lv_embarques.selectedItem
       frm_embarques_vivos.Visible = False
       txt_numero_embarque.SetFocus
    End If
End Sub

Private Sub lv_embarques_LostFocus()
   frm_embarques_vivos.Visible = False
End Sub

Private Sub lv_movimientos_ItemClick(ByVal item As MSComctlLib.ListItem)
   'txt_renglones = lv_movimientos.selectedItem.SubItems(7)
   'txt_de = lv_movimientos.selectedItem.SubItems(8)
   'txt_a = lv_movimientos.selectedItem.SubItems(9)
   var_almacen = lv_movimientos.selectedItem.SubItems(18)
   var_clave_movimiento = lv_movimientos.selectedItem.SubItems(1)
   var_numero_mov = lv_movimientos.selectedItem.SubItems(2)
   var_clave_moneda = lv_movimientos.selectedItem.SubItems(19)
   var_tipo_Cambio = lv_movimientos.selectedItem.SubItems(20)
End Sub


Private Sub txt_embarque_activo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Dim var_contador As Double
      Dim var_tipo As String
      lv_embarques.ListItems.Clear
      rs.Open "select * from VW_EMBARQUES_ACTIVOS where vcha_emp_empresa_id = '" + var_empresa + "' AND CHAR_MOV_DOCUMENTO = 'N'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_contador = 1
         While Not rs.EOF
            If Trim(rs!CHAR_EMB_ESTATUS) = "I" Then
               var_tipo = IIf(IsNull(rs!char_emb_tipo), "", rs!char_emb_tipo)
               Set list_item = Me.lv_envio_informacion.ListItems.Add(, , rs!inte_emb_embarque)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
               Me.lv_envio_informacion.ListItems.item(var_contador).Selected = True
               If var_tipo = "R" Then
                  Me.lv_envio_informacion.selectedItem.ForeColor = &HFF&
                  Me.lv_envio_informacion.selectedItem.ListSubItems.item(1).ForeColor = &HFF&
               End If
               var_contador = var_contador + 1
            End If
            rs.MoveNext
         Wend
         rs.Close
         If var_contador > 8 Then
            Me.lv_envio_informacion.ColumnHeaders(2).Width = 3800
         Else
            Me.lv_envio_informacion.ColumnHeaders(2).Width = 4000.25
         End If
         Me.frm_envio_informacion.Visible = True
         Me.lv_envio_informacion.SetFocus
      Else
         rs.Close
         MsgBox "No existen embarques activos", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_embarque_activo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Dim var_referencia_vi As String
   Dim var_contador_renglones As Integer
   Dim var_cadena_cajas As String
   Dim var_posible As Boolean
   Dim var_copia As Boolean
   Dim var_eliminar As Boolean
   Dim var_nombre_archivo As String
   Dim var_numero_folio_anterior As Double
   Dim var_clave_moneda As String
   Dim var_moneda_local As Integer
   Dim var_tipo_Cambio As Double
   Dim var_posible_tipo_cambio As Boolean
   Dim var_clave_movimiento_anterior As String
   Dim var_catalogo_1 As String
   Dim var_catalogo_2 As String
   Dim var_fecha_surtido_catalogo As Date
   Dim var_importe_posible_surtido As Double
   Dim var_importe_surtir As Double
   Dim var_lista_precios_catalogo As String
   Dim var_precio_catalogo_1 As Double
   Dim var_precio_catalogo_2 As Double
   Dim var_importe_disponible As Double
   Dim var_importe_catalogos As Double
   Dim var_mes_catalogo As Integer
   Dim var_año_catalogo As Integer
   
   rs.Open "select * from tb_encabezado_embarques where inte_emb_embarque = " + Str(var_numero_embarque), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_embarque_cerrado = IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", Trim(rs!CHAR_EMB_ESTATUS))
   End If
   rs.Close
   si = MsgBox("¿Desea generar el reporte del embarque?", vbYesNo, "ATENCION")
   If si = 6 Then
         var_clave_movimiento_anterior = var_clave_movimiento
         
         If Trim(var_embarque_cerrado) = "" Then
            If rsaux3.State = 1 Then
               rsaux3.Close
            End If
            rsaux3.Open "select * from vw_embarques_cerrar where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque_activo + " and char_emb_estatus = 'I'", cnn, adOpenDynamic, adLockOptimistic
            var_tipo_Cambio = 0
            var_posible_tipo_cambio = True
            While Not rsaux3.EOF
               var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
               If var_moneda_local = 0 Then
                  var_tipo_Cambio = IIf(IsNull(rsaux3!mone_tca_importe), 0, rsaux3!mone_tca_importe)
                  If var_tipo_Cambio = 0 Then
                     var_posible_tipo_cambio = False
                  End If
               End If
               rsaux3.MoveNext
            Wend
            If var_posible_tipo_cambio = True Then
               var_numero_folio_anterior = var_numero_folio
               If rsaux3.RecordCount > 0 Then
                  rsaux3.MoveFirst
               End If
               While Not rsaux3.EOF
                  var_almacen_Destino = rsaux3!VCHA_EMO_ALMACEN_DESTINO
                  var_clave_movimiento = rsaux3!VCHA_MOV_MOVIMIENTO_ID
                  var_numero_folio = rsaux3!INTE_SAL_NUMERO
                  var_clave_moneda = rsaux3!vcha_mon_moneda_id
                  var_almacen_origen = rsaux3!VCHA_ALM_ALMACEN_ID
                  var_clave_titular = IIf(IsNull(rsaux3!vcha_tit_titular_id), "", rsaux3!vcha_tit_titular_id)
                  var_clave_cliente = IIf(IsNull(rsaux3!vcha_cli_clave_id), "", rsaux3!vcha_cli_clave_id)
                  var_almacen_OS = var_almacen_origen
                  var_estatus_movimiento = rsaux3!char_Emo_estatus
                  var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
                  If var_moneda_local = 0 Then
                     var_tipo_Cambio = IIf(IsNull(rsaux3!mone_tca_importe), 0, rsaux3!mone_tca_importe)
                  Else
                     var_tipo_Cambio = 1
                  End If
                  If var_numero_folio > 0 Then
                     'Set reporte = appl.OpenReport(App.Path + "\rep_notas_envio.rpt")
                     'reporte.RecordSelectionFormula = "{VW_orden_surtido_mov.inte_emo_numero} = " + Str(var_numero_folio) + " and {VW_ORDEN_SURTIDO_MOV.FLOA_SAL_CANTIDAD} > 0 and {VW_ORDEN_SURTIDO_MOV.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "'"
                     If var_clave_movimiento = "SV" Then
                        Set reporte = appl.OpenReport(App.Path + "\rep_salidas_VISTAS.rpt")
                        reporte.RecordSelectionFormula = "{VW_SALIDAS_VISTAS.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_Destino + "' and {VW_SALIDAS_VISTAS.INTE_ENT_NUMERO} = " + Str(var_numero_folio) + " and {VW_SALIDAS_VISTAS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_SALIDAS_VISTAS.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_SALIDAS_VISTAS.vcha_uor_unidad_id} = '" + var_unidad_organizacional + "'"
                        frmvistasprevias.cr.ReportSource = reporte
                        For ntablas = 1 To reporte.Database.Tables.Count
                            reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                        Next ntablas
                        frmvistasprevias.cr.ViewReport
                        frmvistasprevias.Caption = "Reporte de Movimientos"
                        frmvistasprevias.Show 1
                        Set reporte = Nothing
                     Else
                        Set reporte = appl.OpenReport(App.Path + "\rep_nota_ENVIO_TIENDAS.rpt")
                        reporte.RecordSelectionFormula = "{VW_ENVIO_TIENDAS.inte_emo_numero} = " + Str(var_numero_folio) + " and {VW_ENVIO_TIENDAS.FLOA_SAL_CANTIDAD} > 0 and {VW_ENVIO_TIENDAS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_ENVIO_TIENDAS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "'"
                        frmvistasprevias.cr.ReportSource = reporte
                        For ntablas = 1 To reporte.Database.Tables.Count
                            reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                        Next ntablas
                        frmvistasprevias.cr.ViewReport
                        frmvistasprevias.Caption = "Reporte de Movimientos"
                        frmvistasprevias.Show 1
                        Set reporte = Nothing
                        var_nombre_archivo = ""
                        If Len(Trim(Str(var_numero_folio))) = 1 Then
                           var_nombre_archivo = "00000" + Trim(Str(var_numero_folio))
                        End If
                        If Len(Trim(Str(var_numero_folio))) = 2 Then
                           var_nombre_archivo = "0000" + Trim(Str(var_numero_folio))
                        End If
                        If Len(Trim(Str(var_numero_folio))) = 3 Then
                           var_nombre_archivo = "000" + Trim(Str(var_numero_folio))
                        End If
                        If Len(Trim(Str(var_numero_folio))) = 4 Then
                           var_nombre_archivo = "00" + Trim(Str(var_numero_folio))
                        End If
                        If Len(Trim(Str(var_numero_folio))) = 5 Then
                           var_nombre_archivo = "0" + Trim(Str(var_numero_folio))
                        End If
                        If Len(Trim(Str(var_numero_folio))) = 6 Then
                           var_nombre_archivo = Trim(Str(var_numero_folio))
                        End If
                        If Dir(App.Path & "\nota_env.dbf") <> "" Then
                           Set var_tabla = CreateObject("ADODB.connection")
                           var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + App.Path + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
                           rsaux2.Open "delete from nota_env", var_tabla, adOpenDynamic, adLockOptimistic
                           var_eliminar = DeleteFile(App.Path & "\temp_" + Trim(var_nombre_archivo) + ".dbf")
                           var_eliminar = DeleteFile(App.Path & "\" + Trim(var_nombre_archivo) + ".dbf")
                           var_copia = CopyFile(App.Path & "\nota_env.dbf", App.Path & "\temp_" + Trim(var_nombre_archivo) + ".dbf", 1)
                           Cadena = "select * from VW_ORDEN_SURTIDO_MOV where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_emo_numero = " + Str(var_numero_folio)
                           var_si = MsgBox("          ¿Enviar Correo?", vbYesNo, "ATENCION")
                           If var_si = 6 Then
                              'rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                              'var_numero_pedido_cliente = 0
                              'If Not rs.EOF Then
                              '   var_numero_pedido_cliente = IIf(IsNull(rs!INTE_PED_REFERENCIA), 0, rs!INTE_PED_REFERENCIA)
                              'Else
                              '   var_numero_pedido_cliente = 0
                              'End If
                              'rs.Close
                              var_numero_pedido_cliente = 0
                              Cadena = "select * from tb_salidas where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "'"
                              rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                              While Not rs.EOF
                                    Cadena = "insert into " + App.Path + "\temp_" + Trim(var_nombre_archivo) + ".dbf (cvenota, cvecliente, clapr, canp1, canp2, canp3, canp4, canp5, canp6, prepr, cvepedido, anocosto,tallas, talla1, talla2, talla3, talla4, talla5, talla6) values ('" + Trim(Str(var_numero_folio)) + "', '" + var_clave_cliente + "', '" + Mid(Trim(rs!vcha_Art_Articulo_id), 7, 5) + "', " + Trim(CStr(IIf(IsNull(rs!floa_Sal_Cantidad), 0, rs!floa_Sal_Cantidad))) + ", 0, 0, 0, 0, 0, " + Trim(CStr(Round(rs!floa_Sal_costo, 4))) + ", " + CStr(var_numero_pedido_cliente) + ", '" + Trim(CStr(rs!INTE_sAL_AÑO)) + "',0,0,0,0,0,0,0)"
                                    rsaux2.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
                                     rs.MoveNext
                              Wend
                              rs.Close
                              var_tabla.Close
                              var_copia = CopyFile(App.Path & "\temp_" + Trim(var_nombre_archivo) + ".dbf", App.Path & "\" + Trim(var_nombre_archivo) + ".dbf", 1)
                              var_correo_electronico = ""
                              rsaux4.Open "select * from tb_clientes where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux4.EOF Then
                                 var_correo_electronico = rsaux4!vcha_cli_email
                              End If
                              rsaux4.Close
                              If Trim(var_correo_electronico) <> "" Then
                                 If MAPISession1.SessionID = 0 Then
                                    MAPISession1.SignOn
                                 End If
                                 MAPIMessages1.SessionID = MAPISession1.SessionID
                                 MAPIMessages1.Compose
                                 MAPIMessages1.RecipDisplayName = var_correo_electronico
                                 MAPIMessages1.RecipAddress = var_correo_electronico
                                 MAPIMessages1.AddressResolveUI = True
                                 MAPIMessages1.ResolveName
                                 MAPIMessages1.MsgSubject = "Nota de envio " + Str(var_numero_folio)
                                 MAPIMessages1.MsgNoteText = "Se adjunta nota de envio número " + Str(var_numero_folio)
                                 MAPIMessages1.AttachmentPathName = App.Path + "\" + Trim(var_nombre_archivo) + ".dbf"
                                 MAPIMessages1.Send True
                                 If MAPISession1.SessionID > 0 Then
                                    MAPISession1.SignOff
                                 End If
                              Else
                                 MsgBox "El cliente no cuenta con una cuenta de correo electronico", vbOKOnly, "ATENCION"
                              End If
                           End If
                        Else
                           MsgBox "No se encuentra el archivo " + App.Path + "\nota_env.dbf, consulte con el administrador del sistema", vbOKOnly, "ATENCION"
                        End If
                     End If
                     rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux3.MoveNext
               Wend
               rsaux3.Close
            Else
               MsgBox "No se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El embarque no a sido cerrado", vbOKOnly, "ATENCION"
      End If
End If
Me.frm_embarque_envio.Visible = False
End If
If KeyAscii = 27 Then
   Me.frm_envio_informacion.Visible = False
End If

End Sub

Private Sub txt_embarque_correo_ft_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_embarque_correo_ft) Then
         rs.Open "SELECT * FROM VW_FT_FACTURACION WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque_correo_ft, cnn, adOpenDynamic, adLockOptimistic
         var_correo_electronico = IIf(IsNull(rs!VCHA_AGE_EMAIL), "", rs!VCHA_AGE_EMAIL)
         If Trim(var_correo_electronico) <> "" Then
            If MAPISession1.SessionID = 0 Then
               MAPISession1.SignOn
            End If
            MAPIMessages1.SessionID = MAPISession1.SessionID
            MAPIMessages1.Compose
            MAPIMessages1.RecipDisplayName = var_correo_electronico
            MAPIMessages1.RecipAddress = var_correo_electronico
            MAPIMessages1.AddressResolveUI = True
            MAPIMessages1.ResolveName
            MAPIMessages1.MsgSubject = "Información del pedido " + CStr(IIf(IsNull(rs!inte_ped_numero), "", rs!inte_ped_numero)) + " del cliente " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            MAPIMessages1.MsgNoteText = "Se anexa archivo con información del pedido  " + CStr(IIf(IsNull(rs!inte_ped_numero), "", rs!inte_ped_numero))
            var_Archivo = App.Path & "\Pedido_" + Trim(CStr(IIf(IsNull(rs!inte_ped_numero), "", rs!inte_ped_numero))) + ".txt"
            Open (App.Path & "\Pedido_" + Trim(CStr(IIf(IsNull(rs!inte_ped_numero), "", rs!inte_ped_numero))) + ".txt") For Output As #1
            Print #1, "Se facturo el pedido " + Trim(CStr(IIf(IsNull(rs!inte_ped_numero), "", rs!inte_ped_numero))) + " con los siguientes datos"
            Print #1, ""
            Print #1, "Cliente: " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            Print #1, ""
            rsaux8.Open "select * from tb_sellos where vcha_Emp_empresa_id = '" + var_empresa + "' and  inte_emb_embarque = " + Me.txt_embarque_correo_ft, cnn, adOpenDynamic, adLockOptimistic
            Print #1, "Guias: "
            While Not rsaux8.EOF
                  Print #1, IIf(IsNull(rsaux8!vcha_sel_Sello), "", rsaux8!vcha_sel_Sello)
                  rsaux8.MoveNext
            Wend
            rsaux8.Close
            Print #1, ""
            Print #1, "Lugar de entrega de la mercancia: "
            rsaux8.Open "SELECT * FROM VW_ESTABLECIMIENTOS_EMBARQUES WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_EMB_EMBARQUE = " + Me.txt_embarque_correo_ft, cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
               Print #1, "   Dirección: " + IIf(IsNull(rsaux8!vcha_esb_domicilio), "", rsaux8!vcha_esb_domicilio)
               Print #1, "   Colonia:   " + IIf(IsNull(rsaux8!vcha_col_nombre), "", rsaux8!vcha_col_nombre)
               Print #1, "   Ciudad:    " + IIf(IsNull(rsaux8!vcha_ciu_nombre), "", rsaux8!vcha_ciu_nombre)
               Print #1, "   Municipio: " + IIf(IsNull(rsaux8!vcha_mun_nombre), "", rsaux8!vcha_mun_nombre)
               Print #1, "   Estado:    " + IIf(IsNull(rsaux8!vcha_est_nombre), "", rsaux8!vcha_est_nombre)
               Print #1, "   Pais:      " + IIf(IsNull(rsaux8!vcha_pai_nombre), "", rsaux8!vcha_pai_nombre)
            End If
            rsaux8.Close
            var_i = 0
            var_importe_total = 0
            Print #1, ""
            Print #1, "Facturas:"
            var_moneda = CStr(rs!vcha_mon_nombre_plural)
            While Not rs.EOF
                  var_cadena = ""
                  var_importe_total = var_importe_total + rs!floa_Car_importe_neto
                  var_cadena = var_cadena + " " + CStr(IIf(IsNull(rs!inte_Car_numero), "", rs!inte_Car_numero))
                  var_importe_str = Format(CStr(rs!floa_Car_importe_neto), "###,###,##0.00")
                  For var_i = 1 + Len(Trim(var_importe_str)) To 14
                      var_importe_str = " " + var_importe_str
                  Next var_i
                  var_cadena = var_cadena + " con importe de " + var_importe_str + " " + CStr(rs!vcha_mon_nombre_plural)
                  Print #1, var_cadena
                  rs.MoveNext
            Wend
            Print #1, "=================================="
            var_importe_total_str = Format(var_importe_total, "###,###,##0.00#")
            For var_i = 1 + Len(Trim(var_importe_total_str)) To 26
                var_importe_total_str = " " + var_importe_total_str
            Next var_i
            Print #1, "Por un total de " + var_importe_total_str + " " + var_moneda
             Close #1
            MAPIMessages1.AttachmentPathName = var_Archivo
            MAPIMessages1.Send True
            If MAPISession1.SessionID > 0 Then
               MAPISession1.SignOff
            End If
         Else
            MsgBox "El cliente no cuenta con una cuenta de correo electronico", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Embarque incorrecto", vbOKOnly, "ATENCION"
      End If
      Me.frm_embarque_correo_ft.Visible = False
   End If
End Sub

Private Sub txt_embarque_correo_ft_LostFocus()
   Me.frm_embarque_correo_ft.Visible = False
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If rsaux3.State = 1 Then
         rsaux3.Close
      End If
      rsaux3.Open "select * from vw_embarques_cerrar where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_embarque + " and (char_emb_estatus = 'I' or char_emb_estatus = 'F')", cnn, adOpenDynamic, adLockOptimistic
      var_tipo_Cambio = 0
      var_posible_tipo_cambio = True
      'While Not rsaux3.EOF
      '      var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
      '      If var_moneda_local = 0 Then
      '         var_tipo_Cambio = IIf(IsNull(rsaux3!mone_tca_importe), 0, rsaux3!mone_tca_importe)
      '         If var_tipo_Cambio = 0 Then
      '            var_posible_tipo_cambio = False
      '         End If
      '      End If
      '      rsaux3.MoveNext
      'Wend
      If var_posible_tipo_cambio = True Then
         var_numero_folio_anterior = var_numero_folio
         If rsaux3.RecordCount > 0 Then
            rsaux3.MoveFirst
         End If
         While Not rsaux3.EOF
               var_clave_movimiento = rsaux3!VCHA_MOV_MOVIMIENTO_ID
               var_numero_folio = rsaux3!INTE_SAL_NUMERO
               var_clave_moneda = rsaux3!vcha_mon_moneda_id
               var_almacen_origen = rsaux3!VCHA_ALM_ALMACEN_ID
               var_clave_titular = IIf(IsNull(rsaux3!vcha_tit_titular_id), "", rsaux3!vcha_tit_titular_id)
               var_clave_cliente = IIf(IsNull(rsaux3!vcha_cli_clave_id), "", rsaux3!vcha_cli_clave_id)
               rsaux4.Open "select * from tb_clientes where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                  var_correo_electronico = rsaux4!vcha_cli_email
               End If
               rsaux4.Close
               If Trim(var_correo_electronico) <> "" Then
                  var_almacen_OS = var_almacen_origen
                  var_estatus_movimiento = rsaux3!char_Emo_estatus
                  var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
                  If var_moneda_local = 0 Then
                     var_tipo_Cambio = IIf(IsNull(rsaux3!mone_tca_importe), 0, rsaux3!mone_tca_importe)
                  Else
                     var_tipo_Cambio = 1
                  End If
                  Cadena = "SELECT max(dbo.TB_SALIDAS.INTE_CAR_NUMERO) as maximo FROM dbo.TB_SALIDAS INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_SALIDAS.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO WHERE (dbo.TB_SALIDAS.INTE_sal_NUMERO = " + Str(var_numero_folio) + ") AND (dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento
                  Cadena = Cadena + "') AND (dbo.TB_SALIDAS.vcha_Emp_empresa_id = '" + var_empresa + "')"
                  rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  var_numero_folio_2 = IIf(IsNull(rs!maximo), 0, rs!maximo)
                  rs.Close
                  If var_numero_folio_2 > 0 Then
                      var_nombre_archivo = ""
                     If Len(Trim(Str(var_numero_folio_2))) = 1 Then
                        var_nombre_archivo = "00000" + Trim(Str(var_numero_folio_2))
                     End If
                     If Len(Trim(Str(var_numero_folio_2))) = 2 Then
                        var_nombre_archivo = "0000" + Trim(Str(var_numero_folio_2))
                     End If
                     If Len(Trim(Str(var_numero_folio_2))) = 3 Then
                        var_nombre_archivo = "000" + Trim(Str(var_numero_folio_2))
                     End If
                     If Len(Trim(Str(var_numero_folio_2))) = 4 Then
                        var_nombre_archivo = "00" + Trim(Str(var_numero_folio_2))
                     End If
                     If Len(Trim(Str(var_numero_folio_2))) = 5 Then
                        var_nombre_archivo = "0" + Trim(Str(var_numero_folio_2))
                     End If
                     If Len(Trim(Str(var_numero_folio_2))) = 6 Then
                        var_nombre_archivo = Trim(Str(var_numero_folio_2))
                     End If
                     If Dir(App.Path & "\nota_env.dbf") <> "" Then
                        Set var_tabla = CreateObject("ADODB.connection")
                        var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + App.Path + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
                        rsaux2.Open "delete from nota_env", var_tabla, adOpenDynamic, adLockOptimistic
                        var_eliminar = DeleteFile(App.Path & "\temp_" + Trim(var_nombre_archivo) + ".dbf")
                        var_eliminar = DeleteFile(App.Path & "\" + Trim(var_nombre_archivo) + ".dbf")
                        var_copia = CopyFile(App.Path & "\nota_env.dbf", App.Path & "\temp_" + Trim(var_nombre_archivo) + ".dbf", 1)
                        Cadena = "select * from VW_ORDEN_SURTIDO_MOV where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_emo_numero = " + Str(var_numero_folio)
                        var_si = MsgBox("              ¿Enviar Correo?", vbYesNo, "ATENCION")
                        If var_si = 6 Then
                        
                           'rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                           'var_numero_pedido_cliente = 0
                           'If Not rs.EOF Then
                           '   var_numero_pedido_cliente = IIf(IsNull(rs!INTE_PED_REFERENCIA), 0, rs!INTE_PED_REFERENCIA)
                           'Else
                           '   var_numero_pedido_cliente = 0
                           'End If
                           'rs.Close
                           
                           

                           
                           Cadena = "SELECT dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID, dbo.TB_SALIDAS.INTE_CAR_NUMERO, dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID, dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, dbo.TB_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_ENCABEZADO_MOVIMIENTOS.FLOA_EMO_DESCUENTO_1, dbo.TB_ENCABEZADO_MOVIMIENTOS.FLOA_EMO_DESCUENTO_2, dbo.TB_ENCABEZADO_MOVIMIENTOS.FLOA_EMO_TIPO_CAMBIO, dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_SALIDAS.inte_sal_año FROM dbo.TB_SALIDAS INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_SALIDAS.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO WHERE (dbo.TB_SALIDAS.INTE_sal_NUMERO = " + Str(var_numero_folio) + ") AND (dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento
                           Cadena = Cadena + "') AND (dbo.TB_SALIDAS.vcha_Emp_empresa_id = '" + var_empresa + "') order by inte_car_numero"
                           
                           'Cadena = "select inte_sal_numero, vcha_Art_articulo_id, floa_sal_cantidad, floa_sal_precio, floa_emo_tipo_cambio from tb_salidas, tb_encabezado_movimientos where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio)
                           
                           rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                           While Not rs.EOF
                                 var_precio_articulo = IIf(IsNull(rs!floa_Sal_precio), 0, rs!floa_Sal_precio) / IIf(IsNull(rs!floa_emo_tipo_cambio), 1, rs!floa_emo_tipo_cambio)
                                 var_precio_articulo = var_precio_articulo * (1 - (IIf(IsNull(rs!floa_emo_descuento_1), 0, rs!floa_emo_descuento_1) / 100))
                                 var_precio_articulo = var_precio_articulo * (1 - (IIf(IsNull(rs!floa_emo_descuento_2), 0, rs!floa_emo_descuento_2) / 100))
                                 Cadena = "insert into " + App.Path + "\temp_" + Trim(var_nombre_archivo) + ".dbf (cvenota, cvecliente, clapr, canp1, canp2, canp3, canp4, canp5, canp6, prepr, cvepedido, anocosto, tallas, talla1, talla2, talla3, talla4, talla5, talla6) values ('" + Trim(Str(rs!inte_Car_numero)) + "', '" + var_clave_cliente + "', '" + Mid(Trim(rs!vcha_Art_Articulo_id), 7, 5) + "', " + Trim(CStr(IIf(IsNull(rs!floa_Sal_Cantidad), 0, rs!floa_Sal_Cantidad))) + ", 0, 0, 0, 0, 0, " + Trim(CStr(Round(var_precio_articulo, 4))) + ", 0, '" + Trim(CStr(rs!INTE_sAL_AÑO)) + "',0,0,0,0,0,0,0)"
                                 rsaux2.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
                                 rs.MoveNext
                           Wend
                           rs.Close
                           var_tabla.Close
                           var_copia = CopyFile(App.Path & "\temp_" + Trim(var_nombre_archivo) + ".dbf", App.Path & "\" + Trim(var_nombre_archivo) + ".dbf", 1)
                           var_correo_electronico = ""
                           rsaux4.Open "select * from tb_clientes where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux4.EOF Then
                              var_correo_electronico = rsaux4!vcha_cli_email
                           End If
                           rsaux4.Close
                           If Trim(var_correo_electronico) <> "" Then
                              If MAPISession1.SessionID = 0 Then
                                 MAPISession1.SignOn
                              End If
                              MAPIMessages1.SessionID = MAPISession1.SessionID
                              MAPIMessages1.Compose
                              MAPIMessages1.RecipDisplayName = var_correo_electronico
                              MAPIMessages1.RecipAddress = var_correo_electronico
                              MAPIMessages1.AddressResolveUI = True
                              MAPIMessages1.ResolveName
                              MAPIMessages1.MsgSubject = "Nota de envio " + Str(var_numero_folio_2)
                              MAPIMessages1.MsgNoteText = "Se adjunta nota de envio número " + Str(var_numero_folio_2)
                              MAPIMessages1.AttachmentPathName = App.Path + "\" + Trim(var_nombre_archivo) + ".dbf"
                              MAPIMessages1.Send True
                              If MAPISession1.SessionID > 0 Then
                                 MAPISession1.SignOff
                              End If
                           Else
                              MsgBox "El cliente no cuenta con una cuenta de correo electronico", vbOKOnly, "ATENCION"
                           End If
                        End If
                     Else
                        MsgBox "No se encuentra el archivo " + App.Path + "\nota_env.dbf, consulte con el administrador del sistema", vbOKOnly, "ATENCION"
                     End If
                  End If
               Else
                  MsgBox "El cliente no cuenta con una cuenta de correo electronico", vbOKOnly, "ATENCION"
               End If
               rsaux3.MoveNext
          Wend
          rsaux3.Close
       Else
           MsgBox "No se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
       End If
       Me.frm_correo.Visible = False
   End If
End Sub

Private Sub txt_embarque_LostFocus()
   frm_correo.Visible = False
End Sub

Private Sub txt_embarque_relacion_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      rs.Open "select * from vw_facturas_distintas where VCHA_EMP_EMPRESA_ID ='" + var_empresa + "' and inte_emb_embarque = " + txt_embarque_relacion, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open "select * from tb_inventario_documentos where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and VCHA_CAR_TIPO_DOCUMENTO = '" + rs!vcha_Car_tipo_documento + "' and VCHA_CAR_DOCUMENTO = '" + rs!vcha_car_documento + "' and VCHA_AGE_AGENTE_ID = '" + rs!VCHA_AGE_AGENTE_ID + "' and VCHA_CAR_CLASE_ID = '" + rs!vcha_Car_clase_id + "' and INTE_CAR_NUMERO = " + CStr(rs!inte_Car_numero) + " and  VCHA_SER_SERIE_ID = '" + rs!vcha_Ser_Serie_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If rsaux.EOF Then
               var_cadena = "INSERT INTO [TB_INVENTARIO_DOCUMENTOS] ([VCHA_EMP_EMPRESA_ID],[VCHA_AGE_AGENTE_ID],[VCHA_CAR_TIPO_DOCUMENTO], [VCHA_CAR_DOCUMENTO],  [VCHA_CAR_CLASE_ID], [INTE_CAR_NUMERO], [CHAR_CAR_AFECTACION], [VCHA_SER_SERIE_ID], [CHAR_IDO_ESTATUS], [FLOA_IDO_CANTIDAD], [FLOA_CAR_IMPORTE_NETO], [FLOA_CAR_TIPO_CAMBIO], [VCHA_MON_MONEDA_ID],[DTIM_IDO_FECHA_ENTRAGA],[VCHA_CLI_CLAVE_ID], [INTE_EMB_EMBARQUE])"
               var_cadena = var_cadena + " Values ( '" + var_empresa + "', '" + rs!VCHA_AGE_AGENTE_ID + "', '" + rs!vcha_Car_tipo_documento + "', '" + rs!vcha_car_documento + "', '" + rs!vcha_Car_clase_id + "', " + CStr(rs!inte_Car_numero) + ", '+', '" + rs!vcha_Ser_Serie_id + "',  'A', " + CStr(rs!Cantidad) + ", " + CStr(rs!floa_Car_importe_neto) + ", " + CStr(rs!floa_car_tipo_cambio) + ", '" + rs!vcha_mon_moneda_id + "', '" + CStr(Date) + "', '" + rs!vcha_cli_clave_id + "', " + txt_embarque_relacion + ")"
               rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            Else
               rsaux2.Open "update tb_inventario_documentos set FLOA_IDO_CANTIDAD = " + CStr(rs!Cantidad) + ", inte_emb_embarque = " + txt_embarque_relacion + " where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and VCHA_CAR_TIPO_DOCUMENTO = '" + rs!vcha_Car_tipo_documento + "' and VCHA_CAR_DOCUMENTO = '" + rs!vcha_car_documento + "' and VCHA_AGE_AGENTE_ID = '" + rs!VCHA_AGE_AGENTE_ID + "' and VCHA_CAR_CLASE_ID = '" + rs!vcha_Car_clase_id + "' and INTE_CAR_NUMERO = " + CStr(rs!inte_Car_numero) + " and  VCHA_SER_SERIE_ID = '" + rs!vcha_Ser_Serie_id + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            rsaux.Close
            rs.MoveNext
         Wend
         rsaux4.Open "select distinct vcha_age_agente_id from tb_inventario_documentos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_embarque_relacion, cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux4.EOF
            Set reporte = appl.OpenReport(App.Path + "\rep_inventario_documentos_embarque.rpt")
            var_cadena = "{VW_INVENTARIO_DOCUMENTOS.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_INVENTARIO_DOCUMENTOS.inte_emb_embarque} = " + txt_embarque_relacion + " and {VW_INVENTARIO_DOCUMENTOS.vcha_age_agente_id} = '" + rsaux4!VCHA_AGE_AGENTE_ID + "'"
            reporte.RecordSelectionFormula = var_cadena
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Movimientos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            
            Set reporte = appl.OpenReport(App.Path + "\rep_inventario_documentos_ENTREGA.rpt")
            var_cadena = "{VW_INTVENTARIO_DOCUMENTOS_ENTREGA.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_INTVENTARIO_DOCUMENTOS_ENTREGA.inte_emb_embarque} = " + txt_embarque_relacion + " and {VW_INTVENTARIO_DOCUMENTOS_ENTREGA.vcha_age_agente_id} = '" + rsaux4!VCHA_AGE_AGENTE_ID + "'"
            reporte.RecordSelectionFormula = var_cadena
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Movimientos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            
            
            rsaux4.MoveNext
         Wend
         rsaux4.Close
      Else
         MsgBox "No existen facturas en el embarque indicado", vbOKOnly, "ATENCION"
      End If
      rs.Close
      frm_embarque_relacion.Visible = False
   End If
End Sub

Private Sub txt_factura_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_factura) Then
      
         rsaux9.Open "select distinct VCHA_MOV_MOVIMIENTO_ID, inte_SAL_numero from tb_salidas where inte_car_numero = " + Me.txt_factura + " AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "' ", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux9.EOF Then
            var_clave_movimiento = IIf(IsNull(rsaux9!VCHA_MOV_MOVIMIENTO_ID), "", rsaux9!VCHA_MOV_MOVIMIENTO_ID)
            var_cadena = "select inte_Emb_embarque from tb_Detalle_embarqueS where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_SAL_numero = " + CStr(IIf(IsNull(rsaux9(0).Value), 0, rsaux9(0).Value))
            rsaux8.Open "select inte_Emb_embarque from tb_Detalle_embarqueS where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_SAL_numero = " + CStr(IIf(IsNull(rsaux9(1).Value), 0, rsaux9(1).Value)) + " AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
               txt_numero_embarque = IIf(IsNull(rsaux8(0).Value), 0, rsaux8(0).Value)
               If IsNumeric(Me.txt_numero_embarque) Then
                  cnn.BeginTrans
                  If rs.State = 1 Then
                     rs.Close
                  End If
                  rs.Open "select isnull(max(inte_tem_consecutivo),0) from tb_temp_factura_embarques", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_consecutivo = rs(0).Value
                  Else
                     var_consecutivo = 0
                  End If
                  rs.Close
                  var_consecutivo = var_consecutivo + 1
                  rs.Open "insert into tb_temp_factura_embarques (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  cnn.CommitTrans
                  If var_empresa = "28" Then
                     Cadena = "EXEC SP_CREA_TABLA_FACTURAS_VIANNEY_CATALOG " + CStr(var_consecutivo) + ",'" + var_empresa + "'," + txt_numero_embarque
                  Else
                     Cadena = "EXEC SP_CREA_TABLA_FACTURAS_CHIQUIBLANCOS " + CStr(var_consecutivo) + ",'" + var_empresa + "'," + txt_numero_embarque
                  End If
                  rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  Me.frm_mensaje.Visible = False
                  rsaux3.Open "select distinct inte_car_numero, vcha_ser_Serie_id, VCHA_CLI_CLAVE_ID from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " AND INTE_CAR_NUMERO  = " + Me.txt_factura + " ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     VAR_CLAVE_RETENCION = IIf(IsNull(rsaux3!vcha_cli_clave_id), "", rsaux3!vcha_cli_clave_id)
                     While Not rsaux3.EOF
                           If VAR_CLAVE_RETENCION = "C000008200" Then
                              Set reporte = appl.OpenReport(App.Path + "\rep_factura_chiquiblancos_RETENCION.rpt")
                              reporte.RecordSelectionFormula = "{TB_TEMP_FACTURA_EMBARQUES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {TB_TEMP_FACTURA_EMBARQUES.inte_car_numero} = " + CStr(rsaux3!inte_Car_numero) + " and {TB_TEMP_FACTURA_EMBARQUES.vcha_ser_serie_id} = '" + Trim(rsaux3!vcha_Ser_Serie_id) + "' and {TB_TEMP_FACTURA_EMBARQUES.vcha_emp_empresa_id} = '" + var_empresa + "'"
                              frmvistasprevias.cr.ReportSource = reporte
                              For ntablas = 1 To reporte.Database.Tables.Count
                                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), "chiquiblancos_sid", parametros(4), parametros(5)
                              Next ntablas
                              frmvistasprevias.cr.ViewReport
                              frmvistasprevias.Caption = "Reimpresion de factura " + Me.txt_factura
                              frmvistasprevias.Show 1
                           Else
                              If var_empresa = "28" Then
                                 Set reporte = appl.OpenReport(App.Path + "\rep_factura_vianney_catalog.rpt")
                              Else
                                 Set reporte = appl.OpenReport(App.Path + "\rep_factura_chiquiblancos.rpt")
                              End If
                              reporte.RecordSelectionFormula = "{TB_TEMP_FACTURA_EMBARQUES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {TB_TEMP_FACTURA_EMBARQUES.inte_car_numero} = " + CStr(rsaux3!inte_Car_numero) + " and {TB_TEMP_FACTURA_EMBARQUES.vcha_ser_serie_id} = '" + Trim(rsaux3!vcha_Ser_Serie_id) + "' and {TB_TEMP_FACTURA_EMBARQUES.vcha_emp_empresa_id} = '" + var_empresa + "'"
                              frmvistasprevias.cr.ReportSource = reporte
                              For ntablas = 1 To reporte.Database.Tables.Count
                                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), "chiquiblancos_sid", parametros(4), parametros(5)
                              Next ntablas
                              frmvistasprevias.cr.ViewReport
                              frmvistasprevias.Caption = "Reimpresion de factura " + Me.txt_factura
                              frmvistasprevias.Show 1
                           End If
                           rsaux3.MoveNext
                     Wend
                  End If
                  rsaux3.Close
                  rsaux3.Open "delete from TB_TEMP_FACTURA_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               Else
                  MsgBox "La factura no se encuentra en ningun embarque", vbOKOnly, "ATENCION"
               End If
            Else
                  MsgBox "La factura no se encuentra en ningun embarque", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "La factura no existe", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Número de factura incorrecta", vbOKOnly, "ATENCION"
      End If
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
   Me.frm_reimpresion.Visible = False
   End If
   If KeyAscii = 27 Then
      Me.frm_reimpresion.Visible = False
   End If
   Me.txt_numero_embarque = ""
End Sub

Private Sub txt_factura_nueva_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_factura_nueva) Then
         rsaux9.Open "select distinct VCHA_MOV_MOVIMIENTO_ID, inte_SAL_numero from tb_salidas where inte_car_numero = " + txt_factura_nueva, cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux9.EOF Then
            var_si = MsgBox("Deseas reimprimir la factura " + Me.txt_factura_nueva, vbYesNo, "ATENCION")
            If var_si = 6 Then
            var_clave_movimiento = IIf(IsNull(rsaux9!VCHA_MOV_MOVIMIENTO_ID), "", rsaux9!VCHA_MOV_MOVIMIENTO_ID)
            var_cadena = "select inte_Emb_embarque from tb_Detalle_embarqueS where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_SAL_numero = " + CStr(IIf(IsNull(rsaux9(0).Value), 0, rsaux9(0).Value))
            rsaux8.Open "select inte_Emb_embarque from tb_Detalle_embarqueS where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_SAL_numero = " + CStr(IIf(IsNull(rsaux9(1).Value), 0, rsaux9(1).Value)), cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
               txt_numero_embarque = IIf(IsNull(rsaux8(0).Value), 0, rsaux8(0).Value)
               If IsNumeric(Me.txt_numero_embarque) Then
               
               
                           cnn.BeginTrans
                           If rs.State = 1 Then
                              rs.Close
                           End If
                           rs.Open "select isnull(max(inte_tem_consecutivo),0) from tb_temp_factura_embarques", cnn, adOpenDynamic, adLockOptimistic
                           If Not rs.EOF Then
                              var_consecutivo = rs(0).Value
                           Else
                              var_consecutivo = 0
                           End If
                           rs.Close
                           var_consecutivo = var_consecutivo + 1
                           rs.Open "insert into tb_temp_factura_embarques (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                           cnn.CommitTrans
                           cnn.CommandTimeout = 360
                           If var_empresa = "18" Then
                              Cadena = "EXEC SP_CREA_TABLA_FACTURAS " + CStr(var_consecutivo) + ",'" + var_empresa + "'," + Me.txt_numero_embarque
                           Else
                              Cadena = "EXEC SP_CREA_TABLA_FACTURAS_nuevo " + CStr(var_consecutivo) + ",'" + var_empresa + "'," + txt_numero_embarque
                           End If
                           rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                           Me.frm_mensaje.Visible = False
                        
                           If var_empresa = "01" Or var_empresa = "02" Or var_empresa = "18" Then
                              rsaux3.Open "select distinct inte_car_numero from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_car_numero  = " + Me.txt_factura_nueva + " ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 var_Archivo = App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat"
                                 Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat") For Output As #2
                                 While Not rsaux3.EOF
                                       If rs.State = 1 Then
                                          rs.Close
                                       End If
                                       If var_empresa <> "03" Then
                                          rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' ORDER BY INTE_SAL_CONSECUTIVO_FACTURA", cnn, adOpenDynamic, adLockOptimistic
                                       Else
                                          rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' ORDER BY vcha_sal_descripcion_factura", cnn, adOpenDynamic, adLockOptimistic
                                       End If
                                       If Not rs.EOF Then
                                         'AQUI EMPIEZA LA FACTURA
                                          Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt") For Output As #1
                                          'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                                          'Print #1, Chr(27) + Chr(67) + Chr(44) + Chr(13)
                                          'Print #1, ""
                                          Print #1, Chr(15) + Chr(27) + Chr(64)
                                          Print #1, ""
                                          Print #1, ""
                                          Print #1, Spc(80); CStr(rs!inte_Car_numero)
                                          Print #1, Spc(95); Format(rs!dtim_Car_fecha, "Short Date")
                                          Print #1, ""
                                          Print #1, Spc(95); IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                                          Print #1, ""
                                          Print #1, ""
                                          'Print #1, Spc(92); Str(rs!inte_car_PLAZO) + " DIAS DE VENCIMIENTO"
                                          var_cliente = "CLIENTE:   " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                                          var_cliente_coppel = "cxcxcxcx"
                                          var_cliente_sigo = "cxcxcxccc"
                                          For var_j = 1 + Len(Trim(var_cliente)) To 70
                                              var_cliente = var_cliente + " "
                                          Next var_j
                                          var_cliente = var_cliente + " AGUASCALIENTES, AGS."
                                          Print #1, var_cliente
                                          var_domicilio = "DOMICILIO: " + Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION)) + " COLONIA: " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
                                          var_agente = ""
                                          var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                                          For var_j = 1 + Len(Trim(var_agente)) To 8
                                              var_agente = var_agente + " "
                                          Next var_j
                                          rsaux4.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux4.EOF Then
                                             var_agente = var_agente + IIf(IsNull(rsaux4!VCHA_AGE_NOMBRE), "", rsaux4!VCHA_AGE_NOMBRE)
                                          Else
                                             var_agente = var_agente + ""
                                          End If
                                          rsaux4.Close
                                          var_domicilio = var_domicilio
                                         'Print #1, Spc(111); var_agente
                                          Print #1, var_domicilio
                                          var_ciudad = ""
                                          var_ciudad = "CIUDAD:    " + IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                                          var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                                          If Trim(var_estado) <> "" Then
                                             var_ciudad = var_ciudad + ", " + var_estado
                                          Else
                                             var_ciudad = var_ciudad
                                          End If
                                          var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                          var_ciudad = var_ciudad
                                          For var_j = 1 + Len(Trim(var_ciudad)) To 70
                                             var_ciudad = var_ciudad + " "
                                          Next var_j
                                  
                                          var_ciudad = var_ciudad + " " + var_agente
                                  
                                          VAR_EMBARQUE = "EMB.: " + txt_numero_embarque
                                          var_ordern_surtido = x
                                          Print #1, var_ciudad
                                          If Trim(var_rfc) <> "" Then
                                             var_rfc = "  RFC:  " + var_rfc
                                          Else
                                             var_rfc = "  RFC:  "
                                          End If
                                          var_rfc = "C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP) + var_rfc
                                          For var_j = 1 + Len(Trim(var_rfc)) To 70
                                              var_rfc = var_rfc + " "
                                          Next var_j
                                          var_rfc = var_rfc + "PED.: " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))) + " "
                                          var_rfc = var_rfc + " O.S.: " + Trim(Str(IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), 0, rs!INTE_ORS_ORDEN_SURTIDO))) + " " + VAR_EMBARQUE
                                          Print #1, var_rfc
                                          'Print #1, Spc(10); IIf(IsNull(rs!vcha_esb_establecimiento_id), "", rs!vcha_esb_establecimiento_id)
                                          Print #1, ""
                                          Print #1, ""
                                          Print #1, ""
                                          var_importe_descuento_1 = 0
                                          var_importe_descuento_2 = 0
                                          var_importe_descuento_3 = 0
                                          var_contador_promociones = 0
                                          var_cantidad_total = 0
                                        MMM = 1  ''temporal
                                          If MMM = 1 Then ''temporal
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
                                                 var_linea = Trim(IIf(IsNull(rs!vcha_agr_agrupador_id), "", rs!vcha_agr_agrupador_id))
                                                 If Len(var_linea) = 13 Or Len(var_linea) = 12 Then
                                                    var_linea = Mid(var_linea, 9, 4)
                                                 End If
                                                 For var_j = 1 + Len(Trim(var_linea)) To 9
                                                     var_linea = var_linea + " "
                                                 Next var_j
                                                 var_cantidad_nueva = Format(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                                                 If Len(Trim(var_cantidad_nueva)) < 7 Then
                                                    For var_j = 1 + Len(Trim(var_cantidad_nueva)) To 7
                                                        var_cantidad_nueva = " " + var_cantidad_nueva
                                                    Next var_j
                                                 End If
                                                 
                                                 var_linea = var_linea + var_cantidad_nueva + "  " + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                                 var_i = 0
                                                 While Len((var_linea)) < 75
                                                       var_linea = var_linea + " "
                                                 Wend
                                                 var_linea = var_linea + " "
                                                 var_linea = var_linea + var_marca_promocion
                                                 var_cantidad = Format(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                                                 var_cantidad_total = var_cantidad_total + IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                                                 If Len(Trim(var_cantidad)) < 14 Then
                                                    For var_j = 1 + Len(Trim(var_cantidad)) To 14
                                                        var_cantidad = " " + var_cantidad
                                                    Next var_j
                                                 End If
                                                 var_precio = IIf(IsNull(rs!Importe), 0, rs!Importe)
                                                 var_descuento_1 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
                                                 var_descuento_2 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2)
                                                 var_descuento_3 = IIf(IsNull(rs!floa_car_porcentaje_descuento_3), 0, rs!floa_car_porcentaje_descuento_3)
                                                 var_porcentaje = (100 - var_descuento_1) / 100
                                                 var_precio = var_precio * var_porcentaje
                                                 var_importe_descuento_1_2 = (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                                 var_importe_descuento_1 = var_importe_descuento_1 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                                 var_precio = var_precio * ((100 - var_descuento_2) / 100)
                                                 var_importe_descuento_2 = var_importe_descuento_2 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - (var_importe_descuento_1_2 + var_precio))
                                                 var_precio = var_precio * ((100 - var_descuento_3) / 100)
                                                 var_precio = var_precio / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                                                 'var_precio_str = Format(var_precio / IIf(IsNull(rs!cantidad), 0, rs!cantidad), "###,###,##0.00")
                                                 var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                                 If Len(Trim(var_rfc)) > 0 Then
                                                    var_precio_str = Format(IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                                                 Else
                                                    var_precio_str = Format((IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) * (1 + (rs!floa_car_porcentaje_iva / 100)), "###,###,##0.00")
                                                 End If
                                                 If Len(Trim(var_precio_str)) < 14 Then
                                                    For var_j = 1 + Len(Trim(var_precio_str)) To 14
                                                       var_precio_str = " " + var_precio_str
                                                    Next var_j
                                                 End If
                                                 var_linea = var_linea + var_precio_str + "          "
                                                 If Len(Trim(var_rfc)) > 0 Then
                                                    var_importe = Format((IIf(IsNull(rs!Importe), 0, rs!Importe)), "###,###,##0.00")
                                                    If Len(Trim(var_importe)) < 14 Then
                                                        For var_j = 1 + Len(Trim(var_importe)) To 14
                                                            var_importe = " " + var_importe
                                                        Next var_j
                                                    End If
                                                 Else
                                                    var_importe = Format((IIf(IsNull(rs!Importe), 0, rs!Importe) * (1 + (rs!floa_car_porcentaje_iva / 100))), "###,###,##0.00")
                                                    If Len(Trim(var_importe)) < 14 Then
                                                       For var_j = 1 + Len(Trim(var_importe)) To 14
                                                           var_importe = " " + var_importe
                                                       Next var_j
                                                    End If
                                                 End If
                                                 var_linea = var_linea + var_importe
                                        
                                                 Print #1, var_linea
                                                 rs.MoveNext
                                              Else
                                                 Print #1, ""
                                              End If
                                              Next var_k
                                              End If ''temporal
                                              rs.MoveFirst
                                 
                                              var_cantidad_total_str = Format(var_cantidad_total, "###,###,##0.00")
                                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                              If Len(Trim(var_rfc)) > 0 Then
                                                 var_cantidad_letra = rs!vcha_car_importe_letra
                                                 var_importe_descuento_1_str = Format(IIf(IsNull(rs!floa_car_importe_descuento_1), 0, rs!floa_car_importe_descuento_1) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio), "###,###,##0.00")
                                                 If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                                    For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                                        var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                                    Next var_j
                                                 End If
                                                 var_importe_descuento_2_str = Format(IIf(IsNull(rs!floa_car_importe_descuento_2), 0, rs!floa_car_importe_descuento_2) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio), "###,###,##0.00")
                                                 If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                                    For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                                        var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                                    Next var_j
                                                 End If
                                              Else
                                                 var_cantidad_letra = rs!vcha_car_importe_letra
                                                 var_importe_descuento_1_str = Format((IIf(IsNull(rs!floa_car_importe_descuento_1), 0, rs!floa_car_importe_descuento_1)) * (1 + (rs!floa_car_porcentaje_iva / 100) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                                 If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                                    For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                                        var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                                    Next var_j
                                                 End If
                                                 var_importe_descuento_2_str = Format((IIf(IsNull(rs!floa_car_importe_descuento_2), 0, rs!floa_car_importe_descuento_2)) * (1 + (rs!floa_car_porcentaje_iva / 100) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                                 If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                                    For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                                        var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                                    Next var_j
                                                 End If
                                              End If
                                              
                                              var_cantidad_total_str = Trim(var_cantidad_total_str)
                                              If Len(Trim(var_cantidad_total_str)) < 14 Then
                                                 For var_j = 1 + Len(Trim(var_cantidad_total_str)) To 14
                                                     var_cantidad_total_str = " " + var_cantidad_total_str
                                                 Next var_j
                                              End If
                                              
                                              If Len(var_importe_descuento_1_str) >= 0 Then
                                                 var_linea = "      - DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                                              Else
                                                 var_linea = ""
                                              End If
                                              var_linea = var_cantidad_total_str + " " + var_linea
                                              For var_j = 1 + Len(var_linea) To 99
                                                  var_linea = var_linea + " "
                                              Next var_j
                                              var_linea = var_linea + var_importe_descuento_1_str
                                              Print #1, Spc(2); var_linea
                                              
                                              
                                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                              
                                              
                                              var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                                              Print #1, Spc(25); var_linea
                                              If Len(Trim(var_rfc)) = 0 Then
                                                 var_subimporte = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                                 If Len(Trim(var_subimporte)) < 14 Then
                                                    For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                                        var_subimporte = " " + var_subimporte
                                                    Next var_j
                                                 End If
                                                 var_iva = "-"
                                                 For var_j = 1 + Len(Trim(var_iva)) To 11
                                                     var_iva = " " + var_iva
                                                 Next var_j
                                              Else
                                                 var_subimporte = Format(Round(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                                 If Len(Trim(var_subimporte)) < 14 Then
                                                    For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                                        var_subimporte = " " + var_subimporte
                                                    Next var_j
                                                 End If
                                                 var_iva = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                                 If Len(Trim(var_iva)) < 14 Then
                                                    For var_j = 1 + Len(Trim(var_iva)) To 14
                                                        var_iva = " " + var_iva
                                                    Next var_j
                                                 End If
                                              End If
                                           var_importe = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                           If Len(Trim(var_importe)) < 14 Then
                                              For var_j = 1 + Len(Trim(var_importe)) To 14
                                                  var_importe = " " + var_importe
                                              Next var_j
                                           End If
                                           var_linea = var_importe
                                           Print #1, Spc(101); var_subimporte
                                           Print #1, Spc(67); var_linea + "                    " + var_iva
                                           Print #1, Spc(101); var_linea
                                           var_linea = ""
                                           MMM = 1
                                           If MMM = 1 Then ''' THENPORAL
                                           Print #1, ""
                                           Print #1, ""
                                           Print #1, ""
                                           Print #1, ""
                                           Print #1, ""
                                           Print #1, ""
                                           Print #1, ""
                                           Print #1, ""
                                           End If
                                           Close #1
                                           If Trim(var_empresa) = "02" Then
                                              Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt lpt1"
                                           Else
                                              Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt lpt1"
                                           End If
                                          'AQUI TERMINA LA FACTURA
                                       End If
                                       rs.Close
                                       rsaux3.MoveNext
                                 Wend
                                 Close #2
                                 x = Shell(var_Archivo, vbHide)
                              End If
                              rsaux3.Close
                              'Aqui se termina de imprimir la factura
                           End If
                           rsaux3.Open "delete from TB_TEMP_FACTURA_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                           rsaux3.Open "DELETE FROM TB_TEMP_SALIDAS_FACTURACION WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                           MsgBox "Se a terminado el proceso de facturación", vbOKOnly, "ATENCION"

               Else
                  MsgBox "La factura no se encuentra en ningun embarque", vbOKOnly, "ATENCION"
               End If
            End If
            Else
                  MsgBox "La factura no se encuentra en ningun embarque", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "La factura no existe", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Número de factura incorrecta", vbOKOnly, "ATENCION"
      End If
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
      Me.frm_reimpresion_nueva.Visible = False
   End If
   If KeyAscii = 27 Then
      Me.frm_reimpresion.Visible = False
   End If
End Sub

Private Sub txt_factura_nueva_LostFocus()
   Me.frm_reimpresion_nueva.Visible = False
End Sub

Private Sub txt_numero_embarque_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Dim var_contador As Double
      Dim var_tipo As String
      lv_embarques.ListItems.Clear
      rs.Open "select * from VW_EMBARQUES_ACTIVOS where vcha_emp_empresa_id = '" + var_empresa + "' AND CHAR_MOV_DOCUMENTO = 'F'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_contador = 1
         While Not rs.EOF
            If Trim(rs!CHAR_EMB_ESTATUS) = "I" Then
               var_tipo = IIf(IsNull(rs!char_emb_tipo), "", rs!char_emb_tipo)
               Set list_item = lv_embarques.ListItems.Add(, , rs!inte_emb_embarque)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
               lv_embarques.ListItems.item(var_contador).Selected = True
               If var_tipo = "R" Then
                  lv_embarques.selectedItem.ForeColor = &HFF&
                  lv_embarques.selectedItem.ListSubItems.item(1).ForeColor = &HFF&
               End If
               var_contador = var_contador + 1
            End If
            rs.MoveNext
         Wend
         rs.Close
         If var_contador > 8 Then
            lv_embarques.ColumnHeaders(2).Width = 3800
         Else
            lv_embarques.ColumnHeaders(2).Width = 4000.25
         End If
         frm_embarques_vivos.Visible = True
         lv_embarques.SetFocus
      Else
         rs.Close
         MsgBox "No existen embarques activos", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_numero_embarque_KeyPress(KeyAscii As Integer)
   Dim list_item As ListItem
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      var_total = 0
      txt_agente = ""
      txt_clave_agente = ""
      txt_fecha = ""
      txt_jaula = ""
      txt_de = ""
      txt_a = ""
      txt_renglones = ""
      txt_importe = ""
      txt_piezas = ""
      lv_movimientos.ListItems.Clear
      If Trim(txt_numero_embarque) <> "" Then
         rsaux4.Open "select * from VW_EMBARQUES_ACTIVOS_2 where vcha_emp_empresa_id = '" + var_empresa + "' AND CHAR_MOV_DOCUMENTO = 'F' and inte_emb_embarque = " + txt_numero_embarque, cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux4.EOF Then
            rs.Open "select * from tb_encabezado_embarques where inte_emb_embarque = " + txt_numero_embarque + " and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_estatus_embarque = IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", rs!CHAR_EMB_ESTATUS)
               If Trim(rs!CHAR_EMB_ESTATUS) = "I" Then
                  var_total_facturas = 0
                  rsaux2.Open "Select * from tb_agentes where vcha_age_agente_id ='" + rs!VCHA_AGE_AGENTE_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_agente = rsaux2!VCHA_AGE_NOMBRE
                  txt_clave_agente = rsaux2!VCHA_AGE_AGENTE_ID
                  txt_jaula = rs!inte_jau_jaula_id
                  txt_fecha = Date
                  var_total_importe = 0
                  var_total_piezas = 0
                  var_numero_renglones = 0
                  rsaux2.Close
                  rs.Close
                  lv_movimientos.ListItems.Clear
                  rs.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_factura_inicio = rs!inte_ser_factura
                  var_total_de = var_factura_inicio
                  var_total_a = var_factura_inicio
                  rs.Close
                  rs.Open "SELECT * FROM VW_EMBARQUES_ACTIVOS_2 WHERE INTE_EMB_EMBARQUE = " + txt_numero_embarque + " and vcha_mov_movimiento_id <> 'AV' AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_total_facturas = 0
                     While Not rs.EOF
                        rsaux2.Open "Select * from vw_datos_factura where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           var_plazo = IIf(IsNull(rsaux2!inte_pla_dias), 0, rsaux2!inte_pla_dias)
                           var_agrupador = IIf(IsNull(rsaux2!VCHA_FAG_FAMILIA_AGRUPADOR_ID), "", rsaux2!VCHA_FAG_FAMILIA_AGRUPADOR_ID)
                           var_iva = IIf(IsNull(rsaux2!FLOA_TPE_IVA), 0, rsaux2!FLOA_TPE_IVA)
                        Else
                           var_plazo = 0
                           var_agrupador = ""
                           var_iva = 0
                        End If
                        rsaux2.Close
                        Set list_item = lv_movimientos.ListItems.Add(, , rs!inte_emo_numero_origen)
                        list_item.SubItems(1) = IIf(IsNull(rs!VCHA_MOV_MOVIMIENTO_ID), "", rs!VCHA_MOV_MOVIMIENTO_ID)
                        var_clave_mov = IIf(IsNull(rs!VCHA_MOV_MOVIMIENTO_ID), "", rs!VCHA_MOV_MOVIMIENTO_ID)
                        list_item.SubItems(2) = IIf(IsNull(rs!INTE_SAL_NUMERO), 0, rs!INTE_SAL_NUMERO)
                        var_numero_mov = IIf(IsNull(rs!INTE_SAL_NUMERO), 0, rs!INTE_SAL_NUMERO)
                        
                        rsaux3.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           list_item.SubItems(4) = IIf(IsNull(rsaux3!VCHA_CLI_NOMBRE), "", rsaux3!VCHA_CLI_NOMBRE)
                        End If
                        rsaux3.Close
                        rsaux3.Open "SELECT * FROM TB_ESTABLECIMIENTOS WHERE VCHA_ESB_ESTABLECIMIENTO_ID = '" + rs!vcha_ESB_ESTABLECIMIENTO_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           list_item.SubItems(3) = IIf(IsNull(rsaux3!VCHA_ESB_NOMBRE), "", rsaux3!VCHA_ESB_NOMBRE)
                        End If
                        rsaux3.Close
                        
                        Call facturas
                        list_item.SubItems(5) = Format(var_piezas, "###,###,##0.00")
                        list_item.SubItems(6) = Format(var_total * (1 + (var_iva / 100)), "###,###,##0.00")
                        list_item.SubItems(8) = var_total_de
                        list_item.SubItems(7) = (var_total_a + 1) - var_total_de
                        var_total_facturas = var_total_facturas + ((var_total_a + 1) - var_total_de)
                        list_item.SubItems(9) = var_total_a
                        var_total_de = var_total_de + ((var_total_a + 1) - var_total_de)
                        list_item.SubItems(10) = var_subimporte
                        list_item.SubItems(11) = var_imp_total_desc_1
                        list_item.SubItems(12) = var_imp_total_desc_2
                        list_item.SubItems(13) = 0
                        list_item.SubItems(14) = var_iva
                        list_item.SubItems(15) = var_total * (var_iva / 100)
                        list_item.SubItems(16) = var_plazo
                        list_item.SubItems(17) = var_agrupador
                        list_item.SubItems(18) = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
                        list_item.SubItems(19) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                        var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                        list_item.SubItems(20) = IIf(IsNull(rs!floa_emo_tipo_cambio), 1, rs!floa_emo_tipo_cambio)
                        var_tipo_Cambio = IIf(IsNull(rs!floa_emo_tipo_cambio), 1, rs!floa_emo_tipo_cambio)
                        var_factura_inicio = var_factura_inicio + (0 + list_item.SubItems(7))
                        var_factura_inicio = var_factura_inicio + Val(txt_renglones)
                        var_total_importe = var_total_importe + (var_total + (var_total * var_iva / 100))
                        var_total_piezas = var_total_piezas + var_piezas
                        rs.MoveNext
                     Wend
                     txt_piezas = Format(var_total_piezas, "###,###,##0.00")
                     txt_importe = Format(var_total_importe, "###,###,##0.00")
                     txt_renglones = lv_movimientos.selectedItem.SubItems(7)
                     txt_de = lv_movimientos.selectedItem.SubItems(8)
                     txt_a = lv_movimientos.selectedItem.SubItems(9)
                     var_almacen = lv_movimientos.selectedItem.SubItems(18)
                     var_clave_movimiento = lv_movimientos.selectedItem.SubItems(1)
                     var_numero_mov = lv_movimientos.selectedItem.SubItems(2)
                     lv_movimientos.SetFocus
                  Else
                     MsgBox "El embarque no tiene movimientos asignados", vbOKOnly, "ATENCION"
                  End If
               Else
                  If Trim(rs!CHAR_EMB_ESTATUS) = "F" Then
                     MsgBox "El embarque ya fue facturado", vbOKOnly, "ATENCION"
                  Else
                     MsgBox "El embarque no a sido cerrado aun", vbOKOnly, "ATENCION"
                  End If
               End If
               rs.Close
            Else
               rs.Close
               MsgBox "El número de embarque no existe", vbOKOnly, "ATENCION"
               txt_agente = ""
               txt_clave_agente = ""
               lv_movimientos.ListItems.Clear
            End If
         Else
            MsgBox "Número de embarque no existe", vbOKOnly, "ATENCION"
         End If
         rsaux4.Close
      End If
   End If
End Sub
Private Sub facturas()
'On Error GoTo SALIR:
   rsaux2.Open "select * from vw_sumatoria_salidas_total where vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_mov + "' and inte_emo_numero = " + Str(var_numero_mov) + " and floa_sal_cantidad > 0", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux2.EOF Then
      var_subtotal = 0
      var_total = 0
      var_piezas = 0
      var_descuento = 0
      var_numero_renglones = 0
      var_imp_total_desc_1 = 0
      var_imp_total_desc_2 = 0
      var_cantidad = 0
      var_precio = 0
      var_subimporte = 0
      var_importe_iva = 0
      var_total_a = var_total_de
      var_imp_descuento_1 = 0
      var_imp_descuento_2 = 0
      var_imp_descuento_3 = 0
      While Not rsaux2.EOF
         var_numero_renglones = var_numero_renglones + 1
         If var_numero_renglones > var_renglones_factura Then
            var_numero_renglones = 1
            var_total_a = var_total_a + 1
         End If
         
         var_cantidad = IIf(IsNull(rsaux2!floa_Sal_Cantidad), 0, rsaux2!floa_Sal_Cantidad)
         var_precio = ((IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio)) / IIf(IsNull(rsaux2!floa_Sal_Cantidad), 0, rsaux2!floa_Sal_Cantidad)) / (IIf(IsNull(floa_emo_tipo_cambio), 1, rsaux2!floa_emo_tipo_cambio))
         var_subtotal = var_subtotal + (var_cantidad * var_precio)
         var_subimporte = var_subimporte + (var_cantidad * var_precio)
         var_descuento_1 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_1), 0, rsaux2!FLOA_SAL_DESCUENTO_1)
         var_descuento_2 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_2), 0, rsaux2!FLOA_SAL_DESCUENTO_2)
         var_piezas = var_piezas + IIf(IsNull(rsaux2!floa_Sal_Cantidad), 0, rsaux2!floa_Sal_Cantidad)
         
         
         If var_descuento_1 > 0 Then
            var_imp_descuento_1 = (var_precio - (var_precio * ((100 - var_descuento_1) / 100)))
            var_precio_desc_1 = var_precio - var_imp_descuento_1
            var_imp_total_desc_1 = var_imp_total_desc_1 + (var_imp_descuento_1 * var_cantidad)
         End If
         If var_descuento_2 > 0 Then
            var_imp_descuento_2 = (var_precio_desc_1 - (var_precio_desc_1 * ((100 - var_descuento_2) / 100)))
            var_imp_total_desc_2 = var_imp_total_desc_2 + (var_imp_descuento_2 * var_cantidad)
         End If
         var_descuento = var_descuento + (var_cantidad * (var_imp_descuento_1 + var_imp_descuento_2))
         rsaux2.MoveNext
      Wend
      var_total = var_subtotal - var_descuento
   End If
   rsaux2.Close
salir:

End Sub

Private Sub imprimir_facturas()
On Error GoTo salir:
   Dim list_item As ListItem
   Dim var_descuento_1 As Double
   Dim var_descuento_2 As Double
   Dim var_descuento_3 As Double
   Dim var_imp_descuento_1 As Double
   Dim var_imp_descuento_2 As Double
   Dim var_imp_neto As Double
   Dim var_precio_desc_1 As Double
   Dim var_cantidad As Double
   Dim var_precio As Double
   lv_detalle.ListItems.Clear
   rs.Open "select * from vw_orden_surtido_mov where vcha_mov_movimiento_id = '" + lv_movimientos.selectedItem.SubItems(2) + "' and inte_emo_numero = " + Str(lv_movimientos.selectedItem.SubItems(3)) + " and floa_sal_cantidad > 0", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_subtotal = 0
      var_total = 0
      var_descuento = 0
      While Not rs.EOF
         'If rs!vcha_mov_movimiento_id <> "AV" Then
            Set list_item = lv_detalle.ListItems.Add(, , rs!vcha_Art_Articulo_id)
                 list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
                 list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_Sal_Cantidad), 0, rs!floa_Sal_Cantidad), "###,###,##0.00")
                 var_cantidad = IIf(IsNull(rs!floa_Sal_Cantidad), 0, rs!floa_Sal_Cantidad)
                 list_item.SubItems(3) = Format(IIf(IsNull(rs!floa_Sal_precio), 0, rs!floa_Sal_precio), "###,###,##0.00")
                 var_precio = IIf(IsNull(rs!floa_Sal_precio), 0, rs!floa_Sal_precio)
                 var_subtotal = var_subtotal + (var_cantidad * var_precio)
                 list_item.SubItems(4) = Format(IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE), "###,###,##0.00")
                 list_item.SubItems(3) = Format(IIf(IsNull(rs!floa_Sal_precio), 0, rs!floa_Sal_precio), "###,###,##0.00")
                 var_descuento_1 = IIf(IsNull(rs!FLOA_ORS_DESCUENTO_1), 0, rs!FLOA_ORS_DESCUENTO_1)
                 var_descuento_2 = IIf(IsNull(rs!FLOA_ORS_DESCUENTO_2), 0, rs!FLOA_ORS_DESCUENTO_2)
                 If var_descuento_1 > 0 Then
                    var_imp_descuento_1 = (var_precio - (var_precio * ((100 - var_descuento_1) / 100)))
                    var_precio_desc_1 = var_precio - var_imp_descuento_1
                 End If
                 If var_descuento_2 > 0 Then
                    var_imp_descuento_2 = (var_precio_desc_1 - (var_precio_desc_1 * ((100 - var_descuento_2) / 100)))
                 End If
                 
                 list_item.SubItems(4) = Str(var_descuento_1) + "% + " + Str(var_descuento_2) + "%"
                 list_item.SubItems(5) = Format(var_imp_descuento_1 + var_imp_descuento_2, "###,###,##0.00")
                 var_descuento = var_descuento + (var_cantidad * (var_imp_descuento_1 + var_imp_descuento_2))
                 list_item.SubItems(6) = Format(var_cantidad * (var_precio - (var_imp_descuento_1 + var_imp_descuento_2)), "###,###,##0.00")
                 
         'End If
         rs.MoveNext
      Wend
      txt_subtotal = Format(var_subtotal, "###,###,##0.00")
      txt_descuento = Format(var_descuento, "###,###,##0.00")
      var_total = var_subtotal - var_descuento
      txt_total = Format(var_total, "###,###,##0.00")
   End If
   rs.Close
salir:

End Sub

Private Sub txt_numero_embarque_LostFocus()
   x = 0
   If x = 1 Then
   Dim list_item As ListItem
   var_total = 0
   txt_agente = ""
   txt_clave_agente = ""
   txt_fecha = ""
   txt_jaula = ""
   txt_de = ""
   txt_a = ""
   txt_renglones = ""
   txt_importe = ""
   txt_piezas = ""
   lv_movimientos.ListItems.Clear
   If Trim(txt_numero_embarque) <> "" Then
      rsaux4.Open "select * from VW_EMBARQUES_ACTIVOS_2 where vcha_emp_empresa_id = '" + var_empresa + "' AND CHAR_MOV_DOCUMENTO = 'F' and inte_emb_embarque = " + txt_numero_embarque, cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux4.EOF Then
         If rsaux4!char_mov_documento = "F" Then
            rs.Open "select * from tb_encabezado_embarques where inte_emb_embarque = " + txt_numero_embarque + " and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_estatus_embarque = IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", rs!CHAR_EMB_ESTATUS)
               If Trim(rs!CHAR_EMB_ESTATUS) = "I" Then
                  var_total_facturas = 0
                  rsaux2.Open "Select * from tb_agentes where vcha_age_agente_id ='" + rs!VCHA_AGE_AGENTE_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_agente = rsaux2!VCHA_AGE_NOMBRE
                  txt_clave_agente = rsaux2!VCHA_AGE_AGENTE_ID
                  txt_jaula = rs!inte_jau_jaula_id
                  txt_fecha = Date
                  var_total_importe = 0
                  var_total_piezas = 0
                  var_numero_renglones = 0
                  rsaux2.Close
                  rs.Close
                  lv_movimientos.ListItems.Clear
                  rs.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_factura_inicio = rs!inte_ser_factura
                  var_total_de = var_factura_inicio
                  var_total_a = var_factura_inicio
                  rs.Close
                  rs.Open "SELECT * FROM VW_EMBARQUES_ACTIVOS_2 WHERE INTE_EMB_EMBARQUE = " + txt_numero_embarque + " and vcha_mov_movimiento_id <> 'AV'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_total_facturas = 0
                     While Not rs.EOF
                        rsaux2.Open "Select * from vw_datos_factura where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           var_plazo = IIf(IsNull(rsaux2!inte_pla_dias), 0, rsaux2!inte_pla_dias)
                           var_agrupador = IIf(IsNull(rsaux2!VCHA_FAG_FAMILIA_AGRUPADOR_ID), "", rsaux2!VCHA_FAG_FAMILIA_AGRUPADOR_ID)
                           var_iva = IIf(IsNull(rsaux2!FLOA_TPE_IVA), 0, rsaux2!FLOA_TPE_IVA)
                        Else
                           var_plazo = 0
                           var_agrupador = ""
                           var_iva = 0
                        End If
                        rsaux2.Close
                        
                        Set list_item = lv_movimientos.ListItems.Add(, , rs!EMO_NUMERO_ORIGEN)
                        list_item.SubItems(1) = IIf(IsNull(rs!VCHA_MOV_MOVIMIENTO_ID), "", rs!VCHA_MOV_MOVIMIENTO_ID)
                        var_clave_mov = IIf(IsNull(rs!VCHA_MOV_MOVIMIENTO_ID), "", rs!VCHA_MOV_MOVIMIENTO_ID)
                        list_item.SubItems(2) = IIf(IsNull(rs!INTE_SAL_NUMERO), 0, rs!INTE_SAL_NUMERO)
                        var_numero_mov = IIf(IsNull(rs!INTE_SAL_NUMERO), 0, rs!INTE_SAL_NUMERO)
                        rsaux4.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux4.EOF Then
                           list_item.SubItems(4) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                        End If
                        rsaux4.Close
                        rsaux4.Open "SELECT * FROM TB_ESTABLECIMIENTOS WHERE VCHA_ESB_ESTABLECIMIENTO_ID = '" + rs!vcha_ESB_ESTABLECIMIENTO_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux4.EOF Then
                           list_item.SubItems(3) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
                        End If
                        rsaux4.Close
                        Call facturas
                        list_item.SubItems(5) = Format(var_piezas, "###,###,##0.00")
                        list_item.SubItems(6) = Format(var_total * (1 + (var_iva / 100)), "###,###,##0.00")
                        list_item.SubItems(8) = var_total_de
                        list_item.SubItems(7) = (var_total_a + 1) - var_total_de
                        var_total_facturas = var_total_facturas + ((var_total_a + 1) - var_total_de)
                        list_item.SubItems(9) = var_total_a
                        var_total_de = var_total_de + ((var_total_a + 1) - var_total_de)
                        list_item.SubItems(10) = var_subimporte
                        list_item.SubItems(11) = var_imp_total_desc_1
                        list_item.SubItems(12) = var_imp_total_desc_2
                        list_item.SubItems(13) = 0
                        list_item.SubItems(14) = var_iva
                        list_item.SubItems(15) = var_total * (var_iva / 100)
                        list_item.SubItems(16) = var_plazo
                        list_item.SubItems(17) = var_agrupador
                        list_item.SubItems(18) = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
                        list_item.SubItems(19) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                        var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                        list_item.SubItems(20) = IIf(IsNull(rs!floa_emo_tipo_cambio), 1, rs!floa_emo_tipo_cambio)
                        var_tipo_Cambio = IIf(IsNull(rs!floa_emo_tipo_cambio), 1, rs!floa_emo_tipo_cambio)
                        var_factura_inicio = var_factura_inicio + (0 + list_item.SubItems(7))
                        var_factura_inicio = var_factura_inicio + Val(txt_renglones)
                        var_total_importe = var_total_importe + (var_total + (var_total * var_iva / 100))
                        var_total_piezas = var_total_piezas + var_piezas
                        rs.MoveNext
                     Wend
                     txt_piezas = Format(var_total_piezas, "###,###,##0.00")
                     txt_importe = Format(var_total_importe, "###,###,##0.00")
                     txt_renglones = lv_movimientos.selectedItem.SubItems(7)
                     txt_de = lv_movimientos.selectedItem.SubItems(8)
                     txt_a = lv_movimientos.selectedItem.SubItems(9)
                     var_almacen = lv_movimientos.selectedItem.SubItems(18)
                     var_clave_movimiento = lv_movimientos.selectedItem.SubItems(1)
                     var_numero_mov = lv_movimientos.selectedItem.SubItems(2)
                     var_contador_facturas = 0
                     If lv_movimientos.ListItems.Count > 0 Then
                        lv_movimientos.ListItems.item(1).Selected = True
                        txt_de = lv_movimientos.selectedItem.SubItems(8)
                        For var_i = 1 To lv_movimientos.ListItems.Count
                            lv_movimientos.ListItems.item(var_i).Selected = True
                            var_contador_facturas = var_contador_facturas + CInt(lv_movimientos.selectedItem.SubItems(7))
                        Next var_i
                        txt_a = lv_movimientos.selectedItem.SubItems(9)
                        txt_renglones = var_contador_facturas
                     Else
                        txt_a = ""
                        txt_de = ""
                        txt_renglones = ""
                     End If
                     lv_movimientos.SetFocus
                  Else
                     MsgBox "El embarque no tiene movimientos asignados", vbOKOnly, "ATENCION"
                  End If
               Else
                  If Trim(rs!CHAR_EMB_ESTATUS) = "F" Then
                     MsgBox "El embarque ya fue facturado", vbOKOnly, "ATENCION"
                  Else
                     MsgBox "El embarque no a sido cerrado aun", vbOKOnly, "ATENCION"
                  End If
               End If
               rs.Close
            Else
               rs.Close
               MsgBox "El número de embarque no existe", vbOKOnly, "ATENCION"
               txt_agente = ""
               txt_clave_agente = ""
               lv_movimientos.ListItems.Clear
            End If
         End If
      Else
         MsgBox "El movimiento para el que se hiso el embarque no imprime facturas", vbOKOnly, "ATENCION"
      End If
   Else
       MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
   End If
   rsaux4.Close
   End If
End Sub


