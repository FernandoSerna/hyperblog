VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmfactura_embarques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturacion de Embarques"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   Icon            =   "frmfactura_embarques.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   Begin VB.CommandButton cmd_remision 
      Caption         =   "Remision"
      Height          =   330
      Left            =   8565
      TabIndex        =   70
      Top             =   45
      Width           =   1425
   End
   Begin VB.TextBox txt_embarque_remision 
      Height          =   345
      Left            =   10035
      TabIndex        =   69
      Top             =   30
      Width           =   1125
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   330
      Left            =   6420
      TabIndex        =   68
      Top             =   60
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmd_factura_electronica 
      Caption         =   "Factura electrónica"
      Height          =   315
      Left            =   480
      TabIndex        =   67
      Top             =   45
      Width           =   1680
   End
   Begin VB.Frame frm_correo_clientes 
      Height          =   810
      Left            =   2040
      TabIndex        =   64
      Top             =   420
      Width           =   2115
      Begin VB.TextBox txt_embarque_correo_clientes 
         Height          =   315
         Left            =   75
         TabIndex        =   65
         Top             =   405
         Width           =   1920
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         Caption         =   "Embarque"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   66
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.CommandButton cmd_correo_clientes 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4140
      Picture         =   "frmfactura_embarques.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Enviar información de mercancia facturada a clientes"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton Command7 
      Caption         =   "factura anterior chiquiblancos"
      Height          =   330
      Left            =   7470
      TabIndex        =   62
      Top             =   75
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Frame frm_embarque_reimprimir 
      Height          =   810
      Left            =   1845
      TabIndex        =   59
      Top             =   465
      Width           =   2115
      Begin VB.TextBox txt_embarque_reimprimir 
         Height          =   315
         Left            =   90
         TabIndex        =   60
         Top             =   465
         Width           =   1920
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000D&
         Caption         =   "Embarque"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   61
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2160
      Picture         =   "frmfactura_embarques.frx":028E
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Reimprimir facturas"
      Top             =   45
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame frm_embarque_correo_ft 
      Height          =   810
      Left            =   1620
      TabIndex        =   55
      Top             =   495
      Width           =   2115
      Begin VB.TextBox txt_embarque_correo_ft 
         Height          =   315
         Left            =   90
         TabIndex        =   56
         Top             =   435
         Width           =   1920
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         Caption         =   "Embarque"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   57
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.CommandButton cmd_correo_facturacion_tiendas 
      Caption         =   "FT"
      Height          =   315
      Left            =   3810
      TabIndex        =   54
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_embarques_cerrados 
      Height          =   315
      Left            =   3465
      Picture         =   "frmfactura_embarques.frx":0390
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Embarques cerrados no facturados"
      Top             =   45
      Width           =   345
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   315
      Left            =   1110
      TabIndex        =   52
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   330
      Left            =   7095
      TabIndex        =   51
      Top             =   60
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   7080
      TabIndex        =   50
      Top             =   60
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   315
      Left            =   6630
      TabIndex        =   49
      Top             =   75
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame frm_embarque_envio 
      Height          =   810
      Left            =   1440
      TabIndex        =   43
      Top             =   480
      Width           =   2115
      Begin VB.TextBox txt_embarque_activo 
         Height          =   315
         Left            =   90
         TabIndex        =   44
         Top             =   420
         Width           =   1920
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         Caption         =   "Embarque"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   45
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.Frame frm_envio_informacion 
      Height          =   2745
      Left            =   2790
      TabIndex        =   40
      Top             =   495
      Width           =   5490
      Begin MSComctlLib.ListView lv_envio_informacion 
         Height          =   2235
         Left            =   30
         TabIndex        =   41
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
         TabIndex        =   42
         Top             =   45
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmd_nota_envio 
      Height          =   315
      Left            =   2475
      Picture         =   "frmfactura_embarques.frx":0492
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Generar Nota de Envio y Correo"
      Top             =   45
      Width           =   345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   4260
      TabIndex        =   39
      Top             =   45
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame frm_embarque_relacion 
      Height          =   810
      Left            =   1260
      TabIndex        =   36
      Top             =   480
      Width           =   2115
      Begin VB.TextBox txt_embarque_relacion 
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   405
         Width           =   1920
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         Caption         =   " Embarque para relación"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   38
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.CommandButton cmd_relacion_facturas 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2820
      Picture         =   "frmfactura_embarques.frx":0594
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Relación de Facturas"
      Top             =   45
      Width           =   315
   End
   Begin VB.Frame frm_correo 
      Height          =   810
      Left            =   1140
      TabIndex        =   32
      Top             =   465
      Width           =   2115
      Begin VB.TextBox txt_embarque 
         Height          =   315
         Left            =   90
         TabIndex        =   34
         Top             =   420
         Width           =   1920
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         Caption         =   "Embarque"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   33
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.CommandButton cmd_correo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3135
      Picture         =   "frmfactura_embarques.frx":0696
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Enviar Información"
      Top             =   45
      Width           =   315
   End
   Begin VB.Frame frm_embarques_vivos 
      Height          =   2745
      Left            =   2940
      TabIndex        =   29
      Top             =   555
      Width           =   5490
      Begin MSComctlLib.ListView lv_embarques 
         Height          =   2235
         Left            =   30
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   120
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11220
      Picture         =   "frmfactura_embarques.frx":0798
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmfactura_embarques.frx":0DD2
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   45
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Caption         =   " Facturas Sugeridas "
      Height          =   1095
      Left            =   6390
      TabIndex        =   21
      Top             =   525
      Width           =   5205
      Begin VB.ComboBox cmb_series 
         Height          =   315
         Left            =   615
         TabIndex        =   11
         Top             =   285
         Width           =   795
      End
      Begin VB.TextBox txt_renglones 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4575
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   615
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txt_a 
         Height          =   315
         Left            =   4575
         TabIndex        =   13
         Top             =   285
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txt_de 
         Height          =   315
         Left            =   2970
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   285
         Width           =   1920
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   135
         TabIndex        =   28
         Top             =   345
         Width           =   405
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Total de Facturas a Imprimir:"
         Height          =   195
         Left            =   1440
         TabIndex        =   24
         Top             =   675
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   195
         Left            =   3510
         TabIndex        =   23
         Top             =   345
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Factura a Imprimir:"
         Height          =   195
         Left            =   1635
         TabIndex        =   22
         Top             =   345
         Width           =   1290
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   120
      TabIndex        =   18
      Top             =   345
      Width           =   11460
   End
   Begin VB.Frame Frame2 
      Caption         =   " Movimientos a Facturar"
      Height          =   5550
      Left            =   135
      TabIndex        =   17
      Top             =   1665
      Width           =   11445
      Begin VB.Frame frm_mensaje 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   1740
         Left            =   1545
         TabIndex        =   46
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
            Left            =   555
            TabIndex        =   48
            Top             =   915
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
            TabIndex        =   47
            Top             =   255
            Width           =   8355
         End
      End
      Begin VB.TextBox txt_piezas 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   6015
         Width           =   1005
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10485
         TabIndex        =   25
         Top             =   6015
         Width           =   1065
      End
      Begin MSComctlLib.ListView lv_movimientos 
         Height          =   5190
         Left            =   60
         TabIndex        =   6
         Top             =   135
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
         TabIndex        =   27
         Top             =   6075
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Embarque "
      Height          =   1095
      Left            =   135
      TabIndex        =   0
      Top             =   525
      Width           =   6210
      Begin VB.TextBox txt_clave_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   885
         TabIndex        =   35
         Top             =   630
         Width           =   1230
      End
      Begin VB.TextBox txt_jaula 
         Height          =   315
         Left            =   4950
         TabIndex        =   9
         Top             =   285
         Width           =   1155
      End
      Begin VB.TextBox txt_fecha 
         Height          =   315
         Left            =   2760
         TabIndex        =   8
         Top             =   285
         Width           =   1620
      End
      Begin VB.TextBox txt_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2115
         TabIndex        =   10
         Top             =   630
         Width           =   3990
      End
      Begin VB.TextBox txt_numero_embarque 
         Height          =   315
         Left            =   885
         TabIndex        =   1
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Jaula:"
         Height          =   195
         Left            =   4515
         TabIndex        =   20
         Top             =   345
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   2235
         TabIndex        =   19
         Top             =   345
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   195
         TabIndex        =   16
         Top             =   675
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   345
         Width           =   600
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5850
      Top             =   -45
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
            Picture         =   "frmfactura_embarques.frx":0ED4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   5295
      Top             =   -30
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
            Picture         =   "frmfactura_embarques.frx":0FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques.frx":18C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   4740
      Top             =   -30
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
            Picture         =   "frmfactura_embarques.frx":219A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques.frx":2A74
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques.frx":334E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques.frx":38EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques.frx":41C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques.frx":4AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques.frx":537A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques.frx":548C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques.frx":559E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques.frx":56B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques.frx":57C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques.frx":58D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_embarques.frx":5A56
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   2805
      Top             =   0
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
      Left            =   3405
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "frmfactura_embarques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_consecutivo As Double
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

Private Sub crea_factura_electronica()
                                       If rs.State = 1 Then
                                          rs.Close
                                       End If
                                       If var_empresa <> "03" Then
                                          rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' ORDER BY INTE_SAL_CONSECUTIVO_FACTURA", cnn, adOpenDynamic, adLockOptimistic
                                       Else
                                          rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' ORDER BY vcha_sal_descripcion_factura", cnn, adOpenDynamic, adLockOptimistic
                                       End If
                                       If Not rs.EOF Then
                                          
                                          rsaux5.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'  ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
                                          rsaux4.Open "select * from tb_encabezado_movimientos where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + rsaux5!VCHA_MOV_MOVIMIENTO_ID + "' and inte_Emo_numero = " + CStr(rsaux5!inte_emo_numero), cnn, adOpenDynamic, adLockOptimistic
                                          var_clasificacion_maquila = ""
                                          If Not rsaux4.EOF Then
                                             var_clasificacion_maquila = IIf(IsNull(rsaux4!vcha_Emo_clasificacion), "", rsaux4!vcha_Emo_clasificacion)
                                          Else
                                             var_clasificacion_maquila = ""
                                          End If
                                          rsaux4.Close
                                          rsaux5.Close
                                          
                                          
                                          
                                          Open (App.Path & "\renombra" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat") For Output As #2
                                          'MsgBox var_ruta_documentos_electronicos
                                          Print #2, "ren " + var_ruta_documentos_electronicos + "\" + Trim(var_serie) + Trim(Str(rsaux3!inte_Car_numero)) + ".fi " + Trim(var_serie) + Trim(Str(rsaux3!inte_Car_numero)) + ".ff"
                                          Close #2
                                          'Close #1
                                          Open (var_ruta_documentos_electronicos & "\" + Trim(var_serie) + Trim(Str(rsaux3!inte_Car_numero)) + ".fi") For Output As #1
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
                                          var_cadena_fecha = var_año + "-" + var_mes + "-" + var_dia + "T" + var_hora + ":" + var_minuto + ":" + var_segundo
                                          If rsaux1.State = 1 Then
                                             rsaux1.Close
                                          End If
                                          rsaux1.Open "select * from tb_plazos where vcha_pla_plazo_id = '" + rs!VCHA_PLA_PLAZO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                          
                                          
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
                                          
                                          
                                          var_cadena = var_cadena + "fecha=" + var_cadena_fecha + Chr(13)
                                          var_cadena = var_cadena + "noAprobacion=" + Chr(13)
                                          var_cadena = var_cadena + "anoAprobacion=" + Chr(13)
                                          var_cadena = var_cadena + "tipoDeComprobante=FACTURA" + Chr(13)
                                          var_cadena = var_cadena + "formaDePago=PAGO HECHO EN UNA SOLA EXHIBICION" + Chr(13)
                                          If rsaux1.EOF Then
                                             var_cadena = var_cadena + "condicionesDePago=CONTADO" + Chr(13)
                                          Else
                                             If var_rfc_cliente = "XAXX010101000" Then
                                                var_cadena = var_cadena + "condicionesDePago=0 DIAS" + Chr(13)
                                             Else
                                                var_cadena = var_cadena + "condicionesDePago=" + CStr(IIf(IsNull(rs!INTE_CAR_PLAZO), 0, rs!INTE_CAR_PLAZO)) + " DIAS" + Chr(13)
                                             End If
                                          End If
                                          rsaux1.Close
                                          var_importe_total = rs!FLOA_CAR_IMPORTE_TOTAL / rs!floa_car_tipo_cambio
                                          
                                          If var_rfc_cliente = "XAXX010101000" Then
                                             If rs!floa_Car_importe_neto = 0 Then
                                                var_cadena = var_cadena + "subtotal=" + Format(CStr((0.01 / rs!floa_car_tipo_cambio)), "###,###,###,##0.000000") + Chr(13)
                                             Else
                                                If Trim(rs!vcha_cli_clave_id) = "C000010568" Or Trim(rs!vcha_cli_clave_id) = "C000008200" Or Trim(rs!vcha_cli_clave_id) = "C000008200" Or Trim(rs!vcha_cli_clave_id) = "C000012251" Or Trim(rs!vcha_cli_clave_id) = "C000011510" Or Trim(rs!vcha_cli_clave_id) = "C000012528" Or Trim(rs!vcha_cli_clave_id) = "C000012558" Or Trim(rs!vcha_cli_clave_id) = "C000007909" Then
                                                   var_cadena = var_cadena + "subtotal=" + Format(CStr((rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio)), "###,###,###,##0.000000") + Chr(13)
                                                Else
                                                   'var_cadena = var_cadena + "subtotal=" + Format(CStr((rs!floa_Car_importe_neto / rs!FLOA_cAR_TIPO_CAMBIO)), "###,###,###,##0.000000") + Chr(13)
                                                   var_cadena = var_cadena + "subtotal=" + Format(CStr((var_importe_total * 1.16)), "###,###,###,##0.000000") + Chr(13)
                                                End If
                                             End If
                                          Else
                                             If rs!floa_Car_importe_neto = 0 Then
                                                var_cadena = var_cadena + "subtotal=" + Format(CStr((0.01 / rs!floa_car_tipo_cambio) / (1 + (rs!floa_car_porcentaje_iva / 100))), "###,###,###,##0.000000") + Chr(13)
                                             Else
                                                If Trim(rs!vcha_cli_clave_id) = "C000010568" Or Trim(rs!vcha_cli_clave_id) = "C000008200" Or Trim(rs!vcha_cli_clave_id) = "C000008200" Or Trim(rs!vcha_cli_clave_id) = "C000012251" Or Trim(rs!vcha_cli_clave_id) = "C000011510" Or Trim(rs!vcha_cli_clave_id) = "C000012528" Or Trim(rs!vcha_cli_clave_id) = "C000012558" Or Trim(rs!vcha_cli_clave_id) = "C000007909" Then
                                                   var_cadena = var_cadena + "subtotal=" + Format(CStr((rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio)), "###,###,###,##0.000000") + Chr(13)
                                                Else
                                                   'var_cadena = var_cadena + "subtotal=" + Format(CStr((rs!floa_Car_importe_neto / rs!FLOA_cAR_TIPO_CAMBIO) / (1 + (rs!floa_car_porcentaje_iva / 100))), "###,###,###,##0.000000") + Chr(13)
                                                   var_cadena = var_cadena + "subtotal=" + Format(CStr((var_importe_total)), "###,###,###,##0.000000") + Chr(13)
                                                End If
                                             End If
                                          End If
                                          var_importe_descuento = ((rs!floa_car_importe_descuento_1 / rs!floa_car_tipo_cambio) + (rs!floa_car_importe_descuento_2 / rs!floa_car_tipo_cambio))
                                          If var_rfc_cliente = "XAXX010101000" Then
                                             var_cadena = var_cadena + "descuento=" + Format(var_importe_descuento * 1.16, "###,###,##0.00000") + Chr(13)
                                          Else
                                             var_cadena = var_cadena + "descuento=" + Format(var_importe_descuento, "###,###,##0.00000") + Chr(13)
                                          End If
                                          var_No_Descuento = 0
                                          If rsaux1.State = 1 Then
                                             rsaux1.Close
                                          End If
                                          If var_empresa = "02" Or var_empresa = "03" Then
                                             rsaux1.Open "select inte_orc_liberada from tb_encabezado_pedidos where inte_ped_numero  = " + CStr(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero)), cnn, adOpenDynamic, adLockOptimistic
                                             If Not rsaux1.EOF Then
                                                var_No_Descuento = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
                                             Else
                                                var_No_Descuento = 0
                                             End If
                                             rsaux1.Close
                                          Else
                                             var_No_Descuento = 0
                                          End If
                                          var_si_promosion = 0
                                          While Not rs.EOF
                                                If rs!floa_sal_promocion_1 > 0 Then
                                                   var_si_promosion = 1
                                                End If
                                                rs.MoveNext
                                          Wend
                                          rs.MoveFirst
                                          If var_No_Descuento = 1 Then
                                             If var_rfc_cliente = "XAXX010101000" Then
                                                var_cadena = var_cadena + "descuento1=" + Chr(13)
                                                var_cadena = var_cadena + "descuento2=" + Chr(13)
                                             Else
                                                var_cadena = var_cadena + "descuento1=" + Chr(13)
                                                var_cadena = var_cadena + "descuento2=" + Chr(13)
                                             End If
                                             var_cadena = var_cadena + "conceptodescuento1=" + "Premios a colaboradores" + Chr(13)
                                             var_cadena = var_cadena + "conceptodescuento2=" + Chr(13)
                                             var_cadena = var_cadena + "tasadescuento1=" + Chr(13)
                                             var_cadena = var_cadena + "tasadescuento2=" + Chr(13)
                                          Else
                                             If var_rfc_cliente = "XAXX010101000" Then
                                                var_cadena = var_cadena + "descuento1=" + Format(CStr((rs!floa_car_importe_descuento_1 / rs!floa_car_tipo_cambio) * (1 + (rs!floa_car_porcentaje_iva / 100))), "###,###,###,##0.000000") + Chr(13)
                                                var_cadena = var_cadena + "descuento2=" + Format(CStr((rs!floa_car_importe_descuento_2) / rs!floa_car_tipo_cambio) * (1 + (rs!floa_car_porcentaje_iva / 100)), "###,###,###,##0.000000") + Chr(13)
                                             Else
                                                var_cadena = var_cadena + "descuento1=" + Format(CStr((rs!floa_car_importe_descuento_1 / rs!floa_car_tipo_cambio)), "###,###,###,##0.000000") + Chr(13)
                                                var_cadena = var_cadena + "descuento2=" + Format(CStr((rs!floa_car_importe_descuento_2) / rs!floa_car_tipo_cambio), "###,###,###,##0.000000") + Chr(13)
                                             End If
                                             var_cadena = var_cadena + "conceptodescuento1=DESCUENTO DEL " + Chr(13)
                                             If var_si_promosion = 1 Then
                                                var_cadena = var_cadena + "conceptodescuento2=" + var_cadena_promocion_171209 + Chr(13)
                                             Else
                                                var_cadena = var_cadena + "conceptodescuento2=DESCUENTO POR PAGO OPORTUNO " + Chr(13)
                                             End If
                                             
                                             var_cadena = var_cadena + "tasadescuento1=" + CStr(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)) + "%" + Chr(13)
                                             If var_si_promosion = 1 Then
                                                var_cadena = var_cadena + "tasadescuento2=" + Chr(13)
                                             Else
                                                var_cadena = var_cadena + "tasadescuento2=" + CStr(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2)) + "%" + Chr(13)
                                             End If
                                          End If
                                          If rsaux1.State = 1 Then
                                             rsaux1.Close
                                          End If
                                          'rsaux1.Open "select * from tb_empresa_FACTURA_ELECTRONICA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                                          If rsaux1.State = 1 Then
                                             rsaux1.Close
                                          End If
                                          rsaux1.Open "select * from tb_empresa_FACTURA_ELECTRONICA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                                          var_certificado = rsaux1!vcha_emp_certificado
                                          var_expedido = rsaux1!vcha_emp_expedido
                                          If var_rfc_cliente = "XAXX010101000" Then
                                             'If rs!vcha_cli_clave_id = "C000005472" Then
                                             '   var_cadena = var_cadena + "iva=" + Format(CStr(rs!floa_car_importe_iva / rs!floa_car_tipo_cambio), "###,###,###,##0.000000") + Chr(13)
                                             'Else
                                                var_cadena = var_cadena + "iva=" + Format(CStr(0), "###,###,###,##0.000000") + Chr(13)
                                             'End If
                                          Else
                                             If rs!floa_car_importe_iva = 0 Then
                                                var_cadena = var_cadena + "iva=" + Format(CStr(0#), "###,###,###,##0.000000") + Chr(13)
                                             Else
                                                var_cadena = var_cadena + "iva=" + Format(CStr(rs!floa_car_importe_iva / rs!floa_car_tipo_cambio), "###,###,###,##0.000000") + Chr(13)
                                             End If
                                          End If
                                          If rs!floa_Car_importe_neto = 0 Then
                                             var_cadena = var_cadena + "total=" + Format(CStr(0.01), "###,###,###,##0.000000") + Chr(13)
                                          Else
                                             var_cadena = var_cadena + "total=" + Format(CStr((rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio)), "###,###,###,##0.000000") + Chr(13)
                                          End If
                                          If Trim(rs!vcha_cli_clave_id) = "C000010568" Or Trim(rs!vcha_cli_clave_id) = "C000008200" Or Trim(rs!vcha_cli_clave_id) = "C000008200" Or Trim(rs!vcha_cli_clave_id) = "C000012251" Or Trim(rs!vcha_cli_clave_id) = "C000011510" Or Trim(rs!vcha_cli_clave_id) = "C000012528" Or Trim(rs!vcha_cli_clave_id) = "C000012558" Or Trim(rs!vcha_cli_clave_id) = "C000007909" Then
                                             var_cadena = var_cadena + "retencion=" + Format(CStr(rs!floa_car_importe_iva / rs!floa_car_tipo_cambio), "###,###,###,##0.000000") + Chr(13)
                                             var_cadena = var_cadena + "factorretencioniva=16%" + Chr(13)
                                          Else
                                             var_cadena = var_cadena + "retencion=" + Chr(13)
                                             var_cadena = var_cadena + "factorretencioniva=" + Chr(13)
                                          End If
                                          var_cadena = var_cadena + "</Comprobante>" + Chr(13) + Chr(13)
                                          var_cadena = var_cadena + "<Emisor>" + Chr(13)
                                          var_cadena = var_cadena + "erfc=" + rsaux1!VCHA_eMP_RFC + Chr(13)
                                          var_cadena = var_cadena + "enombre=" + rsaux1!VCHA_EMP_NOMBRE + Chr(13)
                                          var_cadena = var_cadena + "</Emisor>" + Chr(13) + Chr(13)
                                          var_cadena = var_cadena + "<DomicilioFiscal>" + Chr(13)
                                          var_cadena = var_cadena + "ecalle=" + rsaux1!VCHA_eMP_CALLE + Chr(13)
                                          var_cadena = var_cadena + "enoExterior=" + rsaux1!VCHA_eMP_exterior + Chr(13)
                                          var_cadena = var_cadena + "enoInterior=" + Chr(13)
                                          var_cadena = var_cadena + "ecolonia=" + rsaux1!VCHA_eMP_COLONIA + Chr(13)
                                          var_cadena = var_cadena + "elocalidad=" + rsaux1!VCHA_EMP_LOCALIDAD + Chr(13)
                                          var_cadena = var_cadena + "ereferencia=" + Chr(13)
                                          var_cadena = var_cadena + "emunicipio=" + rsaux1!VCHA_EMP_MUNICIPIO + Chr(13)
                                          var_cadena = var_cadena + "eestado=" + rsaux1!VCHA_EMP_ESTADO + Chr(13)
                                          var_cadena = var_cadena + "epais=" + rsaux1!VCHA_eMP_PAIS + Chr(13)
                                          var_cadena = var_cadena + "ecodigoPostal=" + rsaux1!VCHA_EMP_CODIGO_POSTAL + Chr(13)
                                          var_cadena = var_cadena + "etel=" + IIf(IsNull(rsaux1!VCHA_EMP_TELEFONO), "", rsaux1!VCHA_EMP_TELEFONO) + Chr(13)
                                          var_cadena = var_cadena + "eemail=" + IIf(IsNull(rsaux1!VCHA_EMP_EMAIL), "", rsaux1!VCHA_EMP_EMAIL) + Chr(13)
                                          correo = IIf(IsNull(rsaux1!VCHA_EMP_EMAIL), "", rsaux1!VCHA_EMP_EMAIL)
                                          'MsgBox var_cadena
                                          var_cadena = var_cadena + "</DomicilioFiscal>" + Chr(13) + Chr(13)
                                          If var_empresa = "02" Or var_empresa = "03" Then
                                             var_cadena = var_cadena + "<ExpedidoEn>" + Chr(13) + Chr(13)
                                             var_cadena = var_cadena + "ex_calle=Blvd. Jose Maria Chavez" + Chr(13)
                                             var_cadena = var_cadena + "ex_noExterior=2202" + Chr(13)
                                             var_cadena = var_cadena + "ex_noInterior=" + Chr(13)
                                             var_cadena = var_cadena + "ex_colonia=Ciudad Industrial" + Chr(13)
                                             var_cadena = var_cadena + "ex_localidad=AGUASCALIENTES" + Chr(13)
                                             var_cadena = var_cadena + "ex_referencia=" + Chr(13)
                                             var_cadena = var_cadena + "ex_municipio=AGUASCALIENTES" + Chr(13)
                                             var_cadena = var_cadena + "ex_estado=AGUASCALIENTES" + Chr(13)
                                             var_cadena = var_cadena + "ex_pais=MEXICO" + Chr(13)
                                             var_cadena = var_cadena + "ex_codigoPostal=20290" + Chr(13)
                                             var_cadena = var_cadena + "</ExpedidoEn>"
                                          Else
                                             var_cadena = var_cadena + "<ExpedidoEn>" + Chr(13) + Chr(13)
                                             var_cadena = var_cadena + "ex_calle=" + rsaux1!VCHA_eMP_CALLE + Chr(13)
                                             var_cadena = var_cadena + "ex_noExterior=" + rsaux1!VCHA_eMP_exterior + Chr(13)
                                             var_cadena = var_cadena + "ex_noInterior=" + Chr(13)
                                             var_cadena = var_cadena + "ex_colonia=" + rsaux1!VCHA_eMP_COLONIA + Chr(13)
                                             var_cadena = var_cadena + "ex_localidad=" + rsaux1!VCHA_EMP_LOCALIDAD + Chr(13)
                                             var_cadena = var_cadena + "ex_referencia=" + Chr(13)
                                             var_cadena = var_cadena + "ex_municipio=" + rsaux1!VCHA_EMP_MUNICIPIO + Chr(13)
                                             var_cadena = var_cadena + "ex_estado=" + rsaux1!VCHA_EMP_ESTADO + Chr(13)
                                             var_cadena = var_cadena + "ex_pais=" + rsaux1!VCHA_eMP_PAIS + Chr(13)
                                             var_cadena = var_cadena + "ex_codigoPostal=" + rsaux1!VCHA_EMP_CODIGO_POSTAL + Chr(13)
                                             var_cadena = var_cadena + "</ExpedidoEn>"
                                          End If
                                          
                                          
                                          
                                          
                                          var_cadena = var_cadena + "<Receptor>" + Chr(13)
                                          If rs!vcha_cli_clave_id = "C000001204" Or rs!vcha_cli_clave_id = "C000010618" Or rs!vcha_cli_clave_id = "C000001461" Or rs!vcha_cli_clave_id = "C000002295" Or rs!vcha_cli_clave_id = "C000012359" Or rs!vcha_cli_clave_id = "C000010618" Or rs!vcha_cli_clave_id = "C000002758" Then
                                             var_cadena = var_cadena + "noCliente=" + rs!vcha_cli_clave_id + " " + IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC) + Chr(13)
                                          Else
                                             var_cadena = var_cadena + "noCliente=" + rs!vcha_cli_clave_id + Chr(13)
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
                                          var_cadena = var_cadena + "localidad=" + IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + Chr(13)
                                          rsaux1.Close
                                          rsaux1.Open "select * from vw_clientes where vcha_Cli_clave_id = '" + rs!vcha_cli_clave_id + "'"
                                          var_cadena = var_cadena + "referencia=" + Chr(13)
                                          var_cadena = var_cadena + "municipio=" + IIf(IsNull(rsaux1!vcha_mun_nombre), "", rsaux1!vcha_mun_nombre) + Chr(13)
                                          var_cadena = var_cadena + "estado=" + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre) + Chr(13)
                                          VAR_PAIS_ELECTRONICO = IIf(IsNull(rs!vcha_pai_nombre), "MEXICO", rs!vcha_pai_nombre)
                                          If Trim(VAR_PAIS_ELECTRONICO) = "" Then
                                             VAR_PAIS_ELECTRONICO = "MEXICO"
                                          End If
                                          var_cadena = var_cadena + "pais=" + VAR_PAIS_ELECTRONICO + Chr(13)
                                          var_cadena = var_cadena + Chr(13)
                                          var_cadena = var_cadena + "codigoPostal=" + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP) + Chr(13)
                                          var_cadena = var_cadena + "tel=" + Chr(13)
                                          rsaux11.Open "select * from vw_clientes where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                          'If var_empresa = "31" Or var_empresa = "16" Or var_empresa = "30" Or var_empresa = "15" Then
                                          var_cadena = var_cadena + "email=" + IIf(IsNull(rsaux11!vcha_cli_email), "", rsaux11!vcha_cli_email) + Chr(13)
                                          'Else
                                          '   If IIf(IsNull(rsaux11!VCHA_AGE_AGENTE_ID), "", rsaux11!VCHA_AGE_AGENTE_ID) <> "00100" Then
                                          '      var_cadena = var_cadena + "email=" + Chr(13)
                                          '   Else
                                          '      var_cadena = var_cadena + "email=" + IIf(IsNull(rsaux11!vcha_cli_email), "", rsaux11!vcha_cli_email) + Chr(13)
                                          '   End If
                                          '   rsaux11.Close
                                          'End If
                                          'var_cadena = var_cadena + "email=" + iif(isnull(rs!vcha_cli_email
                                          var_cadena = var_cadena + "</Cliente>" + Chr(13) + Chr(13)
                                          rsaux11.Close
                                          var_cadena = var_cadena + "<EntregarEn>" + Chr(13)
                                          rsaux4.Open "select * from vw_establecimientos_direcciones where vcha_esb_establecimiento_id = '" + rs!vcha_ESB_ESTABLECIMIENTO_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux4.EOF Then
                                             If var_rfc_cliente = "XAXX010101000" Then
                                                var_cadena = var_cadena + "endomicilio=" + IIf(IsNull(rsaux4!vcha_ESB_ESTABLECIMIENTO_id), "", rsaux4!vcha_ESB_ESTABLECIMIENTO_id) + Chr(13)
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
                                             Else
                                                var_cadena = var_cadena + "endomicilio=" + IIf(IsNull(rsaux4!vcha_ESB_ESTABLECIMIENTO_id), "", rsaux4!vcha_ESB_ESTABLECIMIENTO_id) + IIf(IsNull(rsaux4!vcha_esb_domicilio), "", rsaux4!vcha_esb_domicilio) + Chr(13)
                                                var_cadena = var_cadena + "encalle=" + Chr(13)
                                                var_cadena = var_cadena + "ennoExterior=" + Chr(13)
                                                var_cadena = var_cadena + "ennoInterior=" + Chr(13)
                                                var_cadena = var_cadena + "encolonia=" + IIf(IsNull(rsaux4!vcha_col_nombre), "", rsaux4!vcha_col_nombre) + Chr(13)
                                                var_cadena = var_cadena + "enlocalidad=" + IIf(IsNull(rsaux4!vcha_ciu_nombre), "", rsaux4!vcha_ciu_nombre) + Chr(13)
                                                var_cadena = var_cadena + "enreferencia=" + Chr(13)
                                                var_cadena = var_cadena + "enmunicipio=" + IIf(IsNull(rsaux4!vcha_mun_nombre), "", rsaux4!vcha_mun_nombre) + Chr(13)
                                                var_cadena = var_cadena + "enestado=" + IIf(IsNull(rsaux4!vcha_est_nombre), "", rsaux4!vcha_est_nombre) + Chr(13)
                                                var_cadena = var_cadena + "enpais=" + IIf(IsNull(rsaux4!vcha_pai_nombre), "", rsaux4!vcha_pai_nombre) + Chr(13)
                                                var_cadena = var_cadena + "encodigoPostal=" + Chr(13)
                                                var_cadena = var_cadena + "entel=" + Chr(13)
                                                var_cadena = var_cadena + "enemail=" + Chr(13)
                                                var_cadena = var_cadena + "</EntregarEn>" + Chr(13) + Chr(13)
                                             End If
                                          Else
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
                                          End If
                                          rsaux4.Close
                                          
                                          
                                          
                                          
                                          
                                          var_cadena = var_cadena + "<Concepto>" + Chr(13)
                                          rsaux1.Close
                                          'MsgBox var_cadena

                                          var_k = 0
                                          var_piezas_totales = 0
                                          var_si_promosion = 0
                                          While Not rs.EOF
                                                var_k = var_k + 1
                                                pxx = CStr(var_k)
                                                If Len(pxx) = 1 Then
                                                   pxx = "0" + pxx
                                                End If
                                                var_piezas_totales = var_piezas_totales + IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                                                var_cadena = var_cadena + "p" + pxx + "_cantidad=" + CStr(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) + Chr(13)
                                                rsaux1.Open "SELECT dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID, dbo.TB_UNIDADES.VCHA_UNI_UNIDAD_ID, dbo.TB_UNIDADES.VCHA_UNI_NOMBRE FROM dbo.TB_ARTICULOS LEFT OUTER JOIN dbo.TB_UNIDADES ON dbo.TB_ARTICULOS.VCHA_UNI_UNIDAD_ID = dbo.TB_UNIDADES.VCHA_UNI_UNIDAD_ID WHERE (dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID = '" + rs!vcha_agr_agrupador_id + "')"
                                                If Not rsaux1.EOF Then
                                                   var_Unidad_medida = IIf(IsNull(rsaux1!VCHA_UNI_NOMBRE), "PZA", rsaux1!VCHA_UNI_NOMBRE)
                                                   If Len(Trim(var_Unidad_medida)) = 0 Then
                                                      var_Unidad_medida = "PZA"
                                                   End If
                                                   var_cadena = var_cadena + "p" + pxx + "_unidad=" + var_Unidad_medida + Chr(13)
                                                Else
                                                   var_cadena = var_cadena + "p" + pxx + "_unidad=PZA" + Chr(13)
                                                End If
                                                rsaux1.Close
                                                var_cadena = var_cadena + "p" + pxx + "_noIdentificacion=" + IIf(IsNull(rs!vcha_agr_agrupador_id), "", rs!vcha_agr_agrupador_id) + Chr(13)
                                                If var_empresa = "15" Then
                                                   If var_clasificacion_maquila = "" Or var_clasificacion_maquila = "PRIMERA" Then
                                                      var_linea = IIf(IsNull(rs!vcha_agr_agrupador_id), "", rs!vcha_agr_agrupador_id) + " MAQUILA DE " + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                                   Else
                                                      var_linea = IIf(IsNull(rs!vcha_agr_agrupador_id), "", rs!vcha_agr_agrupador_id) + " MAQUILA DE SEGUNDA DE " + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                                   End If
                                                Else
                                                   var_linea = IIf(IsNull(rs!vcha_agr_agrupador_id), "", rs!vcha_agr_agrupador_id) + " " + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                                End If
                                                
                                                If rs!floa_sal_promocion_1 > 0 Then
                                                   var_si_promosion = 1
                                                   var_linea = var_linea + "  *"
                                                End If
                                                If var_empresa = "16" Then
                                                   If var_si_promosion = 1 Then
                                                      var_cadena = var_cadena + "p" + pxx + "_descripcion=" + var_linea + " " + CStr(rs!floa_sal_promocion_1) + "%" + Chr(13)
                                                   Else
                                                      var_cadena = var_cadena + "p" + pxx + "_descripcion=" + var_linea + Chr(13)
                                                   End If
                                                Else
                                                   var_cadena = var_cadena + "p" + pxx + "_descripcion=" + var_linea + Chr(13)
                                                End If
                                                If var_rfc_cliente = "XAXX010101000" Then
                                                   var_precio = IIf(IsNull(rs!Importe), 0, rs!Importe) * (1 + (rs!floa_car_porcentaje_iva / 100))
                                                Else
                                                   var_precio = IIf(IsNull(rs!Importe), 0, rs!Importe)
                                                End If
                                                If IIf(IsNull(rs!vcha_agr_agrupador_id), "", rs!vcha_agr_agrupador_id) <> "---" Then
                                                   'var_descuento_1 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
                                                   'var_descuento_2 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2)
                                                   'var_descuento_3 = IIf(IsNull(rs!floa_car_porcentaje_descuento_3), 0, rs!floa_car_porcentaje_descuento_3)
                                                   'var_porcentaje = (100 - var_descuento_1) / 100
                                                   'var_precio = var_precio * var_porcentaje
                                                   'var_importe_descuento_1_2 = (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                                   'var_importe_descuento_1 = var_importe_descuento_1 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                                   'var_precio = var_precio * ((100 - var_descuento_2) / 100)
                                                   'var_importe_descuento_2 = var_importe_descuento_2 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - (var_importe_descuento_1_2 + var_precio))
                                                   'var_precio = var_precio * ((100 - var_descuento_3) / 100)
                                                End If
                                                'var_precio = var_precio / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                                                If var_precio = 0 Then
                                                   var_precio = 0.00001
                                                End If
                                                var_cadena = var_cadena + "p" + pxx + "_valorUnitario=" + Format(CStr(var_precio / CStr(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad))), "###,###,###,##0.000000") + Chr(13)
                                                var_cadena = var_cadena + "p" + pxx + "_importe=" + Format(CStr(var_precio), "###,###,###,##0.000000") + Chr(13)
                                                
                                                rs.MoveNext
                                          Wend
                                                                                    

                                          var_cadena = var_cadena + "</Concepto>" + Chr(13) + Chr(13)
                                          var_cadena = var_cadena + "<Otros>" + Chr(13)
                                          var_cadena = var_cadena + "certificado=" + IIf(IsNull(var_certificado), "", var_certificado) + Chr(13)
                                          rs.MoveFirst

                                          var_cadena = var_cadena + "cant_letra=" + rs!vcha_car_importe_letra + Chr(13)
                                          var_cadena = var_cadena + "factoriva=" + CStr(rs!floa_car_porcentaje_iva) + "%" + Chr(13)
                                          var_cadena = var_cadena + "moneda=" + IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural) + Chr(13)
                                          var_cadena = var_cadena + "tipodeCambio=" + Format(CStr(IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,###,##0.000000") + Chr(13)
                                          var_cliente_coppel = rs!vcha_cli_clave_id
                                          If Trim(var_cliente_coppel) = "C000002947" Or Trim(var_cliente_coppel) = "C000005566" Or Trim(var_cliente_coppel) = "C000005831" Or var_empresa = "06" Or var_empresa = "17" Then
                                             rsaux5.Open "select * from tb_encabezado_pedidos where inte_ped_numero = " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))), cnn, adOpenDynamic, adLockOptimistic
                                             If Not rsaux5.EOF Then
                                                If Trim(var_cliente_coppel) = "C000005831" Then
                                                   var_numero_pedido = Trim(CStr(IIf(IsNull(rsaux5!VCHA_PED_PEDIDO_EXTERNO), "", rsaux5!VCHA_PED_PEDIDO_EXTERNO)))
                                                Else
                                                   If var_empresa = "06" Or var_empresa = "17" Then
                                                      var_numero_pedido = Trim(CStr(IIf(IsNull(rsaux5!VCHA_PED_PEDIDO_EXTERNO), "", rsaux5!VCHA_PED_PEDIDO_EXTERNO)))
                                                   Else
                                                      var_numero_pedido = Trim(CStr(IIf(IsNull(rsaux5!VCHA_PED_PEDIDO_EXTERNO), "", rsaux5!VCHA_PED_PEDIDO_EXTERNO)))
                                                   End If
                                                End If
                                             End If
                                             rsaux5.Close
                                          Else
                                             var_numero_pedido = Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero)))
                                          End If
                                          'MsgBox var_cadena
                                          If rs!VCHA_MOV_MOVIMIENTO_ID = "FV" Then
                                             rsaux1.Open "select * from tb_remisiones where inte_emo_numero  = " + CStr(rs!inte_emo_numero), cnn, adOpenDynamic, adLockOptimistic
                                             If rsaux1.EOF Then
                                                var_cadena = var_cadena + "pedido=" + Chr(13)
                                             Else
                                                var_cadena = var_cadena + "pedido=" + Trim(IIf(IsNull(rsaux1!VCHA_REM_REMISION_AGENTE), "", rsaux1!VCHA_REM_REMISION_AGENTE)) + Chr(13)
                                             End If
                                             rsaux1.Close
                                          Else
                                             var_cadena = var_cadena + "pedido=" + Trim(var_numero_pedido) + Chr(13)
                                          End If
                                          var_cadena = var_cadena + "Embarque=" + Me.txt_numero_embarque + Chr(13)
                                          var_referencia_Bancaria = ""
                                          If rsaux9.State = 1 Then
                                             rsaux9.Close
                                          End If
                                          rsaux9.Open "select vcha_cli_referencia from tb_Clientes where vcha_Cli_clave_id = '" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                          If Not rsaux9.EOF Then
                                             var_referencia_Bancaria = Trim(IIf(IsNull(rsaux9!VCHA_CLI_REFERENCIA), "", rsaux9!VCHA_CLI_REFERENCIA))
                                          End If
                                          rsaux9.Close
                                          var_cadena = var_cadena + "referenciabancaria=" + var_referencia_Bancaria + Chr(13)
                                          rsaux1.Open "select * from tb_encabezado_pedidos where inte_ped_numero = " + CStr(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero)), cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux1.EOF Then
                                             var_año = CStr(Year(rsaux1!dtim_ped_fecha))
                                             var_mes = CStr(Month(rsaux1!dtim_ped_fecha))
                                             var_dia = CStr(Day(rsaux1!dtim_ped_fecha))
                                             var_hora = CStr(Hour(rsaux1!dtim_ped_fecha))
                                             var_minuto = CStr(Minute(rsaux1!dtim_ped_fecha))
                                             var_segundo = CStr(Second(rsaux1!dtim_ped_fecha))
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
                                             var_cadena_fecha = var_año + "-" + var_mes + "-" + var_dia + "T" + var_hora + ":" + var_minuto + ":" + var_segundo
                                             var_cadena = var_cadena + "fechaPedido=" + var_cadena_fecha + Chr(13)
                                          Else
                                             var_cadena = var_cadena + "fechaPedido=" + var_cadena_fecha + Chr(13)
                                          End If
                                          rsaux1.Close
                                          var_cadena = var_cadena + "expedicion=" + var_expedido + Chr(13)
                                          var_cadena = var_cadena + "observaciones=" + Chr(13)
                                          var_cadena = var_cadena + "conceptoExtra1=" + Chr(13)
                                          var_cadena = var_cadena + "montoconceptoExtra1=" + Chr(13)
                                          var_cadena = var_cadena + "conceptoExtra2=" + Chr(13)
                                          var_cadena = var_cadena + "montoconceptoExtra2=" + Chr(13)
                                          var_cadena = var_cadena + "tipoimpresion=2" + Chr(13)
                                          rsaux9.Open "select * from vw_clientes where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux9.EOF Then
                                             var_cadena = var_cadena + "agente=" + rsaux9!VCHA_AGE_AGENTE_ID + " " + rsaux9!VCHA_AGE_NOMBRE + Chr(13)
                                          Else
                                             var_cadena = var_cadena + "agente=" + rs!VCHA_AGE_AGENTE_ID + " " + rs!VCHA_AGE_NOMBRE + Chr(13)
                                          End If
                                          rsaux9.Close
                                          If var_empresa = "15" Then
                                             'var_cadena = var_cadena + "formato=MHESTAMPADOS_V01.DAT" + Chr(13)
                                             var_cadena = var_cadena + "formato=MHVTH_V01.DAT" + Chr(13)
                                          End If
                                          If var_empresa = "16" Then
                                             'var_cadena = var_cadena + "formato=MHMULTIBONDEADOS_V01.DAT" + Chr(13)
                                             var_cadena = var_cadena + "formato=MHVTH_V01.DAT" + Chr(13)
                                          End If
                                          If var_empresa = "30" Then
                                             var_cadena = var_cadena + "formato=MHTURBINA_V01.DAT" + Chr(13)
                                          End If
                                          If var_empresa = "31" Then
                                             var_cadena = var_cadena + "formato=MHCANTIA_V01.DAT" + Chr(13)
                                          End If
                                          If var_empresa = "02" Or var_empresa = "03" Or var_empresa = "18" Or var_empresa = "17" Or var_empresa = "06" Or var_empresa = "29" Then
                                             If parametros(1) <> "SIDALMACENBKP" Then
                                                var_cadena = var_cadena + "formato=MHVTH_V01.DAT" + Chr(13)
                                             Else
                                                var_cadena = var_cadena + "formato=MHTST_V01.DAT" + Chr(13)
                                             End If
                                          End If
                                          
                                          If var_empresa = "32" Then
                                             var_cadena = var_cadena + "formato=MHARE_V01.DAT" + Chr(13)
                                          End If
                                          If var_empresa = "33" Then
                                             var_cadena = var_cadena + "formato=MHMPU_V01.DAT" + Chr(13)
                                          End If
                                          If var_empresa = "34" Then
                                             var_cadena = var_cadena + "formato=MHMUL_V01.DAT" + Chr(13)
                                          End If
                                          If var_empresa = "36" Then
                                             var_cadena = var_cadena + "formato=MHSME_V01.DAT" + Chr(13)
                                          End If
                                          If var_empresa = "37" Then
                                             var_cadena = var_cadena + "formato=MHVTH_V01.DAT" + Chr(13)
                                          End If
                                          If var_empresa = "38" Then
                                             var_cadena = var_cadena + "formato=MHVIANNEY_V01.DAT" + Chr(13)
                                          End If
                                          If var_empresa = "39" Then
                                             var_cadena = var_cadena + "formato=MHCANTIA_V01.DAT" + Chr(13)
                                          End If
                                          If var_empresa = "40" Then
                                             var_cadena = var_cadena + "formato=MHVIN_V01.DAT" + Chr(13)
                                          End If
                                          If var_empresa = "41" Then
                                             var_cadena = var_cadena + "formato=MHCOP_V01.DAT" + Chr(13)
                                          End If
                                          If var_empresa = "42" Then
                                             var_cadena = var_cadena + "formato=MHCMA_V01.DAT" + Chr(13)
                                          End If
                                          If var_empresa = "43" Then
                                             var_cadena = var_cadena + "formato=MHVOP_V01.DAT" + Chr(13)
                                          End If
                                          If var_empresa = "44" Then
                                             var_cadena = var_cadena + "formato=MHUTV_V01.DAT" + Chr(13)
                                          End If
                                          
                                          
                                          var_cadena = var_cadena + "</Otros>" + Chr(13) + Chr(13)
                                          var_cadena = var_cadena + "piezas_totales=" + CStr(var_piezas_totales) + Chr(13)
                                          var_cadena = var_cadena + "<addenda>" + Chr(13)
                                          var_cadena = var_cadena + "</addenda>" + Chr(13) + Chr(13)
                                          var_cadena = var_cadena + "</Factura>"
                                          Print #1, var_cadena
                                          

                                          
                                          Close #1
                                          
                                          var_Archivo = App.Path & "\renombra" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat"
                                          'MsgBox var_ruta_documentos_electronicos
                                          x = Shell(var_Archivo, vbHide)
                                          'Set fs = CreateObject("Scripting.FileSystemObject")
                                          'ArchivoOrigen = var_ruta_documentos_electronicos & "\" + Trim(var_serie) + Trim(Str(rsaux3!inte_car_numero)) + ".fi"
                                          'ArchivoDestino = var_ruta_documentos_electronicos & "\" + Trim(var_serie) + Trim(Str(rsaux3!inte_car_numero)) + ".ff"
                                          'fs.CopyFile ArchivoOrigen, ArchivoDestino
                                          'fs.DeleteFile ArchivoOrigen
                                          
                                       End If

End Sub

Function fun_copia_archivo(Origen, Destino)
    Copy_File = CopyFile(Origen, Destino, 1)
End Function

Private Sub reimprime_electronica()
                                          Open (App.Path & "\renombra" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat") For Output As #2
                                          Print #2, "ren " + var_ruta_documentos_electronicos + "\" + Trim(var_serie) + Trim(Str(rsaux3!inte_Car_numero)) + ".fi " + Trim(var_serie) + Trim(Str(rsaux3!inte_Car_numero)) + ".ff"
                                          Close #2

                                             If rs.State = 1 Then
                                                rs.Close
                                             End If
                                             If var_empresa <> "03" Then
                                                rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque_reimprimir + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' ORDER BY INTE_SAL_CONSECUTIVO_FACTURA", cnn, adOpenDynamic, adLockOptimistic
                                             Else
                                                rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque_reimprimir + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' ORDER BY vcha_sal_descripcion_factura", cnn, adOpenDynamic, adLockOptimistic
                                             End If
                                             If Not rs.EOF Then
                                                rsaux5.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque_reimprimir + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'  ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
                                                rsaux4.Open "select * from tb_encabezado_movimientos where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + rsaux5!VCHA_MOV_MOVIMIENTO_ID + "' and inte_Emo_numero = " + CStr(rsaux5!inte_emo_numero), cnn, adOpenDynamic, adLockOptimistic
                                                var_clasificacion_maquila = ""
                                                If Not rsaux4.EOF Then
                                                   var_clasificacion_maquila = IIf(IsNull(rsaux4!vcha_Emo_clasificacion), "", rsaux4!vcha_Emo_clasificacion)
                                                Else
                                                   var_clasificacion_maquila = ""
                                                End If
                                                rsaux4.Close
                                                rsaux5.Close
                                                'MsgBox var_ruta_documentos_electronicos
                                                Open (var_ruta_documentos_electronicos & "\" + Trim(var_serie) + Trim(Str(rsaux3!inte_Car_numero)) + ".fi") For Output As #1
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
                                                var_cadena_fecha = var_año + "-" + var_mes + "-" + var_dia + "T" + var_hora + ":" + var_minuto + ":" + var_segundo
                                                If rsaux1.State = 1 Then
                                                   rsaux1.Close
                                                End If
                                                rsaux1.Open "select * from tb_plazos where vcha_pla_plazo_id = '" + rs!VCHA_PLA_PLAZO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                                var_cadena = var_cadena + "fecha=" + var_cadena_fecha + Chr(13)
                                                var_cadena = var_cadena + "noAprobacion=" + Chr(13)
                                                var_cadena = var_cadena + "anoAprobacion=" + Chr(13)
                                                var_cadena = var_cadena + "tipoDeComprobante=FACTURA" + Chr(13)
                                                var_cadena = var_cadena + "formaDePago=PAGO HECHO EN UNA SOLA EXHIBICION" + Chr(13)
                                                If rsaux1.EOF Then
                                                   var_cadena = var_cadena + "condicionesDePago=" + "PLAZO CERO" + Chr(13)
                                                Else
                                                   var_cadena = var_cadena + "condicionesDePago=" + IIf(IsNull(rs!INTE_CAR_PLAZO), "PLAZO CERO", rs!INTE_CAR_PLAZO) + " DIAS" + Chr(13)
                                                End If
                                                rsaux1.Close
                                                'x = Format(CStr((rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio) / (1 + (rs!floa_car_porcentaje_iva / 100))), "###,###,###,##0.000000")
                                                'MsgBox x
                                                If rs!floa_Car_importe_neto = 0 Then
                                                   var_cadena = var_cadena + "subtotal=" + Format(CStr((0.01 / rs!floa_car_tipo_cambio) / (1 + (rs!floa_car_porcentaje_iva / 100))), "###,###,###,##0.000000") + Chr(13)
                                                Else
                                                   If Trim(rs!vcha_cli_clave_id) = "C000010568" Or Trim(rs!vcha_cli_clave_id) = "C000008200" Or Trim(rs!vcha_cli_clave_id) = "C000008200" Or Trim(rs!vcha_cli_clave_id) = "C000012251" Or Trim(rs!vcha_cli_clave_id) = "C000011510" Or Trim(rs!vcha_cli_clave_id) = "C000012528" Or Trim(rs!vcha_cli_clave_id) = "C000012558" Or Trim(rs!vcha_cli_clave_id) = "C000007909" Then
                                                      var_cadena = var_cadena + "subtotal=" + Format(CStr((rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio)), "###,###,###,##0.000000") + Chr(13)
                                                   Else
                                                      var_cadena = var_cadena + "subtotal=" + Format(CStr((rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio) / (1 + (rs!floa_car_porcentaje_iva / 100))), "###,###,###,##0.000000") + Chr(13)
                                                   End If
                                                End If
                                                
                                                var_cadena = var_cadena + "descuento=" + Chr(13)
                                                var_cadena = var_cadena + "descuento1=" + Format(CStr((rs!floa_car_importe_descuento_1 / rs!floa_car_tipo_cambio)), "###,###,##0.000000") + Chr(13)
                                                var_cadena = var_cadena + "descuento2=" + Format(CStr((rs!floa_car_importe_descuento_2) / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                                                var_cadena = var_cadena + "conceptodescuento1=DESCUENTO DEL " + Chr(13)
                                                var_cadena = var_cadena + "conceptodescuento2=DESCUENTO POR PAGO OPORTUNO " + Chr(13)
                                                var_cadena = var_cadena + "tasadescuento1=" + CStr(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)) + "%" + Chr(13)
                                                var_cadena = var_cadena + "tasadescuento2=" + CStr(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2)) + "%" + Chr(13)
                                                If rsaux1.State = 1 Then
                                                   rsaux1.Close
                                                End If
                                                If rsaux1.State = 1 Then
                                                   rsaux1.Close
                                                End If
                                                'rsaux1.Open "select * from tb_empresa_FACTURA_ELECTRONICA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                                                rsaux1.Open "select * from tb_empresa_FACTURA_ELECTRONICA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                                                'MsgBox "select * from tb_empresa_FACTURA_ELECTRONICA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'"
                                                'MsgBox "select * from tb_empresa_FACTURA_ELECTRONICA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'"
                                                var_certificado = rsaux1!vcha_emp_certificado
                                                var_expedido = rsaux1!vcha_emp_expedido
                                                If rs!floa_Car_importe_neto = 0 Then
                                                   var_cadena = var_cadena + "iva=" + Format(CStr(0# / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                                                Else
                                                   var_cadena = var_cadena + "iva=" + Format(CStr(rs!floa_car_importe_iva / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                                                End If
                                                If rs!floa_Car_importe_neto = 0 Then
                                                   var_cadena = var_cadena + "total=" + Format(CStr(0.01 / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                                                Else
                                                   var_cadena = var_cadena + "total=" + Format(CStr(rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                                                End If
                                                If Trim(rs!vcha_cli_clave_id) = "C000010568" Or Trim(rs!vcha_cli_clave_id) = "C000008200" Or Trim(rs!vcha_cli_clave_id) = "C000008200" Or Trim(rs!vcha_cli_clave_id) = "C000012251" Or Trim(rs!vcha_cli_clave_id) = "C000011510" Or Trim(rs!vcha_cli_clave_id) = "C000012528" Or Trim(rs!vcha_cli_clave_id) = "C000012558" Or Trim(rs!vcha_cli_clave_id) = "C000007909" Then
                                                   var_cadena = var_cadena + "retencion=" + Format(CStr(rs!floa_car_importe_iva / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                                                   var_cadena = var_cadena + "factorretencioniva=16%" + Chr(13)
                                                Else
                                                   var_cadena = var_cadena + "retencion=" + Chr(13)
                                                   var_cadena = var_cadena + "factorretencioniva=" + Chr(13)
                                                End If
                                                
                                                var_cadena = var_cadena + "</Comprobante>" + Chr(13) + Chr(13)
                                                var_cadena = var_cadena + "<Emisor>" + Chr(13)
                                                var_cadena = var_cadena + "erfc=" + rsaux1!VCHA_eMP_RFC + Chr(13)
                                                var_cadena = var_cadena + "enombre=" + rsaux1!VCHA_EMP_NOMBRE + Chr(13)
                                                var_cadena = var_cadena + "</Emisor>" + Chr(13) + Chr(13)
                                                var_cadena = var_cadena + "<DomicilioFiscal>" + Chr(13)
                                                var_cadena = var_cadena + "ecalle=" + rsaux1!VCHA_eMP_CALLE + Chr(13)
                                                var_cadena = var_cadena + "enoExterior=" + rsaux1!VCHA_eMP_exterior + Chr(13)
                                                var_cadena = var_cadena + "enoInterior=" + Chr(13)
                                                var_cadena = var_cadena + "ecolonia=" + rsaux1!VCHA_eMP_COLONIA + Chr(13)
                                                var_cadena = var_cadena + "elocalidad=" + rsaux1!VCHA_EMP_LOCALIDAD + Chr(13)
                                                var_cadena = var_cadena + "ereferencia=" + Chr(13)
                                                var_cadena = var_cadena + "emunicipio=" + rsaux1!VCHA_EMP_MUNICIPIO + Chr(13)
                                                var_cadena = var_cadena + "eestado=" + rsaux1!VCHA_EMP_ESTADO + Chr(13)
                                                var_cadena = var_cadena + "epais=" + rsaux1!VCHA_eMP_PAIS + Chr(13)
                                                var_cadena = var_cadena + "ecodigoPostal=" + rsaux1!VCHA_EMP_CODIGO_POSTAL + Chr(13)
                                                var_cadena = var_cadena + "etel=" + IIf(IsNull(rsaux1!VCHA_EMP_TELEFONO), "", rsaux1!VCHA_EMP_TELEFONO) + Chr(13)
                                                var_cadena = var_cadena + "eemail=" + IIf(IsNull(rsaux1!VCHA_EMP_EMAIL), "", rsaux1!VCHA_EMP_EMAIL) + Chr(13)
                                                'MsgBox var_cadena
                                                correo = IIf(IsNull(rsaux1!VCHA_EMP_EMAIL), "", rsaux1!VCHA_EMP_EMAIL)
                                                var_cadena = var_cadena + "</DomicilioFiscal>" + Chr(13) + Chr(13)
                                                var_cadena = var_cadena + "<Receptor>" + Chr(13)
                                                var_cadena = var_cadena + "noCliente=" + rs!vcha_cli_clave_id + Chr(13)
                                             
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
                                                var_cadena = var_cadena + "localidad=" + IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + Chr(13)
                                                rsaux1.Close
                                                rsaux1.Open "select * from vw_clientes where vcha_Cli_clave_id = '" + rs!vcha_cli_clave_id + "'"
                                                var_cadena = var_cadena + "referencia=" + Chr(13)
                                                var_cadena = var_cadena + "municipio=" + IIf(IsNull(rsaux1!vcha_mun_nombre), "", rsaux1!vcha_mun_nombre) + Chr(13)
                                                var_cadena = var_cadena + "estado=" + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre) + Chr(13)
                                                VAR_PAIS_ELECTRONICO = IIf(IsNull(rsaux1!vcha_pai_nombre), "MEXICO", rsaux1!vcha_pai_nombre)
                                                If Trim(VAR_PAIS_ELECTRONICO) = "" Then
                                                   VAR_PAIS_ELECTRONICO = "MEXICO"
                                                End If
                                                var_cadena = var_cadena + "pais=" + VAR_PAIS_ELECTRONICO + Chr(13)
                                                var_cadena = var_cadena + Chr(13)
                                                var_cadena = var_cadena + "codigoPostal=" + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP) + Chr(13)
                                                var_cadena = var_cadena + "tel=" + Chr(13)
                                                rsaux11.Open "select * from vw_clientes where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                                If IIf(IsNull(rsaux11!VCHA_AGE_AGENTE_ID), "", rsaux11!VCHA_AGE_AGENTE_ID) <> "00100" Then
                                                   var_cadena = var_cadena + "email=" + Chr(13)
                                                Else
                                                   var_cadena = var_cadena + "email=" + IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email) + Chr(13)
                                                End If
                                                rsaux11.Close
                                               'var_cadena = var_cadena + "email=" + iif(isnull(rs!vcha_cli_email
                                                var_cadena = var_cadena + "</Cliente>" + Chr(13) + Chr(13)
                                                
                                                var_cadena = var_cadena + "<EntregarEn>" + Chr(13)
                                                rsaux4.Open "select * from vw_establecimientos_direcciones where vcha_esb_establecimiento_id = '" + rs!vcha_ESB_ESTABLECIMIENTO_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                                var_cadena = var_cadena + "endomicilio=" + IIf(IsNull(rsaux4!vcha_esb_domicilio), "", rsaux4!vcha_esb_domicilio) + Chr(13)
                                                var_cadena = var_cadena + "encalle=" + Chr(13)
                                                var_cadena = var_cadena + "ennoExterior=" + Chr(13)
                                                var_cadena = var_cadena + "ennoInterior=" + Chr(13)
                                                var_cadena = var_cadena + "encolonia=" + IIf(IsNull(rsaux4!vcha_col_nombre), "", rsaux4!vcha_col_nombre) + Chr(13)
                                                var_cadena = var_cadena + "enlocalidad=" + IIf(IsNull(rsaux4!vcha_ciu_nombre), "", rsaux4!vcha_ciu_nombre) + Chr(13)
                                                var_cadena = var_cadena + "enreferencia=" + Chr(13)
                                                var_cadena = var_cadena + "enmunicipio=" + IIf(IsNull(rsaux4!vcha_mun_nombre), "", rsaux4!vcha_mun_nombre) + Chr(13)
                                                var_cadena = var_cadena + "enestado=" + IIf(IsNull(rsaux4!vcha_est_nombre), "", rsaux4!vcha_est_nombre) + Chr(13)
                                                var_cadena = var_cadena + "enpais=" + IIf(IsNull(rsaux4!vcha_pai_nombre), "", rsaux4!vcha_pai_nombre) + Chr(13)
                                                var_cadena = var_cadena + "encodigoPostal=" + Chr(13)
                                                var_cadena = var_cadena + "entel=" + Chr(13)
                                                var_cadena = var_cadena + "enemail=" + Chr(13)
                                                var_cadena = var_cadena + "</EntregarEn>" + Chr(13) + Chr(13)
                                                rsaux4.Close
                                                var_cadena = var_cadena + "<Concepto>" + Chr(13)
                                                rsaux1.Close
      
                                                var_k = 0
                                                var_piezas_totales = 0
                                                While Not rs.EOF
                                                     var_k = var_k + 1
                                                      pxx = CStr(var_k)
                                                      If Len(pxx) = 1 Then
                                                         pxx = "0" + pxx
                                                      End If
                                                      var_cadena = var_cadena + "p" + pxx + "_cantidad=" + CStr(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) + Chr(13)
                                                      var_piezas_totales = var_piezas_totales + IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                                                      rsaux1.Open "SELECT dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID, dbo.TB_UNIDADES.VCHA_UNI_UNIDAD_ID, dbo.TB_UNIDADES.VCHA_UNI_NOMBRE FROM dbo.TB_ARTICULOS LEFT OUTER JOIN dbo.TB_UNIDADES ON dbo.TB_ARTICULOS.VCHA_UNI_UNIDAD_ID = dbo.TB_UNIDADES.VCHA_UNI_UNIDAD_ID WHERE (dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID = '" + rs!vcha_agr_agrupador_id + "')"
                                                      If Not rsaux1.EOF Then
                                                         var_Unidad_medida = IIf(IsNull(rsaux1!VCHA_UNI_NOMBRE), "PZA", rsaux1!VCHA_UNI_NOMBRE)
                                                         If Len(Trim(var_Unidad_medida)) = 0 Then
                                                            var_Unidad_medida = "PZA"
                                                         End If
                                                         var_cadena = var_cadena + "p" + pxx + "_unidad=" + var_Unidad_medida + Chr(13)
                                                      Else
                                                         var_cadena = var_cadena + "p" + pxx + "_unidad=PZA" + Chr(13)
                                                      End If
                                                      rsaux1.Close
                                                      var_cadena = var_cadena + "p" + pxx + "_noIdentificacion=" + IIf(IsNull(rs!vcha_agr_agrupador_id), "", rs!vcha_agr_agrupador_id) + Chr(13)
                                                      If var_empresa = "15" Then
                                                         If var_clasificacion_maquila = "" Or var_clasificacion_maquila = "PRIMERA" Then
                                                            var_linea = IIf(IsNull(rs!vcha_agr_agrupador_id), "", rs!vcha_agr_agrupador_id) + " MAQUILA DE " + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                                         Else
                                                            var_linea = IIf(IsNull(rs!vcha_agr_agrupador_id), "", rs!vcha_agr_agrupador_id) + " MAQUILA DE SEGUNDA DE " + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                                         End If
                                                      Else
                                                         var_linea = IIf(IsNull(rs!vcha_agr_agrupador_id), "", rs!vcha_agr_agrupador_id) + " " + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                                      End If
                                                      var_cadena = var_cadena + "p" + pxx + "_descripcion=" + var_linea + Chr(13)
                                                      var_precio = IIf(IsNull(rs!Importe), 0, rs!Importe)
                                                      If IIf(IsNull(rs!vcha_agr_agrupador_id), "", rs!vcha_agr_agrupador_id) <> "---" Then
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
                                                         'var_precio = var_precio / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                                                      End If
                                                      var_cadena = var_cadena + "p" + pxx + "_valorUnitario=" + Format(CStr(var_precio / CStr(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad))), "###,###,##0.000000") + Chr(13)
                                                      var_cadena = var_cadena + "p" + pxx + "_importe=" + Format(CStr(var_precio), "###,###,##0.000000") + Chr(13)
                                                      
                                                      rs.MoveNext
                                                Wend
                                                var_cadena = var_cadena + "</Concepto>" + Chr(13) + Chr(13)
                                                var_cadena = var_cadena + "<Otros>" + Chr(13)
                                                var_cadena = var_cadena + "certificado=" + IIf(IsNull(var_certificado), "", var_certificado) + Chr(13)
                                                rs.MoveFirst
                                                var_cadena = var_cadena + "cant_letra=" + rs!vcha_car_importe_letra + Chr(13)
                                                var_cadena = var_cadena + "factoriva=" + CStr(rs!floa_car_porcentaje_iva) + "%" + Chr(13)
                                                var_cadena = var_cadena + "moneda=" + IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural) + Chr(13)
                                                var_cadena = var_cadena + "tipodeCambio=" + CStr(IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)) + Chr(13)
                                                var_cliente_coppel = rs!vcha_cli_clave_id
                                                If Trim(var_cliente_coppel) = "C000002947" Or Trim(var_cliente_coppel) = "C000005566" Or Trim(var_cliente_coppel) = "C000005831" Or var_empresa = "06" Or var_empresa = "17" Then
                                                   rsaux5.Open "select * from tb_encabezado_pedidos where inte_ped_numero = " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))), cnn, adOpenDynamic, adLockOptimistic
                                                   If Not rsaux5.EOF Then
                                                      If Trim(var_cliente_coppel) = "C000005831" Then
                                                         var_numero_pedido = Trim(CStr(IIf(IsNull(rsaux5!VCHA_PED_PEDIDO_EXTERNO), "", rsaux5!VCHA_PED_PEDIDO_EXTERNO)))
                                                      Else
                                                         If var_empresa = "06" Or var_empresa = "17" Then
                                                            var_numero_pedido = Trim(CStr(IIf(IsNull(rsaux5!VCHA_PED_PEDIDO_EXTERNO), "", rsaux5!VCHA_PED_PEDIDO_EXTERNO)))
                                                         Else
                                                             var_numero_pedido = Trim(CStr(IIf(IsNull(rsaux5!VCHA_PED_PEDIDO_EXTERNO), "", rsaux5!VCHA_PED_PEDIDO_EXTERNO)))
                                                         End If
                                                      End If
                                                   End If
                                                   rsaux5.Close
                                                Else
                                                   var_numero_pedido = Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero)))
                                                End If
                                                var_cadena = var_cadena + "pedido=" + Trim(var_numero_pedido) + Chr(13)
                                                var_cadena = var_cadena + "Embarque=" + Me.txt_numero_embarque + Chr(13)
                                                var_referencia_Bancaria = ""
                                                rsaux11.Open "select vcha_cli_referencia from tb_Clientes where vcha_Cli_clave_id = '" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                                If Not rsaux11.EOF Then
                                                   var_referencia_Bancaria = Trim(IIf(IsNull(rsaux11!VCHA_CLI_REFERENCIA), "", rsaux11!VCHA_CLI_REFERENCIA))
                                                End If
                                                rsaux11.Close
                                                var_cadena = var_cadena + "referenciabancaria=" + var_referencia_Bancaria + Chr(13)
                                                rsaux1.Open "select * from tb_encabezado_pedidos where inte_ped_numero = " + CStr(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero)), cnn, adOpenDynamic, adLockOptimistic
                                                If Not rsaux1.EOF Then
                                                   var_año = CStr(Year(rsaux1!dtim_ped_fecha))
                                                   var_mes = CStr(Month(rsaux1!dtim_ped_fecha))
                                                   var_dia = CStr(Day(rsaux1!dtim_ped_fecha))
                                                   var_hora = CStr(Hour(rsaux1!dtim_ped_fecha))
                                                   var_minuto = CStr(Minute(rsaux1!dtim_ped_fecha))
                                                   var_segundo = CStr(Second(rsaux1!dtim_ped_fecha))
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
                                                   var_cadena_fecha = var_año + "-" + var_mes + "-" + var_dia + "T" + var_hora + ":" + var_minuto + ":" + var_segundo
                                                   var_cadena = var_cadena + "fechaPedido=" + var_cadena_fecha + Chr(13)
                                                Else
                                                   var_cadena = var_cadena + "fechaPedido=" + var_cadena_fecha + Chr(13)
                                                End If
                                                rsaux1.Close
                                                var_cadena = var_cadena + "expedicion=" + var_expedido + Chr(13)
                                                var_cadena = var_cadena + "observaciones=" + Chr(13)
                                                var_cadena = var_cadena + "conceptoExtra1=" + Chr(13)
                                                var_cadena = var_cadena + "montoconceptoExtra1=" + Chr(13)
                                                var_cadena = var_cadena + "conceptoExtra2=" + Chr(13)
                                                var_cadena = var_cadena + "montoconceptoExtra2=" + Chr(13)
                                                var_cadena = var_cadena + "tipoimpresion=2" + Chr(13)
                                                rsaux11.Open "select * from vw_clientes where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                                If Not rsaux11.EOF Then
                                                   var_cadena = var_cadena + "agente=" + rsaux11!VCHA_AGE_AGENTE_ID + " " + rsaux11!VCHA_AGE_NOMBRE + Chr(13)
                                                Else
                                                   var_cadena = var_cadena + "agente=" + rs!VCHA_AGE_AGENTE_ID + " " + rs!VCHA_AGE_NOMBRE + Chr(13)
                                                End If
                                                rsaux11.Close
                                                If var_empresa = "15" Then
                                                   var_cadena = var_cadena + "formato=MHESTAMPADOS_V01.DAT" + Chr(13)
                                                End If
                                                If var_empresa = "16" Then
                                                   var_cadena = var_cadena + "formato=MHMULTIBONDEADOS_V01.DAT" + Chr(13)
                                                End If
                                                If var_empresa = "31" Then
                                                   var_cadena = var_cadena + "formato=MHCANTIA_V01.DAT" + Chr(13)
                                                End If
                                                
                                                If var_empresa = "38" Then
                                                   var_cadena = var_cadena + "formato=MHVIANNEY_V01.DAT" + Chr(13)
                                                End If
                                               If var_empresa = "02" Or var_empresa = "03" Or var_empresa = "18" Or var_empresa = "17" Or var_empresa = "06" Then
                                                  var_cadena = var_cadena + "formato=MHVTH_V01.DAT" + Chr(13)
                                               End If
                                                
                                                
                                                var_cadena = var_cadena + "</Otros>" + Chr(13) + Chr(13)
                                                var_cadena = var_cadena + "piezas_totales =" + CStr(var_piezas_totales) + Chr(13)
                                                var_cadena = var_cadena + "<addenda>" + Chr(13)
                                                var_cadena = var_cadena + "</addenda>" + Chr(13) + Chr(13)
                                                var_cadena = var_cadena + "</Factura>"
                                                Print #1, var_cadena
                                                Close #1
                                                var_Archivo = App.Path & "\renombra" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat"
                                                
                                                
                                             End If
                                          var_Archivo = App.Path & "\renombra" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat"
                                          x = Shell(var_Archivo, vbHide)
                                             

End Sub


Private Sub envio_tb_transito()
    var_cadena = " SELECT dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID, dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID, dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID, dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_SALIDAS.DTIM_SAL_FECHA, dbo.TB_SALIDAS.INTE_SAL_NUMERO, dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID, dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, dbo.TB_SALIDAS.FLOA_SAL_COSTO, dbo.TB_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO,  dbo.TB_SALIDAS.VCHA_SER_SERIE_ID, dbo.TB_SALIDAS.INTE_CAR_NUMERO, dbo.TB_SALIDAS.FLOA_SAL_PROMOCION_1, dbo.TB_SALIDAS.FLOA_SAL_PROMOCION_2, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID , dbo.TB_ENCABEZADO_CARTERA.VCHA_AGE_AGENTE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ENCABEZADO_CARTERA.VCHA_ESB_ESTABLECIMIENTO_ID FROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_SALIDAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND "
    var_cadena = var_cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON  dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_SALIDAS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID WHERE (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_numero_embarque + ") AND (dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')"
    If rs.State = 1 Then
       rs.Close
    End If
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
                   rsaux9.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_aRTICULO_ID = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                   If Not rsaux9.EOF Then
                      var_descripcion_articulo = IIf(IsNull(rsaux9!vcha_Art_nombre_español), "", rsaux9!vcha_Art_nombre_español)
                   Else
                      var_descripcion_articulo = ""
                   End If
                   rsaux9.Close
                   var_descuento_1 = IIf(IsNull(rs!FLOA_SAL_DESCUENTO_1), 0, rs!FLOA_SAL_DESCUENTO_1)
                   var_descuento_2 = IIf(IsNull(rs!FLOA_SAL_DESCUENTO_2), 0, rs!FLOA_SAL_DESCUENTO_2)
                   var_costo = rs!floa_Sal_precio * (1 - (var_descuento_1 / 100))
                   var_costo = var_costo * (1 - (var_descuento_2 / 100))
                   rsaux10.Open "select * from tb_transito where vcha_tra_nota_envio = '" + var_clave_planta_origen + "_" + CStr(rs!inte_Car_numero) + "' and vcha_Art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
                   If rsaux10.EOF Then
                      var_cadena = "insert into tb_transito (vcha_tra_nota_envio, vcha_Art_Articulo_id,                                                              vcha_Art_descripcion,           floa_Tra_cantidad_Enviada,                            floa_tra_costo, vcha_tra_planta_origen, vcha_tra_planta_destino, floa_tra_Cantidad_recibida, vcha_tra_Calidad,VCHA_TRA_STATUS,VCHA_MOV_MOVIMIENTO_ID, VCHA_EMP_EMPRESA_ID, VCHA_SER_SERIE_ID) "
                      var_cadena = var_cadena + "   values  ('" + var_clave_planta_origen + "_" + CStr(rs!inte_Car_numero) + "', '" + rs!VCHA_ART_ARTICULO_ID + "','" + var_descripcion_articulo + "', " + CStr(rs!floa_Sal_Cantidad) + ", " + CStr(var_costo) + ",'" + var_clave_planta_origen + "','" + var_clave_planta_destino + "',0,'1','A','EI','" + var_empresa + "','" + rs!vcha_Ser_Serie_id + "')"
                      rsaux9.Open var_cadena, cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
                   End If
                   rsaux10.Close
                   rs.MoveNext
             Wend
          End If
       End If
    End If
    rs.Close
End Sub



Private Sub factura_bordalesa()
   Dim var_nombre_unidad As String
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
   Dim ndo As New aClsNodoArbolTrazabilidad
   Dim var_establecimiento_comercial As String
   Dim var_solicitud_sigo As String
   Dim var_empresa_06 As String
   Dim var_clasificacion_maquila As String
   cnn.CommandTimeout = 360
                        
                     
   rsaux3.Open "select distinct inte_car_numero from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'  ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux3.EOF Then
      rsaux5.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'  ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
      
      rsaux4.Open "select * from tb_encabezado_movimientos where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + rsaux5!VCHA_MOV_MOVIMIENTO_ID + "' and inte_Emo_numero = " + CStr(rsaux5!inte_emo_numero), cnn, adOpenDynamic, adLockOptimistic
      var_clasificacion_maquila = ""
      If Not rsaux4.EOF Then
         var_clasificacion_maquila = IIf(IsNull(rsaux4!vcha_Emo_clasificacion), "", rsaux4!vcha_Emo_clasificacion)
      Else
         var_clasificacion_maquila = ""
      End If
      rsaux4.Close
      rsaux5.Close
      var_Archivo = App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat"
      Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat") For Output As #2
      While Not rsaux3.EOF
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' ORDER BY INTE_SAL_CONSECUTIVO_FACTURA", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               'AQUI EMPIEZA LA FACTURA
               Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt") For Output As #1
               'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
               'Print #1, Chr(27) + Chr(67) + Chr(44) + Chr(13)
               'Print #1, ""
               Print #1, Chr(15) + Chr(27) + Chr(64)
               'Print #1, ""
               'Print #1, ""
               'Print #1, ""
               'Print #1, ""
               'Print #1, ""
               'Print #1, ""
               'Print #1, ""
               'Print #1, ""
               var_dia = Day(rs!dtim_Car_fecha)
               var_mes_numero = Month(rs!dtim_Car_fecha)
               If var_mes_numero = 1 Then
                  var_mes = "ENERO"
               End If
               If var_mes_numero = 2 Then
                  var_mes = "FEBRERO"
               End If
               If var_mes_numero = 3 Then
                  var_mes = "MARZO"
               End If
               If var_mes_numero = 4 Then
                  var_mes = "ABRIL"
               End If
               If var_mes_numero = 5 Then
                  var_mes = "MAYO"
               End If
               If var_mes_numero = 6 Then
                  var_mes = "JUNIO"
               End If
               If var_mes_numero = 7 Then
                  var_mes = "JULIO"
               End If
               If var_mes_numero = 8 Then
                  var_mes = "AGOSTO"
               End If
               If var_mes_numero = 9 Then
                  var_mes = "SEPTIEMBRE"
               End If
               If var_mes_numero = 10 Then
                  var_mes = "OCTUBRE"
               End If
               If var_mes_numero = 11 Then
                  var_mes = "NOVIEMBRE"
               End If
               If var_mes_numero = 12 Then
                  var_mes = "DICIEMBRE"
               End If
               
               var_año = Year(rs!dtim_Car_fecha)
               var_cadena = "                                                                                                                         " + Str(rsaux3!inte_Car_numero)
               For var_j = Len(var_cadena) To 132
                   var_cadena = var_cadena + " "
               Next var_j
               var_cadena = var_cadena + "EMB.: " + txt_numero_embarque
               Print #1, var_cadena
               Print #1, ""
               var_cadena = "                                                                                                                   " + CStr(rs!inte_pla_dias) + " dias"
               For var_j = Len(var_cadena) To 150
                   var_cadena = var_cadena + " "
               Next var_j
               var_cadena = var_cadena + CStr(Format(rs!dtim_Car_fecha, "Short Date"))
               Print #1, " "
               'Print #1, ""
               Print #1, var_cadena
               Print #1, ""
               Print #1, ""
               'Print #1, ""
               var_cliente = "           " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
               For var_j = 1 + Len(Trim(var_cliente)) To 77
                   var_cliente = var_cliente + " "
               Next var_j
               var_cliente = var_cliente + "  " + rs!vcha_cli_clave_id + "               AGUASCALIENTES, AGS."
               Print #1, Spc(5); var_cliente
               'Print #1, ""
               var_domicilio = "           " + IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION)
               For var_j = 1 + Len(Trim(var_domicilio)) To 80
                   var_domicilio = var_domicilio + " "
               Next var_j
               Print #1, Spc(5); var_domicilio
               var_colonia = "           " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
               For var_j = 1 + Len(Trim(var_colonia)) To 48
                   var_colonia = var_colonia + " "
               Next var_j
               var_colonia = var_colonia + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
               Print #1, Spc(5); var_colonia
               'Print #1, ""
               var_ciudad = "           " + IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + ""
               For var_j = Len(var_ciudad) To 58
                   var_ciudad = var_ciudad + " "
               Next var_j
               var_ciudad = var_ciudad + IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
               
               For var_j = 1 + Len(Trim(var_ciudad)) To 77
                   var_ciudad = var_ciudad + " "
               Next var_j
               
               Print #1, Spc(5); var_ciudad + "     " + rs!VCHA_AGE_AGENTE_ID + "       " + IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
               
               'Print #1, ""
               var_cp = "           "
               var_cp = var_cp
               For var_j = 1 + Len(Trim(var_cp)) To 80
                   var_cp = var_cp + " "
               Next var_j
               Print #1, Spc(5); var_cp
               
               Print #1, ""
               var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
               var_importe_descuento_1 = 0
               var_importe_descuento_2 = 0
               var_importe_descuento_3 = 0
               var_contador_promociones = 0
               var_cantidad_total = 0
               'var_renglones_factura = 21
               Print #1, ""
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
                      For var_j = 1 + Len(Trim(var_linea)) To 24
                          var_linea = var_linea + " "
                      Next var_j
                      If var_empresa = "15" Then
                         If var_clasificacion_maquila = "" Or var_clasificacion_maquila = "PRIMERA" Then
                            var_linea = var_linea + "MAQUILA DE " + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                         Else
                            var_linea = var_linea + Mid("MAQUILA DE SEGUNDA DE " + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura), 1, 65)
                         End If
                      Else
                         var_linea = var_linea + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                      End If
                      var_i = 0
                                     
                      ''' imprimir cantidad en la orilla
                                       
                      ''' imprimir cantidad en la orilla
                                      
                                       
                      While Len((var_linea)) < 90
                            var_linea = var_linea + " "
                      Wend
                      var_linea = var_linea + " "
                      var_linea = var_linea + var_marca_promocion
                      var_cantidad = Format(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                      var_cantidad_total = var_cantidad_total + IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                      If Len(Trim(var_cantidad)) < 20 Then
                         For var_j = 1 + Len(Trim(var_cantidad)) To 20
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
                      var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                      If Len(Trim(var_rfc)) > 0 Then
                         var_precio_str = Format(IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                      Else
                         var_precio_str = Format((IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) * (1 + (rs!floa_car_porcentaje_iva / 100)), "###,###,##0.00")
                      End If
                      If Len(Trim(var_precio_str)) < 20 Then
                         For var_j = 1 + Len(Trim(var_precio_str)) To 20
                             var_precio_str = " " + var_precio_str
                         Next var_j
                      End If
                      var_linea = var_linea + var_cantidad + var_precio_str
                      If Len(Trim(var_rfc)) > 0 Then
                         var_importe = Format((IIf(IsNull(rs!Importe), 0, rs!Importe)), "###,###,##0.00")
                         If Len(Trim(var_importe)) < 27 Then
                            For var_j = 1 + Len(Trim(var_importe)) To 27
                                var_importe = " " + var_importe
                            Next var_j
                         End If
                      Else
                         var_importe = Format((IIf(IsNull(rs!Importe), 0, rs!Importe) * (1 + (rs!floa_car_porcentaje_iva / 100))), "###,###,##0.00")
                         If Len(Trim(var_importe)) < 27 Then
                            For var_j = 1 + Len(Trim(var_importe)) To 27
                                var_importe = " " + var_importe
                            Next var_j
                         End If
                      End If
                      var_linea = var_linea + var_importe
                                     
                      Print #1, Spc(5); var_linea
                      rs.MoveNext
                   Else
                      Print #1, ""
                   End If
               Next var_k
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
               var_descuento_leyenda = 0
               var_descuento_leyenda = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
               var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
               If Len(Trim(var_linea)) < 115 Then
                  For var_j = 1 + Len(Trim(var_linea)) To 115
                      var_linea = var_linea + " "
                  Next var_j
               End If
               var_linea = var_linea + var_importe_descuento_1_str
               var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%"
               If Len(Trim(var_linea)) < 115 Then
                  For var_j = 1 + Len(Trim(var_linea)) To 115
                      var_linea = var_linea + " "
                  Next var_j
               End If
               var_linea = var_linea + var_importe_descuento_2_str
               var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
               Print #1, Spc(25); IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
               ' aqui se puso la leyenda de pago en una sola exhibicion el 190310
               'Print #1, "                                                                          ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                        "
               Print #1, ""
               Dim var_fecha_pago As Date
               var_fecha_pago = rs!dtim_Car_fecha + IIf(IsNull(rs!inte_pla_dias), 0, rs!inte_pla_dias)
               var_dia_Pago = Day(var_fecha_pago)
               var_mes_pago = Month(var_fecha_pago)
               var_año_pago_str = CStr(Year(var_fecha_pago) - 2000)
               
               If var_mes_pago = 1 Then
                  var_mes = "ENERO"
               End If
               If var_mes_pago = 2 Then
                  var_mes = "FEBRERO"
               End If
               If var_mes_pago = 3 Then
                  var_mes = "MARZO"
               End If
               If var_mes_pago = 4 Then
                  var_mes = "ABRIL"
               End If
               If var_mes_pago = 5 Then
                  var_mes = "MAYO"
               End If
               If var_mes_pago = 6 Then
                  var_mes = "JUNIO"
               End If
               If var_mes_pago = 7 Then
                  var_mes = "JULIO"
               End If
               If var_mes_pago = 8 Then
                  var_mes = "AGOSTO"
               End If
               If var_mes_pago = 9 Then
                  var_mes = "SEPTIEMBRE"
               End If
               If var_mes_pago = 10 Then
                  var_mes = "OCTUBRE"
               End If
               If var_mes_pago = 11 Then
                  var_mes = "NOVIEMBRE"
               End If
               If var_mes_pago = 12 Then
                  var_mes = "DICIEMBRE"
               End If
               
               If Len(var_año_pago_str) = 1 Then
                  var_año_pago_str = "0" + var_año_pago_str
               End If
               
               If Len(Trim(var_rfc)) = 0 Then
                  var_subimporte = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                  If Len(Trim(var_subimporte)) < 14 Then
                     For var_j = 1 + Len(Trim(var_subimporte)) To 14
                         var_subimporte = " " + var_subimporte
                     Next var_j
                  End If
                  
                  var_linea = "                    ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION"
                  If Len(var_linea) < 145 Then
                     For var_j = 1 + Len(var_linea) To 145
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  var_linea = var_linea + var_subimporte
                  Print #1, Spc(5); var_linea
                  
                  'Print #1, ""
                  var_linea = ""
                  If Len((var_linea)) < 145 Then
                     For var_j = 1 + Len((var_linea)) To 145
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  var_iva = "-"
                  For var_j = 1 + Len(Trim(var_iva)) To 14
                      var_iva = " " + var_iva
                  Next var_j
                  
                  var_linea = var_linea + var_iva
                  Print #1, Spc(5); var_linea
                  'Print #1, ""
                  var_linea = ""
                  var_importe = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                  If Len(Trim(var_importe)) < 14 Then
                     For var_j = 1 + Len(Trim(var_importe)) To 14
                         var_importe = " " + var_importe
                     Next var_j
                  End If
                  
                  'var_linea = "" + var_año_pago_str + "          " + var_importe
                  var_linea = "" + var_linea + "          "
                  If Len((var_linea)) < 145 Then
                     For var_j = 1 + Len((var_linea)) To 145
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  var_linea = var_linea + var_importe
                  Print #1, Spc(5); var_linea
               Else
                  var_subimporte = Format(Round(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                  
                  If Len(Trim(var_subimporte)) < 14 Then
                     For var_j = 1 + Len(Trim(var_subimporte)) To 14
                         var_subimporte = " " + var_subimporte
                     Next var_j
                  End If
                  
                  var_linea = "                      ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION"
                  If Len(var_linea) < 145 Then
                     For var_j = 1 + Len(var_linea) To 145
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  'MsgBox var_linea
                  var_linea = var_linea + var_subimporte
                  Print #1, Spc(5); var_linea
                  
                  'Print #1, ""
                  'var_linea = "                                                              " + CStr(var_dia_Pago) + "   " + var_mes
                  var_linea = ""
                  If Len((var_linea)) < 145 Then
                     For var_j = 1 + Len((var_linea)) To 145
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  
                  var_iva = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                  For var_j = 1 + Len(Trim(var_iva)) To 14
                      var_iva = " " + var_iva
                  Next var_j
                  var_linea = var_linea + var_iva
                  Print #1, Spc(5); var_linea
                  'Print #1, ""
                  
                  
                  var_importe = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                  
                  If Len(Trim(var_importe)) < 14 Then
                     For var_j = 1 + Len(Trim(var_importe)) To 14
                         var_importe = " " + var_importe
                     Next var_j
                  End If
                  var_linea = ""
                  var_linea = "" + var_linea + "          "
                  If Len((var_linea)) < 145 Then
                     For var_j = 1 + Len((var_linea)) To 145
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  var_linea = var_linea + var_importe
                  Print #1, Spc(5); var_linea
               
               End If
               var_linea = ""
               var_linea = "                                                              " + CStr(var_dia_Pago) + "   " + var_mes
               Print #1, Spc(5); var_linea
               Print #1, Spc(5); var_año_pago_str + "          " + var_importe + "    " + IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, "           " + rs!VCHA_CLI_NOMBRE
               Print #1, "           " + rs!VCHA_CLI_DIRECCION
               Print #1, "           " + rs!vcha_ciu_nombre
               If var_empresa <> "03" Then
                  Print #1, ""
                  Print #1, ""
               Else
                  Print #1, ""
                  Print #1, ""
               End If
               Print #1, ""
               Close #1
               Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt lpt1"
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

End Sub

'---
Private Sub factura_estampados()
   Dim var_nombre_unidad As String
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
   Dim ndo As New aClsNodoArbolTrazabilidad
   Dim var_establecimiento_comercial As String
   Dim var_solicitud_sigo As String
   Dim var_empresa_06 As String
   Dim var_clasificacion_maquila As String
   cnn.CommandTimeout = 360
                        
                     
   rsaux3.Open "select distinct inte_car_numero from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'  ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux3.EOF Then
      rsaux5.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'  ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
      
      rsaux4.Open "select * from tb_encabezado_movimientos where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + rsaux5!VCHA_MOV_MOVIMIENTO_ID + "' and inte_Emo_numero = " + CStr(rsaux5!inte_emo_numero), cnn, adOpenDynamic, adLockOptimistic
      var_clasificacion_maquila = ""
      If Not rsaux4.EOF Then
         var_clasificacion_maquila = IIf(IsNull(rsaux4!vcha_Emo_clasificacion), "", rsaux4!vcha_Emo_clasificacion)
      Else
         var_clasificacion_maquila = ""
      End If
      rsaux4.Close
      
      rsaux5.Close
      var_Archivo = App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat"
      Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat") For Output As #2
      While Not rsaux3.EOF
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' ORDER BY INTE_SAL_CONSECUTIVO_FACTURA", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               'AQUI EMPIEZA LA FACTURA
               Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt") For Output As #1
               'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
               'Print #1, Chr(27) + Chr(67) + Chr(44) + Chr(13)
               'Print #1, ""
               Print #1, Chr(15) + Chr(27) + Chr(64)
               'Print #1, ""
               'Print #1, ""
               'Print #1, ""
               'Print #1, ""
               'Print #1, ""
               'Print #1, ""
               'Print #1, ""
               'Print #1, ""
               var_dia = Day(rs!dtim_Car_fecha)
               var_mes_numero = Month(rs!dtim_Car_fecha)
               If var_mes_numero = 1 Then
                  var_mes = "ENERO"
               End If
               If var_mes_numero = 2 Then
                  var_mes = "FEBRERO"
               End If
               If var_mes_numero = 3 Then
                  var_mes = "MARZO"
               End If
               If var_mes_numero = 4 Then
                  var_mes = "ABRIL"
               End If
               If var_mes_numero = 5 Then
                  var_mes = "MAYO"
               End If
               If var_mes_numero = 6 Then
                  var_mes = "JUNIO"
               End If
               If var_mes_numero = 7 Then
                  var_mes = "JULIO"
               End If
               If var_mes_numero = 8 Then
                  var_mes = "AGOSTO"
               End If
               If var_mes_numero = 9 Then
                  var_mes = "SEPTIEMBRE"
               End If
               If var_mes_numero = 10 Then
                  var_mes = "OCTUBRE"
               End If
               If var_mes_numero = 11 Then
                  var_mes = "NOVIEMBRE"
               End If
               If var_mes_numero = 12 Then
                  var_mes = "DICIEMBRE"
               End If
               
               var_año = Year(rs!dtim_Car_fecha)
               var_cadena = "                                                                                                                         " + Str(rsaux3!inte_Car_numero)
               For var_j = Len(var_cadena) To 132
                   var_cadena = var_cadena + " "
               Next var_j
               var_cadena = var_cadena + "EMB.: " + txt_numero_embarque
               Print #1, ""
               Print #1, var_cadena
               var_cadena = "                                                                                 AGUASCALIENTES, AGS.              " + CStr(rs!inte_pla_dias) + " dias"
               For var_j = Len(var_cadena) To 150
                   var_cadena = var_cadena + " "
               Next var_j
               var_cadena = var_cadena + CStr(Format(rs!dtim_Car_fecha, "Short Date"))
               Print #1, " "
               'Print #1, ""
               Print #1, var_cadena
               Print #1, ""
               'Print #1, ""
               rsaux10.Open "select * from vw_establecimientos_direcciones where vcha_esb_establecimiento_id = '" + rs!vcha_ESB_ESTABLECIMIENTO_id + "'", cnn, adOpenDynamic, adLockOptimistic
               var_cliente = "           " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
               For var_j = 1 + Len(Trim(var_cliente)) To 80
                   var_cliente = var_cliente + " "
               Next var_j
               Print #1, Spc(5); var_cliente + IIf(IsNull(rsaux10!VCHA_ESB_NOMBRE), "", rsaux10!VCHA_ESB_NOMBRE)
               'Print #1, ""
               var_domicilio = "           " + IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION)
               For var_j = 1 + Len(Trim(var_domicilio)) To 80
                   var_domicilio = var_domicilio + " "
               Next var_j
               Print #1, Spc(5); var_domicilio + IIf(IsNull(rsaux10!vcha_esb_domicilio), "", rsaux10!vcha_esb_domicilio)
               Print #1, ""
               var_colonia = "           " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
               For var_j = 1 + Len(Trim(var_colonia)) To 80
                   var_colonia = var_colonia + " "
               Next var_j
               
               Print #1, Spc(5); var_colonia + IIf(IsNull(rsaux10!vcha_col_nombre), "", rsaux10!vcha_col_nombre)
               'Print #1, ""
               var_ciudad = "           " + IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + ""
               For var_j = Len(var_ciudad) To 50
                   var_ciudad = var_ciudad + " "
               Next var_j
               var_ciudad = var_ciudad + "  " + rs!vcha_cli_clave_id
               
               For var_j = 1 + Len(Trim(var_ciudad)) To 80
                   var_ciudad = var_ciudad + " "
               Next var_j
               
               Print #1, Spc(5); var_ciudad + IIf(IsNull(rsaux10!vcha_ciu_nombre), "", rsaux10!vcha_ciu_nombre)
               
               'Print #1, ""
               var_cp = "           " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
               For var_j = Len(var_cp) To 45
                   var_cp = var_cp + " "
               Next var_j
               var_cp = var_cp + IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
               For var_j = 1 + Len(Trim(var_cp)) To 80
                   var_cp = var_cp + " "
               Next var_j
               Print #1, Spc(5); var_cp + IIf(IsNull(rsaux10!vcha_esb_cp), "", rsaux10!vcha_esb_cp)
               
               rsaux10.Close
               Print #1, ""
               Print #1, ""
               Print #1, ""
               var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
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
                      For var_j = 1 + Len(Trim(var_linea)) To 24
                          var_linea = var_linea + " "
                      Next var_j
                      If var_empresa = "15" Then
                         If var_clasificacion_maquila = "" Or var_clasificacion_maquila = "PRIMERA" Then
                            var_linea = var_linea + "MAQUILA DE " + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                         Else
                            var_linea = var_linea + Mid("MAQUILA DE SEGUNDA DE " + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura), 1, 65)
                         End If
                      Else
                         var_linea = var_linea + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                      End If
                      var_i = 0
                                     
                      ''' imprimir cantidad en la orilla
                                       
                      ''' imprimir cantidad en la orilla
                                      
                                       
                      While Len((var_linea)) < 90
                            var_linea = var_linea + " "
                      Wend
                      var_linea = var_linea + " "
                      var_linea = var_linea + var_marca_promocion
                      var_cantidad = Format(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                      var_cantidad_total = var_cantidad_total + IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                      If Len(Trim(var_cantidad)) < 20 Then
                         For var_j = 1 + Len(Trim(var_cantidad)) To 20
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
                      var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                      If Len(Trim(var_rfc)) > 0 Then
                         var_precio_str = Format(IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                      Else
                         var_precio_str = Format((IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) * (1 + (rs!floa_car_porcentaje_iva / 100)), "###,###,##0.00")
                      End If
                      If Len(Trim(var_precio_str)) < 20 Then
                         For var_j = 1 + Len(Trim(var_precio_str)) To 20
                             var_precio_str = " " + var_precio_str
                         Next var_j
                      End If
                      var_linea = var_linea + var_cantidad + var_precio_str
                      If Len(Trim(var_rfc)) > 0 Then
                         var_importe = Format((IIf(IsNull(rs!Importe), 0, rs!Importe)), "###,###,##0.00")
                         If Len(Trim(var_importe)) < 27 Then
                            For var_j = 1 + Len(Trim(var_importe)) To 27
                                var_importe = " " + var_importe
                            Next var_j
                         End If
                      Else
                         var_importe = Format((IIf(IsNull(rs!Importe), 0, rs!Importe) * (1 + (rs!floa_car_porcentaje_iva / 100))), "###,###,##0.00")
                         If Len(Trim(var_importe)) < 27 Then
                            For var_j = 1 + Len(Trim(var_importe)) To 27
                                var_importe = " " + var_importe
                            Next var_j
                         End If
                      End If
                      var_linea = var_linea + var_importe
                                     
                      Print #1, Spc(5); var_linea
                      rs.MoveNext
                   Else
                      Print #1, ""
                   End If
               Next var_k
               Print #1, ""
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
               var_descuento_leyenda = 0
               var_descuento_leyenda = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
               var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
               If Len(Trim(var_linea)) < 115 Then
                  For var_j = 1 + Len(Trim(var_linea)) To 115
                      var_linea = var_linea + " "
                  Next var_j
               End If
               var_linea = var_linea + var_importe_descuento_1_str
               'Print #1, Spc(5); var_linea
               Print #1, Spc(5); ""
               var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%"
               If Len(Trim(var_linea)) < 115 Then
                  For var_j = 1 + Len(Trim(var_linea)) To 115
                      var_linea = var_linea + " "
                  Next var_j
               End If
               var_linea = var_linea + var_importe_descuento_2_str
               'Print #1, Spc(5); var_linea
               Print #1, Spc(5); ""
               
               var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
               Print #1, ""
               Print #1, ""
               Print #1, Spc(5); IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
               ' aqui se puso la leyenda de pago en una sola exhibicion el 190310
               Print #1, "                                                                          ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                        "
               Dim var_fecha_pago As Date
               var_fecha_pago = rs!dtim_Car_fecha + IIf(IsNull(rs!inte_pla_dias), 0, rs!inte_pla_dias)
               var_dia_Pago = Day(var_fecha_pago)
               var_mes_pago = Month(var_fecha_pago)
               var_año_pago_str = CStr(Year(var_fecha_pago) - 2000)
               
               If var_mes_pago = 1 Then
                  var_mes = "ENERO"
               End If
               If var_mes_pago = 2 Then
                  var_mes = "FEBRERO"
               End If
               If var_mes_pago = 3 Then
                  var_mes = "MARZO"
               End If
               If var_mes_pago = 4 Then
                  var_mes = "ABRIL"
               End If
               If var_mes_pago = 5 Then
                  var_mes = "MAYO"
               End If
               If var_mes_pago = 6 Then
                  var_mes = "JUNIO"
               End If
               If var_mes_pago = 7 Then
                  var_mes = "JULIO"
               End If
               If var_mes_pago = 8 Then
                  var_mes = "AGOSTO"
               End If
               If var_mes_pago = 9 Then
                  var_mes = "SEPTIEMBRE"
               End If
               If var_mes_pago = 10 Then
                  var_mes = "OCTUBRE"
               End If
               If var_mes_pago = 11 Then
                  var_mes = "NOVIEMBRE"
               End If
               If var_mes_pago = 12 Then
                  var_mes = "DICIEMBRE"
               End If
               
               If Len(var_año_pago_str) = 1 Then
                  var_año_pago_str = "0" + var_año_pago_str
               End If
               
               If Len(Trim(var_rfc)) = 0 Then
                  var_subimporte = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                  If Len(Trim(var_subimporte)) < 14 Then
                     For var_j = 1 + Len(Trim(var_subimporte)) To 14
                         var_subimporte = " " + var_subimporte
                     Next var_j
                  End If
                  
                  var_linea = ""
                  If Len(Trim(var_linea)) < 145 Then
                     For var_j = 1 + Len(Trim(var_linea)) To 145
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  var_linea = var_linea + var_subimporte
                  Print #1, Spc(5); var_linea
                  
                  'Print #1, ""
                  var_linea = "                                                              " + CStr(var_dia_Pago) + "   " + var_mes
                  If Len((var_linea)) < 145 Then
                     For var_j = 1 + Len((var_linea)) To 145
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  
                  var_iva = "-"
                  For var_j = 1 + Len(Trim(var_iva)) To 14
                      var_iva = " " + var_iva
                  Next var_j
                                     
                  var_linea = var_linea + var_iva
                  Print #1, Spc(5); var_linea
                  'Print #1, ""
                                     
                  var_importe = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                  If Len(Trim(var_importe)) < 14 Then
                     For var_j = 1 + Len(Trim(var_importe)) To 14
                         var_importe = " " + var_importe
                     Next var_j
                  End If
                  
                  var_linea = "" + var_año_pago_str + "          " + var_importe
                  If Len((var_linea)) < 145 Then
                     For var_j = 1 + Len((var_linea)) To 145
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  var_linea = var_linea + var_importe
                  Print #1, Spc(5); var_linea
               Else
                  var_subimporte = Format(Round(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                  
                  If Len(Trim(var_subimporte)) < 14 Then
                     For var_j = 1 + Len(Trim(var_subimporte)) To 14
                         var_subimporte = " " + var_subimporte
                     Next var_j
                  End If
                  
                  var_linea = ""
                  If Len(Trim(var_linea)) < 145 Then
                     For var_j = 1 + Len(Trim(var_linea)) To 145
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  var_linea = var_linea + var_subimporte
                  Print #1, Spc(5); var_linea
                  
                  'Print #1, ""
                  var_linea = "                                                              " + CStr(var_dia_Pago) + "   " + var_mes
                  If Len((var_linea)) < 145 Then
                     For var_j = 1 + Len((var_linea)) To 145
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  
                  var_iva = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                  For var_j = 1 + Len(Trim(var_iva)) To 14
                      var_iva = " " + var_iva
                  Next var_j
                  var_linea = var_linea + var_iva
                  Print #1, Spc(5); var_linea
                  'Print #1, ""
                  
                  
                  var_importe = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                  
                  If Len(Trim(var_importe)) < 14 Then
                     For var_j = 1 + Len(Trim(var_importe)) To 14
                         var_importe = " " + var_importe
                     Next var_j
                  End If
                  
                  var_linea = "" + var_año_pago_str + "          " + var_importe
                  If Len((var_linea)) < 145 Then
                     For var_j = 1 + Len((var_linea)) To 145
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  var_linea = var_linea + var_importe
                  Print #1, Spc(5); var_linea
               
               End If
               var_linea = ""
               Print #1, ""
               Print #1, ""
               Print #1, "      " + rs!VCHA_CLI_NOMBRE
               Print #1, "      " + rs!VCHA_CLI_DIRECCION
               Print #1, "      " + rs!vcha_ciu_nombre
               If var_empresa <> "03" Then
                  Print #1, ""
                  Print #1, ""
               Else
                  Print #1, ""
                  Print #1, ""
               End If
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Close #1
               Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt lpt1"
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
                     

End Sub





Private Sub factura_turbina()
   Dim var_nombre_unidad As String
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
   Dim ndo As New aClsNodoArbolTrazabilidad
   Dim var_establecimiento_comercial As String
   Dim var_solicitud_sigo As String
   Dim var_empresa_06 As String
   cnn.CommandTimeout = 360
                        
                     
   rsaux3.Open "select distinct inte_car_numero from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'  ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux3.EOF Then
      var_Archivo = App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat"
      Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat") For Output As #2
      While Not rsaux3.EOF
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' ORDER BY INTE_SAL_CONSECUTIVO_FACTURA", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               'AQUI EMPIEZA LA FACTURA
               Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt") For Output As #1
               'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
               'Print #1, Chr(27) + Chr(67) + Chr(44) + Chr(13)
               'Print #1, ""
               Print #1, Chr(15) + Chr(27) + Chr(64)
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               var_dia = Day(rs!dtim_Car_fecha)
               var_mes_numero = Month(rs!dtim_Car_fecha)
               If var_mes_numero = 1 Then
                  var_mes = "ENERO"
               End If
               If var_mes_numero = 2 Then
                  var_mes = "FEBRERO"
               End If
               If var_mes_numero = 3 Then
                  var_mes = "MARZO"
               End If
               If var_mes_numero = 4 Then
                  var_mes = "ABRIL"
               End If
               If var_mes_numero = 5 Then
                  var_mes = "MAYO"
               End If
               If var_mes_numero = 6 Then
                  var_mes = "JUNIO"
               End If
               If var_mes_numero = 7 Then
                  var_mes = "JULIO"
               End If
               If var_mes_numero = 8 Then
                  var_mes = "AGOSTO"
               End If
               If var_mes_numero = 9 Then
                  var_mes = "SEPTIEMBRE"
               End If
               If var_mes_numero = 10 Then
                  var_mes = "OCTUBRE"
               End If
               If var_mes_numero = 11 Then
                  var_mes = "NOVIEMBRE"
               End If
               If var_mes_numero = 12 Then
                  var_mes = "DICIEMBRE"
               End If
               
               var_año = Year(rs!dtim_Car_fecha)
               var_cadena = "                                                                                                           " + CStr(var_dia) + "   " + var_mes + "   " + CStr(var_año)
               Print #1, var_cadena
               Print #1, "    " + Str(rsaux3!inte_Car_numero)
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               var_cliente = "CLIENTE: " + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
               var_cliente_coppel = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
               var_cliente_sigo = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
               For var_j = 1 + Len(Trim(var_cliente)) To 83
                   var_cliente = var_cliente + " "
               Next var_j
               var_cliente = var_cliente
               Print #1, Spc(5); var_cliente
               var_domicilio = "DOMICILIO: " + IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " COLONIA: " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
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
               Print #1, Spc(5); var_domicilio
               var_ciudad = ""
               var_ciudad = "CIUDAD: " + Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre))
               For var_j = 1 + Len(Trim(var_ciudad)) To 37
                   var_ciudad = var_ciudad + " "
               Next var_j
               var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
               var_ciudad = var_ciudad
               var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
               If Trim(var_rfc) <> "" Then
                  var_ciudad = var_ciudad + " RFC: " + var_rfc
               Else
                  var_ciudad = var_ciudad
               End If
               
               For var_j = 1 + Len((var_ciudad)) To 83
                   var_ciudad = var_ciudad + " "
               Next var_j
                                 
                                 
               For var_j = 1 + Len(Trim(var_estado)) To 46
                   var_estado = var_estado + " "
               Next var_j
                                
   
                                  
               var_ciudad = var_ciudad + "  " + var_agente
                                 
               VAR_EMBARQUE = "EMB.: " + txt_numero_embarque
               var_ordern_surtido = x
               Print #1, Spc(5); var_ciudad
               var_rfc = "RFC:  " + var_rfc
               var_establecimiento_comercial = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
               var_rfc = "ESTADO: " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
               For var_j = 1 + Len(Trim(var_rfc)) To 70
                   var_rfc = var_rfc + " "
               Next var_j
               var_rfc = var_rfc + "               PED.: " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))) + " "
               var_rfc = var_rfc + " O.S.: " + Trim(Str(IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), 0, rs!INTE_ORS_ORDEN_SURTIDO))) + " " + VAR_EMBARQUE
               Print #1, Spc(5); var_rfc
               Print #1, ""
               Print #1, ""
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
                      If var_empresa = "15" Then
                         var_linea = var_linea + "MAQUILA DE " + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                      Else
                         var_linea = var_linea + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                      End If
                      var_i = 0
                                     
                      ''' imprimir cantidad en la orilla
                                       
                      ''' imprimir cantidad en la orilla
                                      
                                       
                      While Len((var_linea)) < 85
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
                                     
                      Print #1, Spc(5); var_linea
                      rs.MoveNext
                   Else
                      Print #1, ""
                   End If
               Next var_k
               Print #1, ""
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
               var_descuento_leyenda = 0
               var_descuento_leyenda = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
               var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
               If Len(Trim(var_linea)) < 115 Then
                  For var_j = 1 + Len(Trim(var_linea)) To 115
                      var_linea = var_linea + " "
                  Next var_j
               End If
               var_linea = var_linea + var_importe_descuento_1_str
               Print #1, Spc(5); var_linea
               
               var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%"
               If Len(Trim(var_linea)) < 115 Then
                  For var_j = 1 + Len(Trim(var_linea)) To 115
                      var_linea = var_linea + " "
                  Next var_j
               End If
               var_linea = var_linea + var_importe_descuento_2_str
               Print #1, Spc(5); var_linea
               
               var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
               Print #1, ""
                
               If Len(Trim(var_rfc)) = 0 Then
                  var_subimporte = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                  If Len(Trim(var_subimporte)) < 14 Then
                     For var_j = 1 + Len(Trim(var_subimporte)) To 14
                         var_subimporte = " " + var_subimporte
                     Next var_j
                  End If
                  
                  var_linea = ""
                  If Len(Trim(var_linea)) < 115 Then
                     For var_j = 1 + Len(Trim(var_linea)) To 115
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  var_linea = var_linea + var_subimporte
                  Print #1, Spc(5); var_linea
                  
                  Print #1, ""
                  var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                  If Len(Trim(var_linea)) < 115 Then
                     For var_j = 1 + Len(Trim(var_linea)) To 115
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  
                  var_iva = "-"
                  For var_j = 1 + Len(Trim(var_iva)) To 14
                      var_iva = " " + var_iva
                  Next var_j
                                     
                  var_linea = var_linea + var_iva
                  Print #1, Spc(5); var_linea
                  Print #1, ""
                                     
                  var_importe = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                  
                  If Len(Trim(var_importe)) < 14 Then
                     For var_j = 1 + Len(Trim(var_importe)) To 14
                         var_importe = " " + var_importe
                     Next var_j
                  End If
                  var_linea = ""
                  If Len(Trim(var_linea)) < 115 Then
                     For var_j = 1 + Len(Trim(var_linea)) To 115
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  var_linea = var_linea + var_importe
                  Print #1, Spc(5); var_linea
               
               
               
               Else
                  var_subimporte = Format(Round(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                  
                  If Len(Trim(var_subimporte)) < 14 Then
                     For var_j = 1 + Len(Trim(var_subimporte)) To 14
                         var_subimporte = " " + var_subimporte
                     Next var_j
                  End If
                  
                  var_linea = ""
                  If Len(Trim(var_linea)) < 115 Then
                     For var_j = 1 + Len(Trim(var_linea)) To 115
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  var_linea = var_linea + var_subimporte
                  Print #1, Spc(5); var_linea
                  
                  Print #1, ""
                  var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                  If Len(Trim(var_linea)) < 115 Then
                     For var_j = 1 + Len(Trim(var_linea)) To 115
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  
                  var_iva = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                  For var_j = 1 + Len(Trim(var_iva)) To 14
                      var_iva = " " + var_iva
                  Next var_j
                  var_linea = var_linea + var_iva
                  Print #1, Spc(5); var_linea
                  Print #1, ""
                  
                  
                  var_importe = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                  
                  If Len(Trim(var_importe)) < 14 Then
                     For var_j = 1 + Len(Trim(var_importe)) To 14
                         var_importe = " " + var_importe
                     Next var_j
                  End If
                  var_linea = ""
                  If Len(Trim(var_linea)) < 115 Then
                     For var_j = 1 + Len(Trim(var_linea)) To 115
                         var_linea = var_linea + " "
                     Next var_j
                  End If
                  var_linea = var_linea + var_importe
                  Print #1, Spc(5); var_linea
               
               End If
               var_linea = ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Close #1
               Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt lpt1"
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

Private Sub cmd_correo_clientes_Click()
   Me.frm_correo_clientes.Visible = True
   Me.txt_embarque_correo_clientes = ""
   Me.txt_embarque_correo_clientes.SetFocus
End Sub

Private Sub cmd_correo_facturacion_tiendas_Click()
   frm_embarque_correo_ft.Visible = True
   txt_embarque_correo_ft = ""
   txt_embarque_correo_ft.SetFocus
End Sub

Private Sub cmd_embarques_cerrados_Click()
On Error GoTo salir:
      rsaux2.Open "select * from VW_EMBARQUES_CERRADOS_NO_FACTURADOS", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux2.EOF Then
         'Me.Enabled = False
         var_activa_forma_detalle_cajas = Me.Name
         frmembarques_cerrados_no_facturados.Show 1
      Else
         MsgBox "No existen embarques cerrados sin facturar", vbOKOnly, "ATENCION"
      End If
      rsaux2.Close
      Exit Sub
salir:
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   MsgBox "No se pueden ver los embarques cerrados", vbOKOnly, "ATENCION"
End Sub

Private Sub cmd_factura_electronica_Click()
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
   Dim var_nombre_unidad As String
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
   Dim ndo As New aClsNodoArbolTrazabilidad
   Dim var_establecimiento_comercial As String
   Dim var_solicitud_sigo As String
   Dim var_empresa_06 As String
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
   If rsaux10.State = 1 Then
      rsaux10.Close
   End If
   If rsaux11.State = 1 Then
      rsaux11.Close
   End If
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
            var_si_saldo_oracle = 0
            If var_unidad_organizacional = "23" Then
               var_cadena = "SELECT  dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_PEDIDO_CREDITO, dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID FROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN"
               var_cadena = var_cadena + " dbo.TB_ENC_ORDEN_SURTIDO ON dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN = dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO INNER JOIN dbo.TB_ENCABEZADO_PEDIDOS ON  dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO = dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO WHERE (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_numero_embarque + ") AND (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')    "
               rsaux9.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               var_pedido_credito_POSIBLE = 0
               If Not rsaux9.EOF Then
                  var_pedido_credito_POSIBLE = IIf(IsNull(rsaux9!inte_ped_pedido_credito), 0, rsaux9!inte_ped_pedido_credito)
               Else
                  var_pedido_credito_POSIBLE = 0
               End If
               rsaux9.Close
               var_cadena = "SELECT ROUND(SUM(((dbo.TB_SALIDAS.FLOA_SAL_PRECIO * 1.16 * dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD) * (1 - dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1 / 100)) * (1 - dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2 / 100)), 2) AS IMPORTE_EMBARQUE, dbo.TB_CLIENTES.VCHA_CLI_REFERENCIA FROM dbo.TB_CLIENTES INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_SALIDAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO ON"
               var_cadena = var_cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO WHERE (dbo.TB_DETALLE_EMBARQUES.inte_emb_embarque = " + Me.txt_numero_embarque + ") AND (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') GROUP BY dbo.TB_CLIENTES.VCHA_CLI_REFERENCIA"
               'MsgBox var_cadena
               rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_importe_embarque = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                  rsaux1.Open "select NUMB_SAL_IMPORTE   from tb_saldo where vcha_sal_Referencia = '" + Trim(IIf(IsNull(rsaux!VCHA_CLI_REFERENCIA), "", rsaux!VCHA_CLI_REFERENCIA)) + "'", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                     var_importe_Saldo_oracle = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
                  Else
                     var_importe_Saldo_oracle = 0
                  End If
                  rsaux1.Close
               Else
                  var_importe_embarque = 0
               End If
               rsaux.Close
               If Round(var_importe_embarque, 2) <= Round(var_importe_Saldo_oracle, 2) Then
                  var_si_saldo_oracle = 0
               Else
                  If var_pedido_credito_POSIBLE = 1 Then
                     var_si_saldo_oracle = 0
                  Else
                     var_si_saldo_oracle = 1
                  End If
               End If
            Else
               var_si_saldo_oracle = 0
            End If
            If var_si_saldo_oracle = 0 Then
               Cadena = "SELECT     dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID, "
               Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO, dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID,"
               Cadena = Cadena + " dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID , dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD"
               Cadena = Cadena + " FROM         dbo.TB_DETALLE_EMBARQUES INNER JOIN"
               Cadena = Cadena + " dbo.TB_SALIDAS WITH (NOLOCK) ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND"
               Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND"
               Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO AND"
               Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID "
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
                              'MsgBox "execute factura_embarques '" + var_empresa + "', '" + var_unidad_organizacional + "', " + txt_numero_embarque + ", '', '','" + var_serie + "', 'FA'"
                              var_clave_movimiento_vistas = ""
                              If var_empresa = "18" Then
                                 rs.Open "select top 1 * from tb_Detalle_embarques where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_numero_embarque, cnn, adOpenDynamic, adLockOptimistic
                                 var_clave_movimiento_vistas = rs!VCHA_MOV_MOVIMIENTO_ID
                                 rs.Close
                              End If
                              If var_empresa = "18" Then
                                 If var_clave_movimiento_vistas = "FV" Then
                                    rs.Open "execute factura_embarques_vistas '" + var_empresa + "', '" + var_unidad_organizacional + "', " + txt_numero_embarque + ", '', '','" + var_serie + "', 'FA'", cnn, adOpenDynamic, adLockOptimistic
                                 Else
                                    rs.Open "execute FACTURA_EMBARQUES_ELECTRONICO '" + var_empresa + "', '" + var_unidad_organizacional + "', " + txt_numero_embarque + ", '', '','" + var_serie + "', 'FA'", cnn, adOpenDynamic, adLockOptimistic
                                 End If
                              Else
                                 If var_empresa = "02" Or var_empresa = "03" Then
                                    rs.Open "execute FACTURA_EMBARQUES '" + var_empresa + "', '" + var_unidad_organizacional + "', " + txt_numero_embarque + ", '', '','" + var_serie + "', 'FA'", cnn, adOpenDynamic, adLockOptimistic
                                 Else
                                    rs.Open "execute FACTURA_EMBARQUES_ELECTRONICO '" + var_empresa + "', '" + var_unidad_organizacional + "', " + txt_numero_embarque + ", '', '','" + var_serie + "', 'FA'", cnn, adOpenDynamic, adLockOptimistic
                                 End If
                              End If
                              ok = TB_ENC_EMBARQUE_M.Anadir(var_empresa, var_unidad_organizacional, CDbl(txt_numero_embarque), "F")
                              rsaux5.Open "select * from tb_detalle_embarques where inte_emb_embarque = " + Me.txt_numero_embarque + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                              var_si_correo_ft = 0
                              var_leyenda_sorteo = ""
                              If var_empresa <> "18" Then
                                 While Not rsaux5.EOF
                                       If rsaux5!VCHA_MOV_MOVIMIENTO_ID = "FT" Then
                                          var_si_correo_ft = 1
                                          cnn.BeginTrans
                                          rsaux10.Open "select inte_pri_activar_sorteo from tb_principal", cnn, adOpenDynamic, adLockOptimistic
                                          var_activar_sorteo = 0
                                          If Not rsaux10.EOF Then
                                             var_activar_sorteo = IIf(IsNull(rsaux10!inte_pri_activar_sorteo), 0, rsaux10!inte_pri_activar_sorteo)
                                          Else
                                             var_activar_sorteo = 0
                                          End If
                                          rsaux10.Close
                                          If var_activar_sorteo = 1 Then
                                             rsaux10.Open "SELECT * FROM VW_SORTEO_NUMERO_BOLETOS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + rsaux5!VCHA_MOV_MOVIMIENTO_ID + "' AND INTE_EMO_NUMERO = " + CStr(rsaux5!INTE_SAL_NUMERO) + " AND VCHA_UOR_UNIDAD_ID = '" + rsaux5!VCHA_UOR_UNIDAD_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                             rsaux11.Open "SELECT * FROM TB_SORTEO_FOLIOS", cnn, adOpenDynamic, adLockOptimistic
                                             VAR_SORTEO_INICIO = rsaux11!inte_sor_folio_actual
                                             var_numero_boletos = rsaux10!numero_boletos
                                             rsaux11.Close
                                             If VAR_SORTEO_INICIO > VAR_SORTEO_INICIO + rsaux10!numero_boletos - 1 Then
                                                var_leyenda_sorteo = ""
                                             Else
                                                var_leyenda_sorteo = "     Folios participantes: Del " + CStr(VAR_SORTEO_INICIO) + " al " + CStr(VAR_SORTEO_INICIO + rsaux10!numero_boletos - 1)
                                             End If
                                             rsaux11.Open "INSERT INTO TB_SORTEO_BOLETOS_MOVIMIENTO (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, FLOA_SOR_IMPORTE, INTE_SOR_NUMERO_BOLETOS, INTE_SOR_BOLETO_INICIO, INTE_SOR_BOLETO_FINAL) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + rsaux5!VCHA_ALM_ALMACEN_ID + "', '" + rsaux5!VCHA_MOV_MOVIMIENTO_ID + "'," + CStr(rsaux5!INTE_SAL_NUMERO) + ", " + CStr(rsaux10!importe_neto) + "," + CStr(rsaux10!numero_boletos) + "," + CStr(VAR_SORTEO_INICIO) + "," + CStr(VAR_SORTEO_INICIO + rsaux10!numero_boletos - 1) + " )", cnn, adOpenDynamic, adLockOptimistic
                                             rsaux11.Open "UPDATE TB_SORTEO_FOLIOS SET INTE_SOR_FOLIO_ACTUAL = INTE_SOR_FOLIO_ACTUAL + " + CStr(var_numero_boletos), cnn, adOpenDynamic, adLockOptimistic
                                             rsaux10.Close
                                          End If
                                          rsaux6.Open "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id = 'FT' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_emo_numero = " + CStr(rsaux5!INTE_SAL_NUMERO), cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux6.EOF Then
                                             var_numero_orden_surtido = IIf(IsNull(rsaux6!inte_emo_numero_origen), 0, rsaux6!inte_emo_numero_origen)
                                             If rsaux9.State = 1 Then
                                                rsaux9.Close
                                             End If
                                             rsaux9.Open "SELECT * FROM VW_PEDIDOS_CREDITO_TIENDAS WHERE INTE_ORS_ORDEN_SURTIDO = " + CStr(var_numero_orden_surtido), cnn, adOpenDynamic, adLockOptimistic
                                             var_pedido_credito = 0
                                             If Not rsaux9.EOF Then
                                                var_pedido_credito = IIf(IsNull(rsaux9!inte_ped_pedido_credito), 0, rsaux9!inte_ped_pedido_credito)
                                             End If
                                             If var_pedido_credito = 0 Then
                                                rsaux7.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id = 'FT' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_emo_numero = " + CStr(rsaux6!inte_emo_numero), cnn, adOpenDynamic, adLockOptimistic
                                                var_importe_pedido_tienda = IIf(IsNull(rsaux7!floa_Car_importe_neto), 0, rsaux7!floa_Car_importe_neto) / IIf(IsNull(rsaux7!floa_car_tipo_cambio), 1, rsaux7!floa_car_tipo_cambio)
                                                var_numero_factura_tienda = IIf(IsNull(rsaux7!inte_Car_numero), 0, rsaux7!inte_Car_numero)
                                                var_tipo_Cambio_tienda = IIf(IsNull(rsaux7!floa_car_tipo_cambio), 1, rsaux7!floa_car_tipo_cambio)
                                                var_importe_descuento_1_tienda = IIf(IsNull(rsaux7!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rsaux7!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
                                                var_importe_descuento_2_tienda = IIf(IsNull(rsaux7!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rsaux7!FLOA_CAR_PORCENTAJE_DESCUENTO_2)
                                                rsaux7.Close
                                                rsaux7.Open "select * from vw_clientes where vcha_cli_clave_id = '" + rsaux6!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                                If Not rsaux7.EOF Then
                                                   var_referencia_cliente_tienda = IIf(IsNull(rsaux7!VCHA_CLI_REFERENCIA), "", rsaux7!VCHA_CLI_REFERENCIA)
                                                   var_clave_cliente_tienda = IIf(IsNull(rsaux7!vcha_cli_clave_id), "", rsaux7!vcha_cli_clave_id)
                                                   var_agente_cliente_tienda = IIf(IsNull(rsaux7!VCHA_AGE_AGENTE_ID), "", rsaux7!VCHA_AGE_AGENTE_ID)
                                                   var_canal_cliente_tienda = IIf(IsNull(rsaux7!vcha_can_canal_venta_id), "", rsaux7!vcha_can_canal_venta_id)
                                                   var_grupo_real_tienda = IIf(IsNull(rsaux7!vcha_gre_grupo_real_id), "", rsaux7!vcha_gre_grupo_real_id)
                                                   var_grupo_actual_tienda = IIf(IsNull(rsaux7!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux7!VCHA_GAC_GRUPO_aCTUAL_ID)
                                                   var_titular_tienda = IIf(IsNull(rsaux7!vcha_tit_titular_id), "", rsaux7!vcha_tit_titular_id)
                                                   var_porcentaje_iva_tienda = IIf(IsNull(rsaux7!FLOA_TPE_IVA), "", rsaux7!FLOA_TPE_IVA)
                                                   var_clave_moneda_tienda = IIf(IsNull(rsaux7!vcha_mon_moneda_id), "1", rsaux7!vcha_mon_moneda_id)
                                                End If
                                                rsaux7.Close
                                                If rsaux8.State = 1 Then
                                                   rsaux8.Close
                                                End If
                                                rsaux8.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id = 'FT' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_emo_numero = " + CStr(rsaux6!inte_emo_numero) + " and (char_car_estatus <> 'C' or char_car_estatus is null)", cnn, adOpenDynamic, adLockOptimistic
                                                var_importe_facturado_ft = 0
                                                While Not rsaux8.EOF
                                                      rsaux11.Open "select * from TB_MAXIMO_PAGO", cnn_sid_quezada, adOpenDynamic, adLockOptimistic
                                                      If rsaux11.EOF Then
                                                         var_numero_folio = 0
                                                      Else
                                                         var_numero_folio = IIf(IsNull(rsaux11!inte_max_maximo_pago), 0, rsaux11!inte_max_maximo_pago)
                                                      End If
                                                      rsaux11.Close
                                                      var_numero_folio = var_numero_folio + 1
                                                      rsaux11.Open "update TB_MAXIMO_PAGO set inte_max_maximo_pago = inte_max_maximo_pago + 1", cnn_sid_quezada, adOpenDynamic, adLockOptimistic
                                                      Cadena = "INSERT INTO TB_ENCABEZADO_CARTERA (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, "
                                                      Cadena = Cadena + "FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3, FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, VCHA_SER_SERIE_ID, CHAR_CAR_ESTATUS, VCHA_CAR_CHEQUE_DEPOSITO, VCHA_CAR_BANCO_CHEQUE, VCHA_CAR_DEPOSITO, VCHA_CAR_BANCO_DEPOSITO, DTIM_CAR_FECHA_DEPOSITO) values ("
                                                      Cadena = Cadena + "'" + var_empresa + "', '" + var_unidad_organizacional + "', 'PA', 'PA', 'PA', " + CStr(var_numero_folio) + ", '-', '', '', 0, getdate(), '" + var_agente_cliente_tienda + "', '" + var_grupo_actual_tienda + "', '" + var_grupo_real_tienda + "', '" + var_titular_tienda + "', '" + var_clave_cliente_tienda + "', '', 0, " + CStr(var_porcentaje_iva_tienda) + ", 0, 0, " + CStr(var_importe_descuento_1_tienda) + ", " + CStr(var_importe_descuento_2_tienda) + ", 0, " + CStr(rsaux8!floa_Car_importe_neto) + ", " + CStr(rsaux8!floa_Car_importe_neto - rsaux8!floa_car_subimporte) + ", 0, 0, 0, 0, 0, " + CStr(rsaux8!floa_car_subimporte) + ", " + CStr(rsaux8!floa_Car_importe_neto) + ", '', '"
                                                      Cadena = Cadena + CStr(var_clave_usuario_global) + "', '', getdate(), 0, getdate(), getdate(), '" + var_clave_moneda + "', " + CStr(var_tipo_Cambio_tienda) + ", 'FT', 'I','', '', '','','')"
                                                      rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                                      rsaux7.Open "update tb_encabezado_cartera set inte_car_pedido_credito = 0 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Car_documento = 'FA' and vcha_Ser_serie_id = '" + var_serie + "' and inte_Car_numero = " + CStr(CDbl(rsaux8!inte_Car_numero)), cnn, adOpenDynamic, adLockOptimistic
                                                      Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
                                                      var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, var_serie, "FA", CDbl(rsaux8!inte_Car_numero), "FT", "PA", CDbl(var_numero_folio), 0, CDbl(rsaux8!floa_Car_importe_neto))
                                                      var_importe_facturado_ft = var_importe_facturado_ft + rsaux8!floa_Car_importe_neto
                                                      If CDbl(Round(rsaux8!floa_Car_importe_neto, 2)) > 0 Then
                                                         rsaux7.Open "select NUMB_SAL_IMPORTE   from tb_saldo where vcha_sal_Referencia = '" + Trim(var_referencia_cliente_tienda) + "'", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                                                         var_importe_Saldo_oracle = IIf(IsNull(rsaux7(0).Value), 0, rsaux7(0).Value)
                                                         var_importe_factura = Round(CDbl(rsaux8!floa_Car_importe_neto), 2)
                                                         var_diferencia_saldo_factura = var_importe_Saldo_oracle - var_importe_factura
                                                         If var_diferencia_saldo_factura < 0 And var_diferencia_saldo_factura > -10 Then
                                                            var_importe_factura = var_importe_factura + (var_diferencia_saldo_factura)
                                                         End If
                                                         rsaux7.Close
                                                         rsaux7.Open "call SP_AGREGA_CARGO ('" + var_canal_cliente_tienda + "','" + var_agente_cliente_tienda + "', " + CStr(CDbl(rsaux8!inte_Car_numero)) + ",'" + Trim(var_referencia_cliente_tienda) + "'," + CStr(CDbl(Round(CDbl(var_importe_factura), 2))) + ", 0,TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'VA')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                                                         VAR_ADI = ""
                                                         VAR_UNFO = 0
                                                         rsaux7.Open "SELECT  dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_ADI, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_UNFO FROM dbo.TB_ENC_ORDEN_SURTIDO INNER JOIN dbo.TB_ENCABEZADO_PEDIDOS ON dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO = dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO Where (dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = " + CStr(rsaux6!inte_emo_numero_origen) + ")", cnn, adOpenDynamic, adLockOptimistic
                                                         If Not rsaux7.EOF Then
                                                            VAR_ADI = IIf(IsNull(rsaux7(0).Value), "", rsaux7(0).Value)
                                                            VAR_UNFO = IIf(IsNull(rsaux7(1).Value), 0, rsaux7(1).Value)
                                                         Else
                                                            VAR_ADI = ""
                                                            VAR_UNFO = 0
                                                         End If
                                                         rsaux7.Close
                                                         rsaux7.Open "SELECT SUBSTRING(VCHA_ART_ARTICULO_ID,7,5) AS ARTICULO,((FLOA_SAL_PRECIO * (1 - (FLOA_SAL_DESCUENTO_1/100))) * (1- (FLOA_SAL_DESCUENTO_2/100))) * FLOA_SAL_cANTIDAD AS IMPORTE  FROM TB_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_sER_SERIE_ID = '" + var_serie + "' AND INTE_cAR_NUMERO = " + CStr(rsaux8!inte_Car_numero), cnn, adOpenDynamic, adLockOptimistic
                                                         If VAR_ADI <> "" Then
                                                            IMPORTE_PUNTOS = 0
                                                            While Not rsaux7.EOF
                                                                  If rsaux10.State = 1 Then
                                                                     rsaux10.Close
                                                                  End If
                                                                  rsaux10.Open "SELECT PORCIENTO_DINERO FROM MONEDERO_ARTICULOS WHERE UNIDAD_NEGOCIO = 1 AND ARTICULO_ID = '" + rsaux7!ARTICULO + "' AND TIPO_CLIENTE = " + VAR_ADI, cnn_puntos_monedero, adOpenDynamic, adLockOptimistic
                                                                  'MsgBox "SELECT PORCIENTO_DINERO FROM MONEDERO_ARTICULOS WHERE UNIDAD_NEGOCIO = 1 AND ARTICULO_ID = '" + rsaux7!ARTICULO + "' AND TIPO_CLIENTE = " + VAR_ADI
                                                                  If Not rsaux10.EOF Then
                                                                     IMPORTE_PUNTOS = IMPORTE_PUNTOS + rsaux7!Importe * ((rsaux10(0).Value) / 100)
                                                                  End If
                                                                  rsaux10.Close
                                                                  rsaux7.MoveNext
                                                            Wend
                                                            'MsgBox "CALL SP_CARGO_ABONO_MONEDERO (0, 1, 3," + CStr(VAR_UNFO) + ",100, 1, 'FACTURA','" + var_serie + CStr(rsaux8!inte_Car_numero) + "',0," + CStr(IMPORTE_PUNTOS) + ",'" + fun_NombrePc + "')"
                                                            rsaux10.Open "CALL FC_CARGO_ABONO_MONEDERO (1, 3," + CStr(VAR_UNFO) + ",100, 1, 'FACTURA','" + var_serie + CStr(rsaux8!inte_Car_numero) + "',0," + CStr(IMPORTE_PUNTOS) + ",'" + fun_NombrePc + "')", cnn_puntos_monedero, adOpenDynamic, adLockOptimistic
                                                          End If
                                                          rsaux7.Close
                                                      End If
                                             
                                                      rsaux8.MoveNext
                                                Wend
                                                If rsaux8.State = 1 Then
                                                   rsaux8.Close
                                                End If
                                                rsaux4.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + CStr(rsaux6!inte_emo_numero_origen), cnn, adOpenDynamic, adLockOptimistic
                                                If Not rsaux4.EOF Then
                                                   var_pedido_tienda = rsaux4!inte_ped_numero
                                                   rsaux4.Close
                                                   rsaux4.Open "select * from VW_IMPORTES_SEGURO_PAQUETERIA where inte_ped_numero = " + CStr(var_pedido_tienda), cnn, adOpenDynamic, adLockOptimistic
                                                   If Not rsaux4.EOF Then
                                                      var_importe_paqueteria_tienda = IIf(IsNull(rsaux4!importe_seguro), 0, rsaux4!importe_seguro)
                                                      var_importe_seguro_tienda = IIf(IsNull(rsaux4!importe_paqueteria), 0, rsaux4!importe_paqueteria)
                                                      var_importe_referencia_tienda = IIf(IsNull(rsaux4!floa_paq_costo_referencia), 0, rsaux4!floa_paq_costo_referencia)
                                                      var_importe_pedido_tienda = IIf(IsNull(rsaux4!importe_pedido), 0, rsaux4!importe_pedido)
                                                      var_importe_total_tienda = var_importe_pedido_tienda + var_importe_paqueteria_tienda + var_importe_seguro_tienda + var_importe_referencia_tienda
                                                      rsaux4.MoveNext
                                                   End If
                                                   rsaux4.Close
                                                Else
                                                   rsuax4.Close
                                                End If
                                                If CDbl(Round(var_importe_total_tienda, 2)) > CDbl(Round(var_importe_facturado_ft, 2)) Then
                                                   var_diferencia_facturado = var_importe_total_tienda - var_importe_facturado_ft
                                                   rsaux8.Open "CALL SP_AGREGA_ABONO('" + Trim(var_referencia_cliente_tienda) + "',0.00," + CStr(var_diferencia_facturado) + ",SYSDATE,SYSDATE,'" + CStr(var_pedido_tienda) + "','','DF','')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                                                End If
                                             End If
                                          Else
                                          End If
                                          rsaux6.Close
                                          cnn.CommitTrans
                                       End If
                                       rsaux5.MoveNext
                                 Wend
                              End If
                              rsaux5.Close
                              Call envio_tb_transito
                              'aqui empieza el archivo de la factura electronica
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
                              cnn.CommandTimeout = 360
                              If var_empresa = "18" Then
                                 Cadena = "EXEC SP_CREA_TABLA_FACTURAS " + CStr(var_consecutivo) + ",'" + var_empresa + "'," + Me.txt_numero_embarque
                              Else
                                 Cadena = "EXEC SP_CREA_TABLA_FACTURAS_nuevo " + CStr(var_consecutivo) + ",'" + var_empresa + "'," + txt_numero_embarque
                              End If
                              rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                              Me.frm_mensaje.Visible = False
                              rsaux3.Open "select distinct inte_car_numero from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'  ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 While Not rsaux3.EOF
                                       Call crea_factura_electronica
                                       rsaux3.MoveNext
                                 Wend
                              End If
                              
                              'rsaux3.MoveFirst
                              'While Not rsaux3.EOF
                              '      If var_empresa = "15" Then
                              '         var_ruta_factura_pdf = "\\FACELECTRONICA\fefiles\ConectorEE\envio\enviados\" + Trim(var_serie) + Trim(Str(rsaux3!inte_car_numero)) + ".pdf"
                              '         frmpdf.Show 1
                              '      End If
                              '      rsaux3.MoveNext
                              'Wend
                              'rsaux3.Close
                              
                              
                              
                              
                              
                              'aqui termina el archivo de la factura electronica
                              
                              
                              
                              '''' AQUI DEBE DE IR EL CORREO DE LAS VENTAS DE TIENDAS
                              If var_si_correo_ft = 1 Then
                                 If IsNumeric(Me.txt_embarque_correo_ft) Then
                                    If rs.State = 1 Then
                                       rs.Close
                                    End If
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
                                          If var_cliente_coppel = "C000006202" Then
                                             Print #1, ""
                                          Else
                                             Print #1, "   Colonia:   " + IIf(IsNull(rsaux8!vcha_col_nombre), "", rsaux8!vcha_col_nombre)
                                          End If
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
                                       var_importe_total_str = Format(var_importe_total, "###,###,##0.00")
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
                              End If
                              ''''' hasta aqui termina el correo de ventas de tiendas
                              MsgBox "Se a terminado el proceso de facturación", vbOKOnly, "ATENCION"
                              var_estatus_embarque = "F"
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
               'fin de la comparacion del saldo con el oracle
               End If
            Else
                MsgBox "El saldo del cliente en ORACLE es menor al importe de las facturas, ", vbOKOnly, "ATENCION"
            End If
         Else
           'pregunta si sorteo
            MsgBox "Se a cancelado la impresión de las facturas", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "No se a seleccionado un embarque", vbOKOnly, "ATENCION"
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
   If rsaux11.State = 1 Then
      rsaux11.Close
   End If
End Sub

Private Sub cmd_imprimir_Click()
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
   Dim var_nombre_unidad As String
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
   Dim ndo As New aClsNodoArbolTrazabilidad
   Dim var_establecimiento_comercial As String
   Dim var_solicitud_sigo As String
   Dim var_empresa_06 As String
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
   If rsaux10.State = 1 Then
      rsaux10.Close
   End If
   If rsaux11.State = 1 Then
      rsaux11.Close
   End If
   
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
            var_si_saldo_oracle = 0
            If var_unidad_organizacional = "23" Then
               var_cadena = "SELECT  dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_PEDIDO_CREDITO, dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID FROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN"
               var_cadena = var_cadena + " dbo.TB_ENC_ORDEN_SURTIDO ON dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN = dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO INNER JOIN dbo.TB_ENCABEZADO_PEDIDOS ON  dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO = dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO WHERE (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_numero_embarque + ") AND (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')    "
               rsaux9.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               var_pedido_credito_POSIBLE = 0
               If Not rsaux9.EOF Then
                  var_pedido_credito_POSIBLE = IIf(IsNull(rsaux9!inte_ped_pedido_credito), 0, rsaux9!inte_ped_pedido_credito)
               Else
                  var_pedido_credito_POSIBLE = 0
               End If
               rsaux9.Close
               
               
               
            
            
               var_cadena = "SELECT ROUND(SUM(((dbo.TB_SALIDAS.FLOA_SAL_PRECIO * 1.16 * dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD) * (1 - dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1 / 100)) * (1 - dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2 / 100)), 2) AS IMPORTE_EMBARQUE, dbo.TB_CLIENTES.VCHA_CLI_REFERENCIA FROM dbo.TB_CLIENTES INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_SALIDAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO ON"
               var_cadena = var_cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO WHERE (dbo.TB_DETALLE_EMBARQUES.inte_emb_embarque = " + Me.txt_numero_embarque + ") AND (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') GROUP BY dbo.TB_CLIENTES.VCHA_CLI_REFERENCIA"
               'MsgBox var_cadena
               rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_importe_embarque = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                  rsaux1.Open "select NUMB_SAL_IMPORTE   from tb_saldo where vcha_sal_Referencia = '" + Trim(IIf(IsNull(rsaux!VCHA_CLI_REFERENCIA), "", rsaux!VCHA_CLI_REFERENCIA)) + "'", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                     var_importe_Saldo_oracle = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
                  Else
                     var_importe_Saldo_oracle = 0
                  End If
                  rsaux1.Close
               Else
                  var_importe_embarque = 0
               End If
               rsaux.Close
               'MsgBox Round(var_importe_embarque, 2)
               'MsgBox Round(var_importe_Saldo_oracle, 2)
               If Round(var_importe_embarque, 2) <= Round(var_importe_Saldo_oracle, 2) Then
                  var_si_saldo_oracle = 0
               Else
                  If var_pedido_credito_POSIBLE = 1 Then
                     var_si_saldo_oracle = 0
                  Else
                     var_si_saldo_oracle = 1
                  End If
               End If
            Else
               var_si_saldo_oracle = 0
            End If
            If var_si_saldo_oracle = 0 Then
         Cadena = "SELECT     dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID, "
         Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO, dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID,"
         Cadena = Cadena + " dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID , dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD"
         Cadena = Cadena + " FROM         dbo.TB_DETALLE_EMBARQUES INNER JOIN"
         Cadena = Cadena + " dbo.TB_SALIDAS WITH (NOLOCK) ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND"
         Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND"
         Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO AND"
         Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID "
         Cadena = Cadena + " WHERE     (dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD IS NULL) AND (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + txt_numero_embarque + ") AND"
         Cadena = Cadena + " (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')"
         'rsaux4.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         rsaux4.Open "select * from tb_lineas", cnn, adOpenDynamic, adLockOptimistic
         'If Not rsaux4.EOF Then
         If rsaux4.EOF Then
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
                     'MsgBox "execute factura_embarques '" + var_empresa + "', '" + var_unidad_organizacional + "', " + txt_numero_embarque + ", '', '','" + var_serie + "', 'FA'"
                     var_clave_movimiento_vistas = ""
                     If var_empresa = "18" Then
                        rs.Open "select top 1 * from tb_Detalle_embarques where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_numero_embarque, cnn, adOpenDynamic, adLockOptimistic
                        var_clave_movimiento_vistas = rs!VCHA_MOV_MOVIMIENTO_ID
                        rs.Close
                     End If
                     If var_empresa = "18" Then
                        If var_clave_movimiento_vistas = "FV" Then
                           rs.Open "execute factura_embarques_vistas '" + var_empresa + "', '" + var_unidad_organizacional + "', " + txt_numero_embarque + ", '', '','" + var_serie + "', 'FA'", cnn, adOpenDynamic, adLockOptimistic
                        Else
                           rs.Open "execute factura_embarques '" + var_empresa + "', '" + var_unidad_organizacional + "', " + txt_numero_embarque + ", '', '','" + var_serie + "', 'FA'", cnn, adOpenDynamic, adLockOptimistic
                        End If
                     Else
                        rs.Open "execute factura_embarques '" + var_empresa + "', '" + var_unidad_organizacional + "', " + txt_numero_embarque + ", '', '','" + var_serie + "', 'FA'", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     ok = TB_ENC_EMBARQUE_M.Anadir(var_empresa, var_unidad_organizacional, CDbl(txt_numero_embarque), "F")
                     rsaux5.Open "select * from tb_detalle_embarques where inte_emb_embarque = " + Me.txt_numero_embarque + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_si_correo_ft = 0
                     var_leyenda_sorteo = ""
                     If var_empresa <> "18" Then
                     While Not rsaux5.EOF
                           If rsaux5!VCHA_MOV_MOVIMIENTO_ID = "FT" Then
                              var_si_correo_ft = 1
                              cnn.BeginTrans
                              rsaux10.Open "select inte_pri_activar_sorteo from tb_principal", cnn, adOpenDynamic, adLockOptimistic
                              var_activar_sorteo = 0
                              If Not rsaux10.EOF Then
                                 var_activar_sorteo = IIf(IsNull(rsaux10!inte_pri_activar_sorteo), 0, rsaux10!inte_pri_activar_sorteo)
                              Else
                                 var_activar_sorteo = 0
                              End If
                              rsaux10.Close
                              If var_activar_sorteo = 1 Then
                                 rsaux10.Open "SELECT * FROM VW_SORTEO_NUMERO_BOLETOS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + rsaux5!VCHA_MOV_MOVIMIENTO_ID + "' AND INTE_EMO_NUMERO = " + CStr(rsaux5!INTE_SAL_NUMERO) + " AND VCHA_UOR_UNIDAD_ID = '" + rsaux5!VCHA_UOR_UNIDAD_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                 rsaux11.Open "SELECT * FROM TB_SORTEO_FOLIOS", cnn, adOpenDynamic, adLockOptimistic
                                 VAR_SORTEO_INICIO = rsaux11!inte_sor_folio_actual
                                 var_numero_boletos = rsaux10!numero_boletos
                                 rsaux11.Close
                                 If VAR_SORTEO_INICIO > VAR_SORTEO_INICIO + rsaux10!numero_boletos - 1 Then
                                    var_leyenda_sorteo = ""
                                 Else
                                    var_leyenda_sorteo = "     Folios participantes: Del " + CStr(VAR_SORTEO_INICIO) + " al " + CStr(VAR_SORTEO_INICIO + rsaux10!numero_boletos - 1)
                                 End If
                                 rsaux11.Open "INSERT INTO TB_SORTEO_BOLETOS_MOVIMIENTO (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, FLOA_SOR_IMPORTE, INTE_SOR_NUMERO_BOLETOS, INTE_SOR_BOLETO_INICIO, INTE_SOR_BOLETO_FINAL) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + rsaux5!VCHA_ALM_ALMACEN_ID + "', '" + rsaux5!VCHA_MOV_MOVIMIENTO_ID + "'," + CStr(rsaux5!INTE_SAL_NUMERO) + ", " + CStr(rsaux10!importe_neto) + "," + CStr(rsaux10!numero_boletos) + "," + CStr(VAR_SORTEO_INICIO) + "," + CStr(VAR_SORTEO_INICIO + rsaux10!numero_boletos - 1) + " )", cnn, adOpenDynamic, adLockOptimistic
                                 rsaux11.Open "UPDATE TB_SORTEO_FOLIOS SET INTE_SOR_FOLIO_ACTUAL = INTE_SOR_FOLIO_ACTUAL + " + CStr(var_numero_boletos), cnn, adOpenDynamic, adLockOptimistic
                                 rsaux10.Close
                              End If
                              
                              rsaux6.Open "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id = 'FT' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_emo_numero = " + CStr(rsaux5!INTE_SAL_NUMERO), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux6.EOF Then
                                 var_numero_orden_surtido = IIf(IsNull(rsaux6!inte_emo_numero_origen), 0, rsaux6!inte_emo_numero_origen)
                                 If rsaux9.State = 1 Then
                                    rsaux9.Close
                                 End If
                                 rsaux9.Open "SELECT * FROM VW_PEDIDOS_CREDITO_TIENDAS WHERE INTE_ORS_ORDEN_SURTIDO = " + CStr(var_numero_orden_surtido), cnn, adOpenDynamic, adLockOptimistic
                                 var_pedido_credito = 0
                                 If Not rsaux9.EOF Then
                                    var_pedido_credito = IIf(IsNull(rsaux9!inte_ped_pedido_credito), 0, rsaux9!inte_ped_pedido_credito)
                                 End If
                                 If var_pedido_credito = 0 Then
                                    rsaux7.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id = 'FT' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_emo_numero = " + CStr(rsaux6!inte_emo_numero), cnn, adOpenDynamic, adLockOptimistic
                                    var_importe_pedido_tienda = IIf(IsNull(rsaux7!floa_Car_importe_neto), 0, rsaux7!floa_Car_importe_neto) / IIf(IsNull(rsaux7!floa_car_tipo_cambio), 1, rsaux7!floa_car_tipo_cambio)
                                    var_numero_factura_tienda = IIf(IsNull(rsaux7!inte_Car_numero), 0, rsaux7!inte_Car_numero)
                                    var_tipo_Cambio_tienda = IIf(IsNull(rsaux7!floa_car_tipo_cambio), 1, rsaux7!floa_car_tipo_cambio)
                                    var_importe_descuento_1_tienda = IIf(IsNull(rsaux7!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rsaux7!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
                                    var_importe_descuento_2_tienda = IIf(IsNull(rsaux7!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rsaux7!FLOA_CAR_PORCENTAJE_DESCUENTO_2)
                                    rsaux7.Close
                                    rsaux7.Open "select * from vw_clientes where vcha_cli_clave_id = '" + rsaux6!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux7.EOF Then
                                       var_referencia_cliente_tienda = IIf(IsNull(rsaux7!VCHA_CLI_REFERENCIA), "", rsaux7!VCHA_CLI_REFERENCIA)
                                       var_clave_cliente_tienda = IIf(IsNull(rsaux7!vcha_cli_clave_id), "", rsaux7!vcha_cli_clave_id)
                                       var_agente_cliente_tienda = IIf(IsNull(rsaux7!VCHA_AGE_AGENTE_ID), "", rsaux7!VCHA_AGE_AGENTE_ID)
                                       var_canal_cliente_tienda = IIf(IsNull(rsaux7!vcha_can_canal_venta_id), "", rsaux7!vcha_can_canal_venta_id)
                                       var_grupo_real_tienda = IIf(IsNull(rsaux7!vcha_gre_grupo_real_id), "", rsaux7!vcha_gre_grupo_real_id)
                                       var_grupo_actual_tienda = IIf(IsNull(rsaux7!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux7!VCHA_GAC_GRUPO_aCTUAL_ID)
                                       var_titular_tienda = IIf(IsNull(rsaux7!vcha_tit_titular_id), "", rsaux7!vcha_tit_titular_id)
                                       var_porcentaje_iva_tienda = IIf(IsNull(rsaux7!FLOA_TPE_IVA), "", rsaux7!FLOA_TPE_IVA)
                                       var_clave_moneda_tienda = IIf(IsNull(rsaux7!vcha_mon_moneda_id), "1", rsaux7!vcha_mon_moneda_id)
                                    End If
                                    rsaux7.Close
                                    If rsaux8.State = 1 Then
                                       rsaux8.Close
                                    End If
                                    rsaux8.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id = 'FT' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_emo_numero = " + CStr(rsaux6!inte_emo_numero) + " and (char_car_estatus <> 'C' or char_car_estatus is null)", cnn, adOpenDynamic, adLockOptimistic
                                    var_importe_facturado_ft = 0
                                    While Not rsaux8.EOF
                                             'rs.Open "select max(inte_car_numero) as maximo_numero from tb_encabezado_cartera where vcha_car_tipo_documento = 'PA'", cnn, adOpenDynamic, adLockOptimistic
                                             'If rs.EOF Then
                                             '   var_numero_folio = 0
                                             'Else
                                             '   var_numero_folio = IIf(IsNull(rs!maximo_numero), 0, rs!maximo_numero)
                                             'End If
                                             'rs.Close
                                             
                                             
                                             rsaux11.Open "select * from TB_MAXIMO_PAGO", cnn_sid_quezada, adOpenDynamic, adLockOptimistic
                                             If rsaux11.EOF Then
                                                var_numero_folio = 0
                                             Else
                                                var_numero_folio = IIf(IsNull(rsaux11!inte_max_maximo_pago), 0, rsaux11!inte_max_maximo_pago)
                                             End If
                                             rsaux11.Close
                                             var_numero_folio = var_numero_folio + 1
                                             rsaux11.Open "update TB_MAXIMO_PAGO set inte_max_maximo_pago = inte_max_maximo_pago + 1", cnn_sid_quezada, adOpenDynamic, adLockOptimistic
                                               
                                             Cadena = "INSERT INTO TB_ENCABEZADO_CARTERA (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, "
                                             Cadena = Cadena + "FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3, FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, VCHA_SER_SERIE_ID, CHAR_CAR_ESTATUS, VCHA_CAR_CHEQUE_DEPOSITO, VCHA_CAR_BANCO_CHEQUE, VCHA_CAR_DEPOSITO, VCHA_CAR_BANCO_DEPOSITO, DTIM_CAR_FECHA_DEPOSITO) values ("
                                             Cadena = Cadena + "'" + var_empresa + "', '" + var_unidad_organizacional + "', 'PA', 'PA', 'PA', " + CStr(var_numero_folio) + ", '-', '', '', 0, getdate(), '" + var_agente_cliente_tienda + "', '" + var_grupo_actual_tienda + "', '" + var_grupo_real_tienda + "', '" + var_titular_tienda + "', '" + var_clave_cliente_tienda + "', '', 0, " + CStr(var_porcentaje_iva_tienda) + ", 0, 0, " + CStr(var_importe_descuento_1_tienda) + ", " + CStr(var_importe_descuento_2_tienda) + ", 0, " + CStr(rsaux8!floa_Car_importe_neto) + ", " + CStr(rsaux8!floa_Car_importe_neto - rsaux8!floa_car_subimporte) + ", 0, 0, 0, 0, 0, " + CStr(rsaux8!floa_car_subimporte) + ", " + CStr(rsaux8!floa_Car_importe_neto) + ", '', '"
                                             Cadena = Cadena + CStr(var_clave_usuario_global) + "', '', getdate(), 0, getdate(), getdate(), '" + var_clave_moneda + "', " + CStr(var_tipo_Cambio_tienda) + ", 'FT', 'I','', '', '','','')"
                                             rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                             rsaux7.Open "update tb_encabezado_cartera set inte_car_pedido_credito = 0 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Car_documento = 'FA' and vcha_Ser_serie_id = '" + var_serie + "' and inte_Car_numero = " + CStr(CDbl(rsaux8!inte_Car_numero)), cnn, adOpenDynamic, adLockOptimistic
                                             Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
                                             var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, var_serie, "FA", CDbl(rsaux8!inte_Car_numero), "FT", "PA", CDbl(var_numero_folio), 0, CDbl(rsaux8!floa_Car_importe_neto))
                                             var_importe_facturado_ft = var_importe_facturado_ft + rsaux8!floa_Car_importe_neto
                                             If CDbl(Round(rsaux8!floa_Car_importe_neto, 2)) > 0 Then
                                                rsaux7.Open "select NUMB_SAL_IMPORTE   from tb_saldo where vcha_sal_Referencia = '" + Trim(var_referencia_cliente_tienda) + "'", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                                                var_importe_Saldo_oracle = IIf(IsNull(rsaux7(0).Value), 0, rsaux7(0).Value)
                                                var_importe_factura = Round(CDbl(rsaux8!floa_Car_importe_neto), 2)
                                                var_diferencia_saldo_factura = var_importe_Saldo_oracle - var_importe_factura
                                                If var_diferencia_saldo_factura < 0 And var_diferencia_saldo_factura > -10 Then
                                                   var_importe_factura = var_importe_factura + (var_diferencia_saldo_factura)
                                                End If
                                                rsaux7.Close
                                                'rsaux7.Open "call SP_AGREGA_CARGO ('" + var_canal_cliente_tienda + "','" + var_agente_cliente_tienda + "', " + CStr(CDbl(rsaux8!INTE_CAR_NUMERO)) + ",'" + Trim(var_referencia_cliente_tienda) + "'," + CStr(CDbl(Round(CDbl(rsaux8!floa_Car_importe_neto), 2))) + ", 0,TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'VA')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                                                rsaux7.Open "call SP_AGREGA_CARGO ('" + var_canal_cliente_tienda + "','" + var_agente_cliente_tienda + "', " + CStr(CDbl(rsaux8!inte_Car_numero)) + ",'" + Trim(var_referencia_cliente_tienda) + "'," + CStr(CDbl(Round(CDbl(var_importe_factura), 2))) + ", 0,TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'VA')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                                             End If
                                             
                                             rsaux8.MoveNext
                                    Wend
                                    If rsaux8.State = 1 Then
                                       rsaux8.Close
                                    End If
                                    
                                    
                                    
                                    rsaux4.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + CStr(rsaux6!inte_emo_numero_origen), cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux4.EOF Then
                                       var_pedido_tienda = rsaux4!inte_ped_numero
                                       rsaux4.Close
                                       rsaux4.Open "select * from VW_IMPORTES_SEGURO_PAQUETERIA where inte_ped_numero = " + CStr(var_pedido_tienda), cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux4.EOF Then
                                          var_importe_paqueteria_tienda = IIf(IsNull(rsaux4!importe_seguro), 0, rsaux4!importe_seguro)
                                          var_importe_seguro_tienda = IIf(IsNull(rsaux4!importe_paqueteria), 0, rsaux4!importe_paqueteria)
                                          var_importe_referencia_tienda = IIf(IsNull(rsaux4!floa_paq_costo_referencia), 0, rsaux4!floa_paq_costo_referencia)
                                          var_importe_pedido_tienda = IIf(IsNull(rsaux4!importe_pedido), 0, rsaux4!importe_pedido)
                                          var_importe_total_tienda = var_importe_pedido_tienda + var_importe_paqueteria_tienda + var_importe_seguro_tienda + var_importe_referencia_tienda
                                          rsaux4.MoveNext
                                       End If
                                       rsaux4.Close
                                    Else
                                       rsuax4.Close
                                    End If
                                    If CDbl(Round(var_importe_total_tienda, 2)) > CDbl(Round(var_importe_facturado_ft, 2)) Then
                                       var_diferencia_facturado = var_importe_total_tienda - var_importe_facturado_ft
                                       rsaux8.Open "CALL SP_AGREGA_ABONO('" + Trim(var_referencia_cliente_tienda) + "',0.00," + CStr(var_diferencia_facturado) + ",SYSDATE,SYSDATE,'" + CStr(var_pedido_tienda) + "','','DF','')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                                    End If
                                 End If
                              Else
                              End If
                              rsaux6.Close
                              cnn.CommitTrans
                           End If
                           rsaux5.MoveNext
                     Wend
                     End If
                     rsaux5.Close
                     Call envio_tb_transito
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
                     cnn.CommandTimeout = 360
                     'Cadena = "INSERT INTO [vianney].[dbo].[TB_TEMP_FACTURA_EMBARQUES] ([INTE_TEM_CONSECUTIVO], [VCHA_AGR_AGRUPADOR_ID], [VCHA_SAL_DESCRIPCION_FACTURA], [IMPORTE], [CANTIDAD], [INTE_CAR_PLAZO], [FLOA_CAR_PORCENTAJE_IVA], [FLOA_CAR_PORCENTAJE_IMPUESTO_1], [FLOA_CAR_PORCENTAJE_IMPUESTO_2], [FLOA_CAR_PORCENTAJE_DESCUENTO_1], [FLOA_CAR_PORCENTAJE_DESCUENTO_2], [FLOA_CAR_PORCENTAJE_DESCUENTO_3], [FLOA_CAR_IMPORTE_TOTAL], [FLOA_CAR_IMPORTE_IVA], [FLOA_CAR_IMPORTE_IMPUESTO_1], [FLOA_CAR_IMPORTE_IMPUESTO_2], [FLOA_CAR_IMPORTE_DESCUENTO_1], [FLOA_CAR_IMPORTE_DESCUENTO_2],"
                     'Cadena = Cadena + " [FLOA_CAR_IMPORTE_DESCUENTO_3], [FLOA_CAR_SUBIMPORTE], [FLOA_CAR_IMPORTE_NETO], [VCHA_CAR_IMPORTE_LETRA], [VCHA_SER_SERIE_ID], [VCHA_CAR_DOCUMENTO], [INTE_CAR_NUMERO], [DTIM_CAR_FECHA], [VCHA_EMP_EMPRESA_ID], [INTE_EMB_EMBARQUE], [VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_CLI_CLAVE_ID], [VCHA_CLI_NOMBRE], [VCHA_CLI_REPRESENTANTE], [VCHA_AGE_AGENTE_ID], [VCHA_RUT_RUTA_ID], [VCHA_CLI_CURP], [VCHA_CLI_RFC], [VCHA_MON_MONEDA_ID], [VCHA_PLA_PLAZO_ID], [VCHA_TCL_TIPO_CLIENTE_ID], [VCHA_LIS_LISTA_ID], [VCHA_CAN_CANAL_VENTA_ID], [VCHA_TRA_TRANSPORTE_ID], [VCHA_FAG_FAMILIA_AGRUPADOR_ID], [INTE_CLI_AGRUPADOR],"
                     'Cadena = Cadena + " [INTE_CLI_ESTATUS], [VCHA_TIT_TITULAR_ID], [CHAR_PRI_PRIORIDAD_ID], [VCHA_CLI_EMAIL], [VCHA_PAI_PAIS_ID], [VCHA_PAI_NOMBRE], [VCHA_EST_ESTADO_ID], [VCHA_EST_NOMBRE], [VCHA_CIU_CIUDAD_ID], [VCHA_CIU_NOMBRE], [VCHA_CLI_COLONIA], [VCHA_CLI_DIRECCION], [VCHA_CLI_CP], [FLOA_GRE_DESCUENTO_1], [FLOA_GRE_DESCUENTO_2], [FLOA_GRE_DESCUENTO_3],  [VCHA_GRE_GRUPO_REAL_ID], [VCHA_GRE_NOMBRE], [VCHA_GAC_GRUPO_ACTUAL_ID], [VCHA_TIT_NOMBRE], [FLOA_TIT_LIMITE_CREDITO], [INTE_PLA_DIAS], [FLOA_GAC_DESCUENTO_1], [FLOA_GAC_DESCUENTO_2], [FLOA_GAC_DESCUENTO_3], [VCHA_CAN_NOMBRE], [INTE_CAN_BUSQUEDA_FACTURA_GRUPO], [FLOA_TPE_IVA], [VCHA_GAC_NOMBRE], "
                     'Cadena = Cadena + " [VCHA_MON_NOMBRE], [VCHA_MON_NOMBRE_PLURAL], [VCHA_AGE_NOMBRE], [VCHA_ESB_ESTABLECIMIENTO_ID], [FLOA_CAR_TIPO_CAMBIO], [INTE_ORS_ORDEN_SURTIDO], [INTE_PED_NUMERO], [FLOA_SAL_PROMOCION_1],  [FLOA_SAL_PROMOCION_2], [VCHA_CAR_TIPO_DOCUMENTO], [VCHA_UOR_UNIDAD_ID], [INTE_JAU_JAULA_ID], [VCHA_VEH_VEHICULO_ID], [DTIM_EMB_FECHA_INICIO], [DTIM_EMB_FECHA_FINAL], [CHAR_EMB_ESTATUS], [VCHA_CHO_CHOFER_ID], [FLOA_EMB_CUBICAJE], [CHAR_CAR_TIPO_FACTURACION], [VCHA_CAR_CLASE_ID], [CHAR_CAR_AFECTACION], [VCHA_ALM_ALMACEN_ID], [Expr1], [INTE_EMO_NUMERO], [Expr2], [Expr3], [Expr4], [Expr5], [Expr6], [Expr7], [VCHA_AUD_USUARIO], [VCHA_AUD_MAQUINA], [VCHA_AUD_FECHA], [FLOA_CAR_SALDO], "
                     'Cadena = Cadena + " [DTIM_CAR_FECHA_VENCIMIENTO], [DTIM_CAR_FECHA_ENTREGA],[Expr8], [CHAR_CAR_ESTATUS], [DTIM_CAR_FECHA_CANCELACION], [VCHA_CAR_USUARIO_CANCELACION],  [VCHA_CAR_MAQUINA_CANCELACION], [INTE_CLI_ENVIO_FACTURA], [FLOA_SAL_PRECIO_PROMEDIO],  [INTE_CAR_FACTURA_CEROS], [FLOA_CAR_COSTO], [INTE_SAL_CONSECUTIVO_FACTURA])"
                     'Cadena = Cadena + " select " + CStr(var_consecutivo) + ", VCHA_AGR_AGRUPADOR_ID, VCHA_SAL_DESCRIPCION_FACTURA, IMPORTE, CANTIDAD, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3, FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_SER_SERIE_ID, VCHA_CAR_DOCUMENTO, INTE_CAR_NUMERO, DTIM_CAR_FECHA, VCHA_EMP_EMPRESA_ID, INTE_EMB_EMBARQUE, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, "
                     'Cadena = Cadena + " VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_CLI_REPRESENTANTE, VCHA_AGE_AGENTE_ID, VCHA_RUT_RUTA_ID, VCHA_CLI_CURP, VCHA_CLI_RFC, VCHA_MON_MONEDA_ID, VCHA_PLA_PLAZO_ID, VCHA_TCL_TIPO_CLIENTE_ID, VCHA_LIS_LISTA_ID, VCHA_CAN_CANAL_VENTA_ID, VCHA_TRA_TRANSPORTE_ID, VCHA_FAG_FAMILIA_AGRUPADOR_ID, INTE_CLI_AGRUPADOR, INTE_CLI_ESTATUS, VCHA_TIT_TITULAR_ID, CHAR_PRI_PRIORIDAD_ID, VCHA_CLI_EMAIL, VCHA_PAI_PAIS_ID, VCHA_PAI_NOMBRE, VCHA_EST_ESTADO_ID, VCHA_EST_NOMBRE, VCHA_CIU_CIUDAD_ID, VCHA_CIU_NOMBRE, VCHA_CLI_COLONIA, VCHA_CLI_DIRECCION, VCHA_CLI_CP, FLOA_GRE_DESCUENTO_1, FLOA_GRE_DESCUENTO_2, FLOA_GRE_DESCUENTO_3, VCHA_GRE_GRUPO_REAL_ID, VCHA_GRE_NOMBRE, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_TIT_NOMBRE, FLOA_TIT_LIMITE_CREDITO, INTE_PLA_DIAS, "
                     'Cadena = Cadena + " FLOA_GAC_DESCUENTO_1, FLOA_GAC_DESCUENTO_2, FLOA_GAC_DESCUENTO_3, VCHA_CAN_NOMBRE, INTE_CAN_BUSQUEDA_FACTURA_GRUPO, FLOA_TPE_IVA, VCHA_GAC_NOMBRE, VCHA_MON_NOMBRE, VCHA_MON_NOMBRE_PLURAL, VCHA_AGE_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, FLOA_CAR_TIPO_CAMBIO, INTE_ORS_ORDEN_SURTIDO, INTE_PED_NUMERO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, VCHA_CAR_TIPO_DOCUMENTO, VCHA_UOR_UNIDAD_ID, INTE_JAU_JAULA_ID, VCHA_VEH_VEHICULO_ID, DTIM_EMB_FECHA_INICIO, DTIM_EMB_FECHA_FINAL, CHAR_EMB_ESTATUS, VCHA_CHO_CHOFER_ID, FLOA_EMB_CUBICAJE, CHAR_CAR_TIPO_FACTURACION, VCHA_CAR_CLASE_ID, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, Expr1, INTE_EMO_NUMERO, Expr2, Expr3, Expr4, Expr5, Expr6, Expr7, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, "
                     'Cadena = Cadena + " FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, Expr8, CHAR_CAR_ESTATUS, DTIM_CAR_FECHA_CANCELACION, VCHA_CAR_USUARIO_CANCELACION, VCHA_CAR_MAQUINA_CANCELACION, INTE_CLI_ENVIO_FACTURA, FLOA_SAL_PRECIO_PROMEDIO, INTE_CAR_FACTURA_CEROS, FLOA_CAR_COSTO, INTE_SAL_CONSECUTIVO_FACTURA from vw_facturas_embarque where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque
                     
                     If var_empresa = "18" Then
                        Cadena = "EXEC SP_CREA_TABLA_FACTURAS " + CStr(var_consecutivo) + ",'" + var_empresa + "'," + Me.txt_numero_embarque
                     Else
                        Cadena = "EXEC SP_CREA_TABLA_FACTURAS_nuevo " + CStr(var_consecutivo) + ",'" + var_empresa + "'," + txt_numero_embarque
                     End If
                     rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     Me.frm_mensaje.Visible = False
''''' inicio turbina

                     If var_empresa = "30" Then
                        Call factura_turbina
                     End If

                     If var_empresa = "15" Then
                        Call factura_estampados
                     End If


''''' fin turbina
                     If var_empresa = "17" Then
                        Call factura_bordalesa
                     End If
                     
                     
                     
                     If var_empresa = "02" Or var_empresa = "18" Or var_empresa = "06" Or var_empresa = "31" Then
                        rsaux3.Open "select distinct inte_car_numero from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'  ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           'If (var_empresa = "06" And Trim(UCase(parametros(0))) = "DISTRIBUCION") Or (var_empresa = "06" And Trim(UCase(parametros(0))) = "sqlquezada2") Then
                           '   var_Archivo = App.Path & "\factura" + Trim(Str(rsaux3!inte_car_numero)) + ".bat"
                           '   Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_car_numero)) + ".bat") For Output As #2
                           'Else
                           '   If var_empresa = "06" Then
                           '      var_Archivo = "c:\fact" + Trim(Str(rsaux3!inte_car_numero)) + ".bat"
                           '      Open ("c:\fact" + Trim(Str(rsaux3!inte_car_numero)) + ".bat") For Output As #2
                           '   End If
                           'End If
                           If var_empresa = "02" Or var_empresa = "18" Or var_empresa = "06" Or var_empresa = "31" Or var_empresa = "17" Then
                              var_Archivo = App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat"
                              Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat") For Output As #2
                           End If
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
                                 'If var_empresa = "06" And Trim(UCase(parametros(0))) <> "DISTRIBUCION" Then
                                 '   If var_empresa = "06" And Trim(UCase(parametros(0))) = "sqlquezada2" Then
                                 '      Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_car_numero)) + ".txt") For Output As #1
                                 '   Else
                                 '      Open ("c:\fact" + Trim(Str(rsaux3!inte_car_numero)) + ".txt") For Output As #1
                                 '   End If
                                 'Else
                                    Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt") For Output As #1
                                 'End If
                                  
                                  'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                                 'Print #1, Chr(27) + Chr(67) + Chr(44) + Chr(13)
                                 'Print #1, ""
                                 Print #1, Chr(15) + Chr(27) + Chr(64)
                                 If var_empresa = "18" Or var_empresa = "06" Or var_empresa = "17" Then
                                    If var_unidad_organizacional = "29" Then
                                    Else
                                       'Print #1, ""
                                    End If
                                 End If
                                 Print #1, Spc(105); Str(rsaux3!inte_Car_numero)
                                 Print #1, ""
                                 Print #1, ""
                                    Print #1, Spc(105); Str(rs!INTE_CAR_PLAZO) + " DIAS DE VENCIMIENTO" + "                  " + Format(rs!dtim_Car_fecha, "Short Date")
                                 Print #1, ""
                                 'Print #1, Spc(92); Str(rs!inte_car_PLAZO) + " DIAS DE VENCIMIENTO"
                                 var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                                 var_cliente_coppel = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                                 var_cliente_sigo = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                                 
                                 For var_j = 1 + Len(Trim(var_cliente)) To 83
                                     var_cliente = var_cliente + " "
                                 Next var_j
                                 If var_unidad_organizacional = "21" Then
                                     var_cliente = var_cliente + "               MEXICO, D.F."
                                 Else
                                     var_cliente = var_cliente + "               AGUASCALIENTES, AGS."
                                 End If
                                 If var_unidad_organizacional = "39" Then
                                    Print #1, ""
                                 End If
                                 Print #1, Spc(10); var_cliente
                                 ''' CAMBIO PARA AGREGAR COLONIA
                                 'var_domicilio = IIf(IsNull(rs!vcha_cli_direccion), "", rs!vcha_cli_direccion) + " C.P. " + IIf(IsNull(rs!vcha_cli_cp), "", rs!vcha_cli_cp)
                                 If var_cliente_coppel = "C000006202" Then
                                    var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION)
                                 Else
                                    var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " COLONIA: " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
                                 End If
                                 'aqui trono 30/10/2009
                                 rsaux11.Open "select vcha_cli_referencia from tb_Clientes where vcha_Cli_clave_id = '" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                 If Not rsaux11.EOF Then
                                    var_referencia_Bancaria = Trim(IIf(IsNull(rsaux11!VCHA_CLI_REFERENCIA), "", rsaux11!VCHA_CLI_REFERENCIA))
                                    If var_referencia_Bancaria <> "" Then
                                       For var_j = 1 + Len(Trim(var_domicilio)) To 105
                                           var_domicilio = var_domicilio + " "
                                       Next var_j
                                       var_domicilio = var_domicilio + " REF. BANCARIA: " + var_referencia_Bancaria
                                    End If
                                 End If
                                 rsaux11.Close
                                 
                                 'For var_j = 1 + Len(Trim(var_domicilio)) To 83
                                 '    var_domicilio = var_domicilio + " "
                                 'Next var_j
                                 ''' FIN DE CAMBIO PARAA AGREGAR COLONIA
                                 
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
                                 If var_empresa = "06" Or var_empresa = "17" Then
                                    var_ciudad = var_ciudad + "                                            " + var_agente
                                 Else
                                    var_ciudad = var_ciudad + "                                                      " + var_agente
                                 End If
                                 VAR_EMBARQUE = "EMB.: " + txt_numero_embarque
                                 var_ordern_surtido = x
                                 Print #1, Spc(10); var_ciudad
                                 var_rfc = "RFC:  " + var_rfc
                                 var_establecimiento_comercial = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
                                 var_rfc = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                                 For var_j = 1 + Len(Trim(var_rfc)) To 89
                                    var_rfc = var_rfc + " "
                                 Next var_j
                                 If Trim(var_cliente_coppel) = "C000002947" Or Trim(var_cliente_coppel) = "C000005566" Or Trim(var_cliente_coppel) = "C000005831" Or var_empresa = "06" Or var_empresa = "17" Then
                                    rsaux5.Open "select * from tb_encabezado_pedidos where inte_ped_numero = " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))), cnn, adOpenDynamic, adLockOptimistic
                                    var_solicitud_sigo = Trim(CStr(IIf(IsNull(rsaux5!VCHA_PED_PEDIDO_EXTERNO), "", rsaux5!VCHA_PED_PEDIDO_EXTERNO)))
                                    If Not rsaux5.EOF Then
                                       If Trim(var_cliente_coppel) = "C000005831" Then
                                          var_rfc = var_rfc + "               O.C.: " + Trim(CStr(IIf(IsNull(rsaux5!VCHA_PED_PEDIDO_EXTERNO), "", rsaux5!VCHA_PED_PEDIDO_EXTERNO))) + " "
                                       Else
                                          If var_empresa = "06" Or var_empresa = "17" Then
                                             var_rfc = var_rfc + "               REF.: " + Trim(CStr(IIf(IsNull(rsaux5!VCHA_PED_PEDIDO_EXTERNO), "", rsaux5!VCHA_PED_PEDIDO_EXTERNO))) + " "
                                          Else
                                             var_rfc = var_rfc + "               PED.: " + Trim(CStr(IIf(IsNull(rsaux5!VCHA_PED_PEDIDO_EXTERNO), "", rsaux5!VCHA_PED_PEDIDO_EXTERNO))) + " "
                                          End If
                                       End If
                                    End If
                                    rsaux5.Close
                                 Else
                                    var_rfc = var_rfc + "               PED.: " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))) + " "
                                 End If
                                 var_rfc = var_rfc + " O.S.: " + Trim(Str(IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), 0, rs!INTE_ORS_ORDEN_SURTIDO))) + " " + VAR_EMBARQUE
                                 Print #1, var_rfc
                                 'Print #1, Spc(10); IIf(IsNull(rs!vcha_esb_establecimiento_id), "", rs!vcha_esb_establecimiento_id)
                                 Print #1, ""
                                 If var_empresa = "06" Or var_empresa = "17" Then
                                    If var_unidad_organizacional = "29" Then
                                       Print #1, ""
                                    End If
                                 Else
                                    Print #1, ""
                                 End If
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
                                       If var_empresa = "31" Then
                                          For var_j = 1 + Len(Trim(var_linea)) To 17
                                              var_linea = var_linea + " "
                                          Next var_j
                                       Else
                                          For var_j = 1 + Len(Trim(var_linea)) To 15
                                              var_linea = var_linea + " "
                                          Next var_j
                                       End If
                                       If var_empresa = "15" Then
                                          var_linea = var_linea + "MAQUILA DE " + UCase(IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura))
                                       Else
                                          var_linea = var_linea + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                       End If
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
                                 
                                 If var_empresa = "18" Or var_empresa = "31" Then
                                    Print #1, ""
                                 End If
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
                                    If Trim(var_cliente_coppel) = "C000010568" Or Trim(var_cliente_coppel) = "C000008200" Then
                                       var_importe_descuento_2_str = "" + Format(rs!floa_Car_importe_neto, "###,###,##0.00")
                                    Else
                                       var_importe_descuento_2_str = Format(IIf(IsNull(rs!floa_car_importe_descuento_2), 0, rs!floa_car_importe_descuento_2) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio), "###,###,##0.00")
                                    End If
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
                                 If var_empresa = "02" Then
                                    var_descuento_leyenda = 0
                                    var_descuento_leyenda = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
                                    If Trim(var_cliente_coppel) = "C000005566" Then
                                       rsaux11.Open "select * from vw_establecimientos_direcciones where vcha_esb_establecimiento_id = '" + var_establecimiento_comercial + "'", cnn, adOpenDynamic, adLockOptimistic
                                       var_linea = "DIRECCION DE ENTREGA: " + IIf(IsNull(rsaux11!vcha_esb_domicilio), "", rsaux11!vcha_esb_domicilio) + " COLONIA: " + IIf(IsNull(rsaux11!vcha_col_nombre), "", rsaux11!vcha_col_nombre)
                                       Print #1, var_linea
                                       var_linea = IIf(IsNull(rsaux11!vcha_ciu_nombre), "", rsaux11!vcha_ciu_nombre) + ", " + IIf(IsNull(rsaux11!vcha_est_nombre), "", rsaux11!vcha_est_nombre) + ", " + IIf(IsNull(rsaux11!vcha_pai_nombre), "", rsaux11!vcha_pai_nombre) + " C.P. " + IIf(IsNull(rsaux11!vcha_esb_cp), "", rsaux11!vcha_esb_cp) + " Tel: " + IIf(IsNull(rsaux11!vcha_esb_telefono), "", rsaux11!vcha_esb_telefono)
                                       Print #1, var_linea
                                       var_linea = ""
                                       rsaux11.Close
                                    Else
                                       If var_descuento_leyenda >= 13 Then
                                          If Trim(var_cliente_coppel) = "C000001636" Then
                                             var_linea = ""
                                          Else
                                             var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                                          End If
                                          If Len(Trim(var_linea)) < 145 Then
                                             For var_j = 1 + Len(Trim(var_linea)) To 145
                                                 var_linea = var_linea + " "
                                             Next var_j
                                          End If
                                          Print #1, var_linea + var_importe_descuento_1_str
                                          If var_empresa = "18" Then
                                             var_linea = ""
                                          Else
                                             If Trim(var_cliente_coppel) = "C000002947" Or Trim(var_cliente_coppel) = "C000001636" Then
                                                If Trim(var_cliente_coppel) = "C000002947" Then
                                                   rsaux11.Open "select * from TB_eSTABLECIMIENTOS where vcha_esb_establecimiento_id = '" + var_establecimiento_comercial + "'", cnn, adOpenDynamic, adLockOptimistic
                                                   var_linea = "ESTABLECIMIENTO: " + IIf(IsNull(rsaux11!VCHA_ESB_NOMBRE), "", rsaux11!VCHA_ESB_NOMBRE)
                                                   rsaux11.Close
                                                Else
                                                   var_linea = ""
                                                End If
                                             Else
                                                var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%" + " " + var_leyenda_sorteo
                                             End If
                                          End If
                                          If Len(Trim(var_linea)) < 145 Then
                                             For var_j = 1 + Len(Trim(var_linea)) To 145
                                                 var_linea = var_linea + " "
                                             Next var_j
                                          End If
                                      Else
                                          If Trim(var_cliente_coppel) = "C000001636" Then
                                             var_linea = ""
                                          Else
                                             var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                                          End If
                                          If Len(Trim(var_linea)) < 145 Then
                                             For var_j = 1 + Len(Trim(var_linea)) To 145
                                                 var_linea = var_linea + " "
                                             Next var_j
                                          End If
                                          Print #1, var_linea + var_importe_descuento_1_str
                                          If var_empresa = "18" Then
                                             var_linea = ""
                                          Else
                                             If Trim(var_cliente_coppel) = "C000002947" Or Trim(var_cliente_coppel) = "C000001636" Then
                                                If Trim(var_cliente_coppel) = "C000002947" Then
                                                   rsaux11.Open "select * from TB_eSTABLECIMIENTOS where vcha_esb_establecimiento_id = '" + var_establecimiento_comercial + "'", cnn, adOpenDynamic, adLockOptimistic
                                                   var_linea = "ESTABLECIMIENTO: " + IIf(IsNull(rsaux11!VCHA_ESB_NOMBRE), "", rsaux11!VCHA_ESB_NOMBRE)
                                                   rsaux11.Close
                                                Else
                                                   var_linea = ""
                                                End If
                                             Else
                                                var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%" + " " + var_leyenda_sorteo
                                             End If
                                          End If
                                          If Len(Trim(var_linea)) < 145 Then
                                             For var_j = 1 + Len(Trim(var_linea)) To 145
                                                 var_linea = var_linea + " "
                                             Next var_j
                                          End If
                                       End If
                                    End If ' comercial
                                 Else
                                    If Trim(var_cliente_coppel) = "C000001636" Then
                                       var_linea = ""
                                    Else
                                       '' aqui debe de ir lo del desperdicio
                                       If Trim(var_cliente_coppel) = "C000010568" Or Trim(var_cliente_coppel) = "C000008200" Then
                                          var_linea = "Impuesto retenido de conformidad con la Ley del Impuesto al Valor Agregado"
                                       Else
                                          var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                                       End If
                                    End If
                                    If Len(Trim(var_linea)) < 145 Then
                                       For var_j = 1 + Len(Trim(var_linea)) To 145
                                           var_linea = var_linea + " "
                                       Next var_j
                                    End If
                                    If Trim(var_cliente_coppel) = "C000010568" Or Trim(var_cliente_coppel) = "C000008200" Then
                                       Print #1, var_linea
                                    Else
                                       Print #1, var_linea + var_importe_descuento_1_str
                                    End If
                                    If var_empresa = "18" Then
                                       var_linea = ""
                                    Else
                                       If Trim(var_cliente_coppel) = "C000002947" Or Trim(var_cliente_coppel) = "C000001636" Then
                                          If Trim(var_cliente_coppel) = "C000002947" Then
                                             rsaux11.Open "select * from TB_eSTABLECIMIENTOS where vcha_esb_establecimiento_id = '" + var_establecimiento_comercial + "'", cnn, adOpenDynamic, adLockOptimistic
                                             var_linea = "ESTABLECIMIENTO: " + IIf(IsNull(rsaux11!VCHA_ESB_NOMBRE), "", rsaux11!VCHA_ESB_NOMBRE)
                                             rsaux11.Close
                                          Else
                                             var_linea = ""
                                          End If
                                       Else
                                          If Trim(var_cliente_coppel) <> "C000010568" Then
                                             If Trim(var_cliente_coppel) = "C000008200" Then
                                                var_linea = "INCISO B DE LA LEY DEL IMPUESTO AL VALOR AGREGADO"
                                             Else
                                                var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%" + " " + var_leyenda_sorteo
                                             End If
                                          Else
                                             var_linea = "INCISO B DE LA LEY DEL IMPUESTO AL VALOR AGREGADO"
                                          End If
                                       End If
                                    End If
                                    If Len(Trim(var_linea)) < 145 Then
                                       For var_j = 1 + Len(Trim(var_linea)) To 145
                                           var_linea = var_linea + " "
                                       Next var_j
                                    End If
                                 End If
                                 'aqui no se pone nada de la comercial'
                                 If var_empresa = "02" Then
                                    If Trim(var_cliente_coppel) <> "C000002947" Then
                                       If Trim(var_cliente_coppel) = "C000010568" Or Trim(var_cliente_coppel) = "C000008200" Then
                                       Else
                                          var_linea = var_linea + var_importe_descuento_2_str
                                       End If
                                    End If
                                    If Trim(var_cliente_coppel) = "C000005566" Then
                                    Else
                                       Print #1, var_linea
                                    End If
                                    ''var_contador_promociones = 1 ' se pone para poder poner la leyenda del IVA del 16%
                                    If var_contador_promociones > 0 Then
                                       If var_cliente_sigo = "C000001636" Then
                                          'Print #1, "Descuento adicional del 2%"
                                          Print #1, var_solicitud_sigo
                                       Else
                                          Print #1, var_cadena_promocion_171209
                                       End If
                                    Else
                                       If var_cliente_sigo = "C000001636" Then
                                          'Print #1, "Descuento adicional del 2%"
                                          Print #1, var_solicitud_sigo
                                       Else
                                          Print #1, ""
                                       End If
                                    End If
                                 Else
                                    If Trim(var_cliente_coppel) = "C000010568" Or Trim(var_cliente_coppel) = "C000008200" Then
                                       var_linea = var_linea + var_importe_descuento_2_str
                                    Else
                                       var_linea = var_linea + var_importe_descuento_2_str
                                    End If
                                    Print #1, var_linea
                                    'var_contador_promociones = 1 ' se pone para poder poner la leyenda del IVA del 16%
                                    If var_contador_promociones > 0 Then
                                       If var_cliente_sigo = "C000001636" Then
                                           'Print #1, "Descuento adicional del 2%"
                                           Print #1, var_solicitud_sigo
                                       Else
                                           Print #1, var_cadena_promocion_171209
                                       End If
                                    Else
                                       If var_cliente_sigo = "C000001636" Then
                                           'Print #1, "Descuento adicional del 2%"
                                           Print #1, var_solicitud_sigo
                                       Else
                                          Print #1, ""
                                       End If
                                    End If
                                 End If
                                 var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                 var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                                 If var_empresa <> "06" Or var_empresa = "17" Then
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
                                 Else
                                    If Len(Trim(var_linea)) < 117 Then
                                       For var_j = 1 + Len(Trim(var_linea)) To 117
                                           var_linea = var_linea + " "
                                       Next var_j
                                    End If
                                 End If
                                 
                                 If Len(Trim(var_rfc)) = 0 Then
                                    If Trim(var_cliente_coppel) = "C000010568" Or Trim(var_cliente_coppel) = "C000008200" Then
                                       var_subimporte = Format(Round((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                    Else
                                       var_subimporte = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                    End If
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
                                    If Trim(var_cliente_coppel) = "C000010568" Or Trim(var_cliente_coppel) = "C000008200" Then
                                       var_subimporte = Format(Round((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                    Else
                                       var_subimporte = Format(Round(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                    End If
                                    If Len(Trim(var_subimporte)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                           var_subimporte = " " + var_subimporte
                                       Next var_j
                                    End If
                                    If Trim(var_cliente_coppel) = "C000010568" Or Trim(var_cliente_coppel) = "C000008200" Then
                                       var_iva = "-" + Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                    Else
                                       var_iva = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                    End If
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
                                 If Trim(var_cliente_coppel) = "C000010568" Or Trim(var_cliente_coppel) = "C000008200" Then
                                    var_linea = "                                                                          ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                        " + var_iva
                                 Else
                                    var_linea = "                                                                          ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                        " + var_iva
                                 End If
                                 Print #1, var_linea
                                 var_dia = Day(rs!dtim_Car_fecha)
                                 var_mes = Month(rs!dtim_Car_fecha)
                                 var_año = Year(rs!dtim_Car_fecha)
                                 If var_empresa = "31" Then
                                    var_linea = "                                                       " + CStr(var_dia) + "     " + CStr(var_mes)
                                 Else
                                    var_linea = "                                                             " + CStr(var_dia) + "     " + CStr(var_mes)
                                 End If
                                 
                                 If Len(var_linea) < 145 Then
                                    For var_j = 1 + Len(var_linea) To 145
                                        var_linea = var_linea + " "
                                    Next var_j
                                 End If
                                 'Print #1, var_linea + var_importe_descuento_1_str
                                 
                                 If var_cliente_coppel = "C000010568" Or Trim(var_cliente_coppel) = "C000008200" Then
                                    var_importe = Format(Round((IIf(IsNull(rs!floa_car_subimporte), 0, rs!floa_car_subimporte)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                 Else
                                    var_importe = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                 End If
                                 
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
                                 If var_empresa = "31" Then
                                    Print #1, Spc(10); Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                                    Print #1, Spc(10); Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre))
                                    Print #1, Spc(10); Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre))
                                 Else
                                    Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                                    Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre))
                                    Print #1, Spc(5); Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre))
                                 End If
                                 If var_empresa <> "03" Then
                                    Print #1, ""
                                    Print #1, ""
                                 Else
                                    Print #1, ""
                                    Print #1, ""
                                 End If
                                 Print #1, ""
                                 If var_empresa = "06" Or var_empresa = "17" Then
                                    If var_unidad_organizacional = "39" Then
                                    Else
                                       Print #1, ""
                                    End If
                                 End If
                                 Print #1, ""
                                 Close #1
                                 'If var_empresa = "06" Then
                                 '   If Trim(UCase(parametros(0))) = "DISTRIBUCION" Or Trim(UCase(parametros(0))) = "sqlquezada2" Then
                                 '      Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_car_numero)) + ".txt lpt1"
                                 '   Else
                                 '      Print #2, "copy c:\fact" + Trim(Str(rsaux3!inte_car_numero)) + ".txt lpt2"
                                 '   End If
                                 'Else
                                    If Trim(var_empresa) = "02" Then
                                       Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt lpt1"
                                    Else
                                       Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt lpt1"
                                    End If
                                 'End If
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
                     '''' AQUI DEBE DE IR EL CORREO DE LAS VENTAS DE TIENDAS
                        If var_si_correo_ft = 1 Then
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
                                    If var_cliente_coppel = "C000006202" Then
                                       Print #1, ""
                                    Else
                                       Print #1, "   Colonia:   " + IIf(IsNull(rsaux8!vcha_col_nombre), "", rsaux8!vcha_col_nombre)
                                    End If
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
                        
                        
                        
                        
                        
                        
                        
                        
                        End If
                     ''''' hasta aqui termina el correo de ventas de tiendas
                     
                     End If
                     
                     If var_empresa = "03" Then
                     var_cliente_import_tp = ""
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
                              var_cliente_sigo = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                              var_cliente_coppel = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                              
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
                              
 ''''' para que imprima la solicitud de sigo
 
                              rsaux5.Open "select * from tb_encabezado_pedidos where inte_ped_numero = " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))), cnn, adOpenDynamic, adLockOptimistic
                              var_solicitud_sigo = Trim(CStr(IIf(IsNull(rsaux5!VCHA_PED_PEDIDO_EXTERNO), "", rsaux5!VCHA_PED_PEDIDO_EXTERNO)))
                              rsaux5.Close
''''
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
                                    If var_empresa = "15" Then
                                       var_linea = var_linea + "MAQUILA DE " + UCase(IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura))
                                    Else
                                       var_linea = var_linea + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                    End If
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
                                 If var_cliente_sigo = "C000005397" Then
                                    'var_importe_descuento_1_str = Format(IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_1), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_1) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio), "###,###,##0.00")
                                    var_importe_descuento_1_str = Format(0, "###,###,##0.00")
                                 Else
                                    var_importe_descuento_1_str = Format(0, "###,###,##0.00")
                                 End If
                                 If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                         var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                    Next var_j
                                 End If
                                 If var_cliente_sigo = "C000005397" Then
                                    'var_importe_descuento_2_str = Format(IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_2), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_2) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio), "###,###,##0.00")
                                    var_importe_descuento_2_str = Format(0, "###,###,##0.00")
                                 Else
                                    var_importe_descuento_2_str = Format(0, "###,###,##0.00")
                                 End If
                                 If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                        var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                    Next var_j
                                 End If
                              Else
                                 var_cantidad_letra = rs!vcha_car_importe_letra
                                 If var_cliente_sigo = "C000005397" Then
                                    'var_importe_descuento_1_str = Format((IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_1), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_1)) * (1 + (rs!floa_car_porcentaje_iva / 100) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                    var_importe_descuento_1_str = Format(0, "###,###,##0.00")
                                 Else
                                    var_importe_descuento_1_str = Format(0, "###,###,##0.00")
                                 End If
                                 If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                         var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                    Next var_j
                                 End If
                                 If var_cliente_sigo = "C000005397" Then
                                    'var_importe_descuento_2_str = Format((IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_2), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_2)) * (1 + (rs!floa_car_porcentaje_iva / 100) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                    var_importe_descuento_1_str = Format(0, "###,###,##0.00")
                                 Else
                                    var_importe_descuento_2_str = Format(0, "###,###,##0.00")
                                 End If
                                 If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                        var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                    Next var_j
                                 End If
                              End If
                              If Trim(var_cliente_coppel) = "C000001636" Or var_cliente_coppel = "C000002912" Or var_cliente_coppel = "C000005397" Then
                                 var_linea = ""
                              Else
                                 var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                              End If
                              If Len(Trim(var_linea)) < 145 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 145
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              If Trim(var_cliente_coppel) = "C000010375" Or var_cliente_coppel = "C000002912" Or var_cliente_coppel = "C000005397" Then  ' del cliente JESSICA KARINA HERANDEZ GARRIDO que no quiere que se impriman las leyendas del descuento
                                 Print #1, ""
                              Else
                                 Print #1, var_linea + var_importe_descuento_1_str
                              End If
                              If var_empresa = "18" Then
                                 var_linea = ""
                              Else
                                 If Trim(var_cliente_coppel) = "C000002947" Or Trim(var_cliente_coppel) = "C000001636" Or Trim(var_cliente_coppel) = "C00000840500" Then
                                    If Trim(var_cliente_coppel) = "C000002947" Then
                                       rsaux11.Open "select * from TB_eSTABLECIMIENTOS where vcha_esb_establecimiento_id = '" + var_establecimiento_comercial + "'", cnn, adOpenDynamic, adLockOptimistic
                                       var_linea = "ESTABLECIMIENTO: " + IIf(IsNull(rsaux11!VCHA_ESB_NOMBRE), "", rsaux11!VCHA_ESB_NOMBRE)
                                       rsaux11.Close
                                    Else
                                       If var_cliente_coppel = "C00000840500" Then
                                          var_linea = "Convenio Aladi 60 dias"
                                       Else
                                          var_linea = ""
                                       End If
                                    End If
                                 Else
                                    var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%" + " " + var_leyenda_sorteo
                                 End If
                              End If
                              If Len(Trim(var_linea)) < 145 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 145
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              var_linea = var_linea + var_importe_descuento_2_str
                              If Trim(var_cliente_coppel) = "C000010375" Or var_cliente_coppel = "C000002912" Or var_cliente_coppel = "C000005397" Then  ' del cliente JESSICA KARINA HERANDEZ GARRIDO que no quiere que se impriman las leyendas del descuento
                                 Print #1, ""
                              Else
                                 Print #1, var_linea
                              End If
                              'var_contador_promociones = 1 ' se pone para poder poner la leyenda del IVA del 16%
                              If var_contador_promociones > 0 Then
                                 If var_cliente_sigo = "C000001636" Then
                                    'Print #1, "Descuento adicional del 2%"
                                    Print #1, var_solicitud_sigo
                                 Else
                                    Print #1, var_cadena_promocion_171209
                                 End If
                              Else
                                 If var_cliente_sigo = "C000001636" Then
                                    'Print #1, "Descuento adicional del 2%"
                                    Print #1, var_solicitud_sigo
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
                     rsaux3.Open "DELETE FROM TB_TEMP_SALIDAS_FACTURACION WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                     MsgBox "Se a terminado el proceso de facturación", vbOKOnly, "ATENCION"
                     If var_empresa = "02" Then
                        If var_unidad_organizacional = "01" Or var_unidad_organizacional = "02" Or var_unidad_organizacional = "03" Or var_unidad_organizacional = "04" Or var_unidad_organizacional = "06" Then
                        Else
                           'var_activa_forma_informacion_pedido_sugerido = Me.Name
                           'frminformacion_pedido_sugerido_rutas.Show
                           'Me.Enabled = False
                        End If
                     End If
                     
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
         'fin de la comparacion del saldo con el oracle
         Else
            MsgBox "El saldo del cliente en ORACLE es menor al importe de las facturas, ", vbOKOnly, "ATENCION"
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

Private Sub cmd_relacion_facturas_Click()
   txt_embarque_relacion = ""
   frm_embarque_relacion.Visible = True
   txt_embarque_relacion.SetFocus
End Sub

Private Sub cmd_remision_Click()
   If IsNumeric(Me.txt_embarque_remision) Then
            Set reporte = appl.OpenReport(App.Path + "\REP_FACTURA_REMISION.rpt")
            reporte.RecordSelectionFormula = "{VW_FACTURA.INTE_EMB_EMBARQUE} = " + Me.txt_embarque_remision
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Movimientos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
   Else
      MsgBox "Embarque incorrecto", vbOKOnly, "ATENCION"
   End If
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
                                        var_catalogo_1 = rsaux4!VCHA_ART_ARTICULO_ID
                                     Else
                                        var_catalogo_2 = rsaux4!VCHA_ART_ARTICULO_ID
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
                                    var_precio_catalogo_1 = IIf(IsNull(rsaux4!floa_dli_Precio), 0, rsaux4!floa_dli_Precio)
                                 Else
                                    var_precio_catalogo_1 = 0
                                 End If
                                 rsaux4.Close
                                 rsaux4.Open "select * from tb_detalle_lista_precios where VCHA_LIS_LISTA_PRECIOS_ID = '" + var_lista_precios_cataloGos + "' and vcha_art_articulo_id = '" + var_catalogo_2 + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux4.EOF Then
                                    var_precio_catalogo_2 = IIf(IsNull(rsaux4!floa_dli_Precio), 0, rsaux4!floa_dli_Precio)
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
                                 ok = TB_LIBERA_APARTADOS.Anadir(var_almacen_OS, rs!VCHA_ART_ARTICULO_ID, 0 - rs!floa_Sal_Cantidad)
                                 var_inserta = False
                                 var_suma_cantidad = 0
                                 var_cantidad_llegar = IIf(IsNull(rs!floa_Sal_Cantidad), 0, rs!floa_Sal_Cantidad)
                                 var_cantidad = 0
                                 While var_suma_cantidad < var_cantidad_llegar
                                       rsaux2.Open "select * from tb_existencias where vcha_art_articulo_id =  '" + rs!VCHA_ART_ARTICULO_ID + "' and vcha_alm_almacen_id = '" + var_almacen_OS + "'", cnn, adOpenDynamic, adLockOptimistic
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
                                          rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id =  '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux4.EOF Then
                                             var_costo = rsaux4!mone_Art_costo_estandar
                                          Else
                                             var_costo = 0
                                          End If
                                          rsaux4.Close
                                       End If
                                       rsaux2.Close
                                       'OTORGAMIENTO DE CATALOGOS
                                       If rs!VCHA_ART_ARTICULO_ID = var_catalogo_1 Then
                                          var_importe_catalogos = var_precio_catalogo_1 * var_cantidad
                                          If var_importe_catalogos <= var_importe_disponible Then
                                             rsaux2.Open "update TB_IMPORTES_FACTURACION_CATALOGOS_TITULAR set  floa_fca_disponible = floa_fca_disponible - " + CStr(var_importe_catalogos) + " where vcha_tit_titular_id = '" + var_clave_titular + "' and inte_fca_año = " + CStr(var_año_catalogo) + " and inte_fca_mes = " + CStr(var_mes_catalogo), cnn, adOpenDynamic, adLockOptimistic
                                             var_importe_disponible = var_importe_disponible - var_importe_catalogos
                                             
                                             Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID],[INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2],[char_ped_tipo], [inte_sal_año]) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!VCHA_ART_ARTICULO_ID + "' , " + CStr(var_cantidad) + ", " + CStr(IIf(IsNull(var_costo), 0, var_costo)) + ", " + CStr(0 * var_tipo_Cambio) + ", 0, 0,  0,'" + rs!char_ped_tipo + "', " + CStr(var_año) + ")"
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
                                                Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID],[INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2],[char_ped_tipo], [inte_sal_año]) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!VCHA_ART_ARTICULO_ID + "' , " + CStr(var_cantidad_gratis) + ", " + CStr(IIf(IsNull(var_costo), 0, var_costo)) + ", " + CStr(0 * var_tipo_Cambio) + ", 0, 0,  0,'" + rs!char_ped_tipo + "', " + CStr(var_año) + ")"
                                                rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                             End If
                                             If var_renglones_factura = var_contador_renglones Then
                                                var_contador_renglones = 0
                                             End If
                                             Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID],[INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2],[char_ped_tipo], [inte_sal_año]) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!VCHA_ART_ARTICULO_ID + "' , " + CStr(var_cantidad_cobrada) + ", " + CStr(IIf(IsNull(var_costo), 0, var_costo)) + ", " + CStr(rs!floa_Sal_precio * var_tipo_Cambio) + ", " + CStr(rs!FLOA_SAL_DESCUENTO) + ", " + CStr(rs!floa_sal_promocion_1) + ",  " + CStr(rs!FLOA_SAL_PROMOCION_2) + ",'" + rs!char_ped_tipo + "', " + CStr(var_año) + ")"
                                             rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                             If var_renglones_factura = var_contador_renglones Then
                                                var_contador_renglones = 0
                                             End If
                                          End If
                                       Else
                                          If rs!VCHA_ART_ARTICULO_ID = var_catalogo_2 Then
                                             var_importe_catalogos = var_precio_catalogo_2 * var_cantidad
                                             If var_importe_catalogos <= var_importe_disponible Then
                                                rsaux2.Open "update TB_IMPORTES_FACTURACION_CATALOGOS_TITULAR set floa_fca_disponible = floa_fca_disponible - " + CStr(var_importe_catalogos) + " where vcha_tit_titular_id = '" + var_clave_titular + "' and inte_fca_año = " + CStr(var_año_catalogo) + " and inte_fca_mes = " + CStr(var_mes_catalogo), cnn, adOpenDynamic, adLockOptimistic
                                                var_importe_disponible = var_importe_disponible - var_importe_catalogos
                                                Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID],[INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2],[char_ped_tipo], [inte_sal_año]) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!VCHA_ART_ARTICULO_ID + "' , " + CStr(var_cantidad) + ", " + CStr(IIf(IsNull(var_costo), 0, var_costo)) + ", " + CStr(rs!floa_Sal_precio * var_tipo_Cambio) + ", " + CStr(rs!FLOA_SAL_DESCUENTO) + ", " + CStr(rs!floa_sal_promocion_1) + ",  " + CStr(rs!FLOA_SAL_PROMOCION_2) + ",'" + rs!char_ped_tipo + "', " + CStr(var_año) + ")"
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
                                                   Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID],[INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2],[char_ped_tipo], [inte_sal_año]) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!VCHA_ART_ARTICULO_ID + "' , " + CStr(var_cantidad_gratis) + ", " + CStr(IIf(IsNull(var_costo), 0, var_costo)) + ", " + CStr(0 * var_tipo_Cambio) + ", 0, 0,  0,'" + rs!char_ped_tipo + "', " + CStr(var_año) + ")"
                                                   rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                                End If
                                                If var_renglones_factura = var_contador_renglones Then
                                                   var_contador_renglones = 0
                                                End If
                                                Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID],[INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2],[char_ped_tipo], [inte_sal_año]) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!VCHA_ART_ARTICULO_ID + "' , " + CStr(var_cantidad_cobrada) + ", " + CStr(IIf(IsNull(var_costo), 0, var_costo)) + ", " + CStr(rs!floa_Sal_precio * var_tipo_Cambio) + ", " + CStr(rs!FLOA_SAL_DESCUENTO) + ", " + CStr(rs!floa_sal_promocion_1) + ",  " + CStr(rs!FLOA_SAL_PROMOCION_2) + ",'" + rs!char_ped_tipo + "', " + CStr(var_año) + ")"
                                                rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                                If var_renglones_factura = var_contador_renglones Then
                                                   var_contador_renglones = 0
                                                End If
                                             End If
                                          Else
                                             Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID],[INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2],[char_ped_tipo], [inte_sal_año]) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!VCHA_ART_ARTICULO_ID + "' , " + CStr(var_cantidad) + ", " + CStr(IIf(IsNull(var_costo), 0, var_costo)) + ", " + CStr(rs!floa_Sal_precio * var_tipo_Cambio) + ", " + CStr(rs!FLOA_SAL_DESCUENTO) + ", " + CStr(rs!floa_sal_promocion_1) + ",  " + CStr(rs!FLOA_SAL_PROMOCION_2) + ",'" + rs!char_ped_tipo + "', " + CStr(var_año) + ")"
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
                                    Cadena = "insert into " + App.Path + "\temp_" + Trim(var_nombre_archivo) + ".dbf (cvenota, cvecliente, clapr, canp1, canp2, canp3, canp4, canp5, canp6, prepr, cvepedido, anocosto) values ('" + Trim(Str(var_numero_folio)) + "', '" + var_clave_cliente + "', '" + Mid(Trim(rs!VCHA_ART_ARTICULO_ID), 7, 5) + "', " + Trim(CStr(rs!floa_Sal_Cantidad)) + ", 0, 0, 0, 0, 0, " + Trim(CStr(Round(rs!floa_Sal_costo, 4))) + ", " + CStr(var_numero_pedido_cliente) + ", '" + Trim(CStr(rs!INTE_sAL_AÑO)) + "')"
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
                                    ok = TB_LIBERA_APARTADOS.Anadir(var_almacen_OS, rs!VCHA_ART_ARTICULO_ID, 0 - rs!floa_Sal_Cantidad)
                                    var_inserta = False
                                    var_suma_cantidad = 0
                                    var_cantidad_llegar = IIf(IsNull(rs!floa_Sal_Cantidad), 0, rs!floa_Sal_Cantidad)
                                    var_cantidad = 0
                                    While var_suma_cantidad < var_cantidad_llegar
                                          rsaux2.Open "select * from tb_existencias where vcha_art_articulo_id =  '" + rs!VCHA_ART_ARTICULO_ID + "' and vcha_alm_almacen_id = '" + var_almacen_OS + "'", cnn, adOpenDynamic, adLockOptimistic
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
                                             rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id =  '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                             If Not rsaux4.EOF Then
                                                var_costo = rsaux4!mone_Art_costo_estandar
                                             Else
                                                var_costo = 0
                                             End If
                                             rsaux4.Close
                                          End If
                                          rsaux2.Close
                                          Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2], [char_ped_tipo],[inte_sal_año]) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!VCHA_ART_ARTICULO_ID + "' , " + Str(var_cantidad) + ", " + CStr(var_costo) + ", " + CStr(rs!floa_Sal_precio * var_tipo_Cambio) + ", " + CStr(rs!FLOA_SAL_DESCUENTO) + ", " + CStr(rs!floa_sal_promocion_1) + ",  " + CStr(rs!FLOA_SAL_PROMOCION_2) + ", '" + rs!char_ped_tipo + "', " + CStr(var_año) + ")"
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
                                       Cadena = "insert into " + App.Path + "\temp_" + Trim(var_nombre_archivo) + ".dbf (cvenota, cvecliente, clapr, canp1, canp2, canp3, canp4, canp5, canp6, prepr, cvepedido,ANOCOSTO ) values ('" + Trim(Str(var_numero_folio)) + "', '" + var_clave_cliente + "', '" + Mid(Trim(rs!VCHA_ART_ARTICULO_ID), 7, 5) + "', " + Trim(CStr(rs!floa_Sal_Cantidad)) + ", 0, 0, 0, 0, 0, " + Trim(CStr(Round(rs!floa_Sal_costo, 4))) + ", " + CStr(var_numero_pedido_cliente) + ",'" + CStr(rs!INTE_sAL_AÑO) + "')"
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
                                    ok = TB_LIBERA_APARTADOS.Anadir(var_almacen_OS, rs!VCHA_ART_ARTICULO_ID, 0 - rs!floa_Sal_Cantidad)
                                    var_inserta = False
                                    Cadena = "insert into tb_Salidas ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD],[FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2], [CHAR_PED_TIPO]) values ('" + rs(0).Value + "', '" + rs(1).Value + "', '" + rs(2).Value + "', '" + rs(3).Value + "', " + CStr(rs(4).Value) + ", '" + rs(5).Value + "' , " + Str(rs(6).Value) + ", " + CStr(rs(7).Value) + ", " + CStr(rs(8).Value * var_tipo_Cambio) + ", " + CStr(rs(9).Value) + ", " + CStr(rs(10).Value) + ",  " + CStr(rs(11).Value) + ", '" + rs!char_ped_tipo + "')"
                                    rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                    'var_inserta = TB_SALIDAS_INSERTA.Anadir(rs(0).Value, rs(1).Value, rs(2).Value, rs(3).Value, rs(4).Value, rs(5).Value, rs(6).Value, rs(7).Value, rs(8).Value * var_tipo_cambio, rs(9).Value)
                                    var_inserta = False
                                    var_inserta = TB_SALIDA_VISTAS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_numero_folio, var_clave_movimiento, rs!VCHA_ART_ARTICULO_ID, rs!floa_Sal_Cantidad, rs!floa_Sal_costo, rs!floa_Sal_precio)
                                    var_inserta = False
                                    var_inserta = TB_ARCH_COMPARACION_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_movimiento_dependencia, var_numero_folio, Date, "G", var_clave_agente, rs!VCHA_ART_ARTICULO_ID, rs!floa_Sal_costo, rs!floa_Sal_Cantidad, 0, "", var_referencia_vi)
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
                                    If var_empresa = "15" Then
                                       var_linea = var_linea + "MAQUILA DE " + UCase(IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura))
                                    Else
                                       var_linea = var_linea + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                    End If
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
                              If Trim(var_cliente_coppel) = "C000001636" Then
                                 var_linea = var_solicitud_sigo
                              Else
                                 var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                              End If
                              If Len(Trim(var_linea)) < 145 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 145
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              Print #1, var_linea + var_importe_descuento_1_str
                              If var_empresa = "18" Then
                                 var_linea = ""
                              Else
                                 If Trim(var_cliente_coppel) = "C000002947" Or Trim(var_cliente_coppel) = "C000001636" Then
                                    If Trim(var_cliente_coppel) = "C000002947" Then
                                       rsaux11.Open "select * from TB_eSTABLECIMIENTOS where vcha_esb_establecimiento_id = '" + var_establecimiento_comercial + "'", cnn, adOpenDynamic, adLockOptimistic
                                       var_linea = "ESTABLECIMIENTO: " + IIf(IsNull(rsaux11!VCHA_ESB_NOMBRE), "", rsaux11!VCHA_ESB_NOMBRE)
                                       rsaux11.Close
                                    Else
                                       var_linea = ""
                                    End If
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
                              'var_contador_promociones = 1 ' se pone para poder poner la leyenda del IVA del 16%
                              If var_contador_promociones > 0 Then
                                 If var_cliente_sigo = "C000001636" Then
                                    'Print #1, "Descuento adicional del 2%"
                                    Print #1, ""
                                 Else
                                    Print #1, var_cadena_promocion_171209
                                 End If
                              Else
                                 If var_cliente_sigo = "C000001636" Then
                                    'Print #1, "Descuento adicional del 2%"
                                    Print #1, ""
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
                                 Cadena = "insert into " + App.Path + "\temp_" + Trim(var_nombre_archivo) + ".dbf (cvenota, cvecliente, clapr, canp1, canp2, canp3, canp4, canp5, canp6, prepr, cvepedido, anocosto) values ('" + Trim(Str(var_numero_folio)) + "', '" + var_clave_cliente + "', '" + Mid(Trim(rs!VCHA_ART_ARTICULO_ID), 7, 5) + "', " + Trim(CStr(IIf(IsNull(rs!floa_Sal_Cantidad), 0, rs!floa_Sal_Cantidad))) + ", 0, 0, 0, 0, 0, " + Trim(CStr(Round(rs!floa_Sal_costo, 4))) + ", " + CStr(var_numero_pedido_cliente) + ", '" + Trim(CStr(rs!INTE_sAL_AÑO)) + "')"
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
                                    rsaux4.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + IIf(IsNull(rsaux2!VCHA_ART_ARTICULO_ID), "", rsaux2!VCHA_ART_ARTICULO_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
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
                                       If var_empresa = "15" Then
                                          var_linea = var_linea + "MAQUILA DE " + UCase(IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura))
                                       Else
                                          var_linea = var_linea + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                       End If
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
                                 If Trim(var_cliente_coppel) = "C000001636" Then
                                    var_linea = var_solicitud_sigo
                                 Else
                                    var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                                 End If
                                 If Len(Trim(var_linea)) < 145 Then
                                    For var_j = 1 + Len(Trim(var_linea)) To 145
                                        var_linea = var_linea + " "
                                    Next var_j
                                 End If
                                 Print #1, var_linea + var_importe_descuento_1_str
                                 If var_empresa = "18" Then
                                    var_linea = ""
                                 Else
                                    If Trim(var_cliente_coppel) = "C000002947" Or Trim(var_cliente_coppel) = "C000001636" Then
                                       If Trim(var_cliente_coppel) = "C000002947" Then
                                          rsaux11.Open "select * from TB_eSTABLECIMIENTOS where vcha_esb_establecimiento_id = '" + var_establecimiento_comercial + "'", cnn, adOpenDynamic, adLockOptimistic
                                          var_linea = "ESTABLECIMIENTO: " + IIf(IsNull(rsaux11!VCHA_ESB_NOMBRE), "", rsaux11!VCHA_ESB_NOMBRE)
                                          rsaux11.Close
                                       Else
                                          var_linea = ""
                                       End If
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
                                 'var_contador_promociones = 1 ' se pone para poder poner la leyenda del IVA del 16%
                                 If var_contador_promociones > 0 Then
                                    If var_cliente_sigo = "C000001636" Then
                                       'Print #1, "Descuento adicional del 2%"
                                       Print #1, ""
                                    Else
                                       Print #1, var_cadena_promocion_171209
                                    End If
                                 Else
                                    If var_cliente_sigo = "C000001636" Then
                                       'Print #1, "Descuento adicional del 2%"
                                       Print #1, ""
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
                                    If var_empresa = "15" Then
                                       var_linea = var_linea + "MAQUILA DE " + UCase(IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura))
                                    Else
                                       var_linea = var_linea + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                    End If
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
                              If Trim(var_cliente_coppel) = "C000001636" Then
                                 var_linea = var_solicitud_sigo
                              Else
                                 var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                              End If
                              If Len(Trim(var_linea)) < 145 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 145
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              Print #1, var_linea + var_importe_descuento_1_str
                              If var_empresa = "18" Then
                                 var_empresa = ""
                              Else
                                 If Trim(var_cliente_coppel) = "C000002947" Or Trim(var_cliente_coppel) = "C000001636" Then
                                    If Trim(var_cliente_coppel) = "C000002947" Then
                                       rsaux11.Open "select * from TB_eSTABLECIMIENTOS where vcha_esb_establecimiento_id = '" + var_establecimiento_comercial + "'", cnn, adOpenDynamic, adLockOptimistic
                                       var_linea = "ESTABLECIMIENTO: " + IIf(IsNull(rsaux11!VCHA_ESB_NOMBRE), "", rsaux11!VCHA_ESB_NOMBRE)
                                       rsaux11.Close
                                    Else
                                       var_linea = ""
                                    End If
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
                              'var_contador_promociones = 1 ' se pone para poder poner la leyenda del IVA del 16%
                              If var_contador_promociones > 0 Then
                                 If var_cliente_sigo = "C000001636" Then
                                    'Print #1, "Descuento adicional del 2%"
                                    Print #1, ""
                                 Else
                                    Print #1, var_cadena_promocion_171209
                                 End If
                              Else
                                 If var_cliente_sigo = "C000001636" Then
                                    'Print #1, "Descuento adicional del 2%"
                                    Print #1, ""
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

Private Sub Command6_Click()
   Me.frm_embarque_reimprimir.Visible = True
   Me.txt_embarque_reimprimir = ""
   Me.txt_embarque_reimprimir.SetFocus
End Sub

Private Sub Command7_Click()
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
                     rs.Open "execute factura_embarques '" + var_empresa + "', '" + var_unidad_organizacional + "', " + txt_numero_embarque + ", '', '','" + var_serie + "', 'FA'", cnn, adOpenDynamic, adLockOptimistic
                     ok = TB_ENC_EMBARQUE_M.Anadir(var_empresa, var_unidad_organizacional, txt_numero_embarque, "F")
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
                     
                     Cadena = "EXEC SP_CREA_TABLA_FACTURAS_CHIQUIBLANCOS " + CStr(var_consecutivo) + ",'" + var_empresa + "'," + txt_numero_embarque
                     rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     Me.frm_mensaje.Visible = False
                     rsaux3.Open "select distinct inte_car_numero, vcha_ser_Serie_id from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux3.EOF Then
                        
                        While Not rsaux3.EOF
                              MsgBox "Se va a imprimir la factura " + Trim(Str(rsaux3!inte_Car_numero)) + ", prepare la impresora", vbOKOnly, "ATENCION"
                              
                              
                              Set reporte = appl.OpenReport(App.Path + "\rep_factura_chiquiblancos.rpt")
                              reporte.RecordSelectionFormula = "{TB_TEMP_FACTURA_EMBARQUES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {TB_TEMP_FACTURA_EMBARQUES.inte_car_numero} = " + CStr(rsaux3!inte_Car_numero) + " and {TB_TEMP_FACTURA_EMBARQUES.vcha_ser_serie_id} = '" + Trim(rsaux3!vcha_Ser_Serie_id) + "' and {TB_TEMP_FACTURA_EMBARQUES.vcha_emp_empresa_id} = '" + var_empresa + "'"
                              For ntablas = 1 To reporte.Database.Tables.Count
                                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), "chiquiblancos_sid", parametros(4), parametros(5)
                              Next ntablas
                              reporte.PrintOut False
                              Set reporte = Nothing
                              
                              Set reporte = appl.OpenReport(App.Path + "\rep_factura_chiquiblancos.rpt")
                              reporte.RecordSelectionFormula = "{TB_TEMP_FACTURA_EMBARQUES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {TB_TEMP_FACTURA_EMBARQUES.inte_car_numero} = " + CStr(rsaux3!inte_Car_numero) + " and {TB_TEMP_FACTURA_EMBARQUES.vcha_ser_serie_id} = '" + Trim(rsaux3!vcha_Ser_Serie_id) + "' and {TB_TEMP_FACTURA_EMBARQUES.vcha_emp_empresa_id} = '" + var_empresa + "'"
                              For ntablas = 1 To reporte.Database.Tables.Count
                                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), "chiquiblancos_sid", parametros(4), parametros(5)
                              Next ntablas
                              reporte.PrintOut False
                              Set reporte = Nothing
                              
                              Set reporte = appl.OpenReport(App.Path + "\rep_factura_chiquiblancos.rpt")
                              reporte.RecordSelectionFormula = "{TB_TEMP_FACTURA_EMBARQUES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {TB_TEMP_FACTURA_EMBARQUES.inte_car_numero} = " + CStr(rsaux3!inte_Car_numero) + " and {TB_TEMP_FACTURA_EMBARQUES.vcha_ser_serie_id} = '" + Trim(rsaux3!vcha_Ser_Serie_id) + "' and {TB_TEMP_FACTURA_EMBARQUES.vcha_emp_empresa_id} = '" + var_empresa + "'"
                              For ntablas = 1 To reporte.Database.Tables.Count
                                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), "chiquiblancos_sid", parametros(4), parametros(5)
                              Next ntablas
                              reporte.PrintOut False
                              Set reporte = Nothing
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

Private Sub Command8_Click()
        var_Archivo = App.Path & "\facturax.bat"
        Open (App.Path & "\facturax.bat") For Output As #2
        Print #2, "ren c:\sistemas\archivo.txt archivo.ff"
        Close #2
        x = Shell(var_Archivo, vbHide)
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
   Me.frm_correo_clientes.Visible = False
   If var_empresa <> "06" Then
      If var_empresa <> "15" Then
         If var_empresa <> "31" Then
            var_numero_embarque_global = 0
         End If
      End If
   End If
   Me.frm_embarque_reimprimir.Visible = False
   If var_empresa <> "18" Then
      If var_unidad_organizacional = "23" Then
         If cnn_clientes_tiendas.State = 0 Then
            cnn_clientes_tiendas.Open var_conexion_pedidos_tiendas
            cnn_clientes_tiendas.CursorLocation = adUseClient
         End If
      End If
   End If
   If var_empresa = "31" Then
       
   End If
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
   rs.Open "select VCHA_PRI_RUTA_ENVIOS_FACTURAS from tb_principal", cnn, adOpenDynamic, adLockOptimistic
   var_ruta = IIf(IsNull(rs(0).Value), "", rs(0).Value)
   rs.Close
   Top = 0
   Left = 0
   If var_empresa = "30" Then
      rs.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockBatchOptimistic
   Else
      If var_empresa = "15" Then
         rs.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         If var_empresa = "17" Then
            rs.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockBatchOptimistic
         Else
            rs.Open "select * from tb_principal", cnn, adOpenDynamic, adLockBatchOptimistic
         End If
      End If
   End If
   var_renglones_factura = rs!INTE_PRI_RENGLONES_FACTURA
   rs.Close
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
      If var_empresa = "30" Or var_empresa = "31" Then
         If var_empresa = "30" Then
            rsaux2.Open "select * from VW_EMBARQUES_CERRADOS_NO_FACTURADOS where vcha_Emp_empresa_id = '30'", cnn, adOpenDynamic, adLockOptimistic
         Else
            If var_empresa = "15" Then
               rsaux2.Open "select * from VW_EMBARQUES_CERRADOS_NO_FACTURADOS where vcha_Emp_empresa_id = '15'", cnn, adOpenDynamic, adLockOptimistic
            Else
               rsaux2.Open "select * from VW_EMBARQUES_CERRADOS_NO_FACTURADOS where vcha_Emp_empresa_id = '31'", cnn, adOpenDynamic, adLockOptimistic
            End If
         End If
      Else
         rsaux2.Open "select * from VW_EMBARQUES_CERRADOS_NO_FACTURADOS where vcha_Emp_empresa_id <> '16'", cnn, adOpenDynamic, adLockOptimistic
      End If
      
      If Not rsaux2.EOF Then
         'Me.Enabled = False
         If var_empresa <> "15" Then
            var_activa_forma_detalle_cajas = Me.Name
            frmembarques_cerrados_no_facturados.Show 1
         End If
      End If
      rsaux2.Close
   Else
      MsgBox "No se a indicado una serie para esta Unidad organizacional", vbOKOnly, "ATENCION"
      txt_numero_embarque.Enabled = False
   End If
   rs.Close
   If var_unidad_organizacional = "01" Or var_unidad_organizacional = "02" Or var_unidad_organizacional = "03" Or var_unidad_organizacional = "04" Or var_unidad_organizacional = "06" Then
      Me.cmd_correo.Visible = False
      Me.cmd_relacion_facturas.Visible = False
      Me.cmd_nota_envio.Visible = False
      Me.cmd_embarques_cerrados.Visible = False
      Me.cmd_correo_facturacion_tiendas.Visible = False
      Me.cmd_correo_clientes.Visible = False
      Me.Command6.Visible = False
   End If
   If var_empresa = "06" Or var_empresa = "15" Or var_empresa = "17" Or var_empresa = "16" Or var_empresa = "30" Or var_empresa = "31" Or var_empresa = "32" Or var_empresa = "33" Or var_empresa = "34" Or var_empresa = "34" Or var_empresa = "35" Or var_empresa = "36" Or var_empresa = "37" Or var_empresa = "38" Or var_empresa = "39" Or var_empresa = "40" Or var_empresa = "41" Or var_empresa = "42" Or var_empresa = "43" Or var_empresa = "44" Or var_empresa = "29" Then
      Me.cmd_correo.Visible = False
      Me.cmd_relacion_facturas.Visible = False
      Me.cmd_nota_envio.Visible = False
      Me.cmd_embarques_cerrados.Visible = False
      Me.cmd_correo_facturacion_tiendas.Visible = False
      Me.cmd_correo_clientes.Visible = False
      Me.Command6.Visible = False
   End If
   If var_numero_embarque_global > 0 Then
      Me.txt_numero_embarque = var_numero_embarque_global
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
         If rsaux4.State = 1 Then
            rsaux4.Close
         End If
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
                  If rs.State = 1 Then
                     rs.Close
                  End If
                  If rs.State = 1 Then
                     rs.Close
                  End If
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
                        'se inivio
                        'Call facturas
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
                     If lv_movimientos.ListItems.Count > 0 Then
                        'lv_movimientos.SetFocus
                     End If
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
   If var_empresa = "15" Or var_empresa = "16" Or var_empresa = "30" Or var_empresa = "31" Or var_empresa = "32" Or var_empresa = "33" Or var_empresa = "34" Or var_empresa = "34" Or var_empresa = "35" Or var_empresa = "36" Or var_empresa = "37" Or var_empresa = "38" Or var_empresa = "39" Or var_empresa = "40" Or var_empresa = "41" Or var_empresa = "42" Or var_empresa = "02" Or var_empresa = "03" Or var_empresa = "06" Or var_empresa = "18" Or var_empresa = "17" Or var_empresa = "43" Or var_empresa = "44" Or var_empresa = "29" Then
      Me.cmd_imprimir.Enabled = False
      Me.cmd_factura_electronica.Enabled = True
      Me.Command6.Visible = True
   Else
      Me.cmd_imprimir.Enabled = True
      Me.cmd_factura_electronica.Enabled = False
      'Me.Command6.Visible = True
   End If
   If var_empresa = "02" Then
      Me.Command6.Visible = True
   End If
   If var_empresa = "06" And var_unidad_organizacional = "39" Or var_empresa = "31" Then
      Me.cmd_nota_envio.Visible = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
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
   If rsaux11.State = 1 Then
      rsaux11.Close
   End If
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


Private Sub Text1_KeyPress(KeyAscii As Integer)
  MsgBox Chr(KeyAscii)
End Sub

Private Sub Pdf1_GotFocus()

End Sub

Private Sub txt_embarque_activo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Dim var_contador As Double
      Dim var_tipo As String
      lv_embarques.ListItems.Clear
      cnn.CommandTimeout = 360
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
   'On Error Resume Next
   If IsNumeric(Me.txt_embarque_activo) Then
   rs.Open "select * from tb_encabezado_embarques where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque_activo, cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_embarque_cerrado = IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", Trim(rs!CHAR_EMB_ESTATUS))
   End If
   rs.Close
   si = MsgBox("¿Desea generar el reporte del embarque?", vbYesNo, "ATENCION")
   If si = 6 Then
         var_clave_movimiento_anterior = var_clave_movimiento
         
         If Trim(var_embarque_cerrado) = "I" Then
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
                  rsaux10.Open "select vcha_cli_email from tb_clientes where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux10.EOF Then
                     var_correo_electronico = IIf(IsNull(rsaux10(0).Value), "", rsaux10(0).Value)
                  Else
                     var_correo_electronico = ""
                  End If
                  rsaux10.Close
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
                        If var_empresa = "02" Then
                           Set reporte = appl.OpenReport(App.Path + "\rep_nota_envio_tiendas_resumen_reporte.rpt")
                           reporte.RecordSelectionFormula = "{VW_ENVIO_TIENDAS_RESUMEN_REPORTE.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_ENVIO_TIENDAS_RESUMEN_REPORTE.FLOA_SAL_CANTIDAD} > 0 and {VW_ENVIO_TIENDAS_RESUMEN_REPORTE.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_ENVIO_TIENDAS_RESUMEN_REPORTE.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "'"
                        Else
                           Set reporte = appl.OpenReport(App.Path + "\rep_nota_ENVIO_TIENDAS.rpt")
                           reporte.RecordSelectionFormula = "{VW_ENVIO_TIENDAS.inte_emo_numero} = " + Str(var_numero_folio) + " and {VW_ENVIO_TIENDAS.FLOA_SAL_CANTIDAD} > 0 and {VW_ENVIO_TIENDAS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_ENVIO_TIENDAS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "'"
                        End If
                        frmvistasprevias.cr.ReportSource = reporte
                        For ntablas = 1 To reporte.Database.Tables.Count
                            reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                        Next ntablas
                        frmvistasprevias.cr.ViewReport
                        frmvistasprevias.Caption = "Reporte de Movimientos"
                        frmvistasprevias.Show 1
                        Set reporte = Nothing
                        If var_empresa = "02" Then
                           var_si = MsgBox("¿Desea generar la nota de envio a detalle?", vbYesNo, "ATENCION")
                           If var_si = 6 Then
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
                           End If
                        End If
                        
                        'var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
                        var_si = 6
                        If var_si = 6 Then
                           Set reporte = appl.OpenReport(App.Path + "\rep_nota_ENVIO_TIENDAS.rpt")
                           reporte.RecordSelectionFormula = "{VW_ENVIO_TIENDAS.inte_emo_numero} = " + Str(var_numero_folio) + " and {VW_ENVIO_TIENDAS.FLOA_SAL_CANTIDAD} > 0 and {VW_ENVIO_TIENDAS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_ENVIO_TIENDAS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "'"
                           For ntablas = 1 To reporte.Database.Tables.Count
                               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                           Next ntablas
                           reporte.ExportOptions.Reset
                           'reporte.ExportOptions.FormatType = crEFTPortableDocFormat
                           'reporte.ExportOptions.FormatType = crEFTExactRichText
                           reporte.ExportOptions.FormatType = crEFTExcel80
                           reporte.ExportOptions.DestinationType = crEDTDiskFile
                           reporte.ExportOptions.UseReportDateFormat = True
                           reporte.ExportOptions.UseReportNumberFormat = True
                           archivo = "c:\reportessid\Nota_envio_" + CStr(var_numero_folio) + ".xls"
                           reporte.ExportOptions.DiskFileName = archivo
                           reporte.Export False
                           Set reporte = Nothing
                           'MsgBox "Se a terminado de guardar el archivo " + archivo
                        End If
                        
                        

                        
                        
                        
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
                           VAR_MAQUINA = UCase(fun_NombrePc)
                           If VAR_MAQUINA = "JFSERNA" Then
                              var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=DSN=dBASE Files;DBQ=" & App.Path & ";DefaultDir=" & App.Path & ";DriverId=533;MaxBufferSize=2048;PageTimeout=5;"
                           Else
                              var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + App.Path + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
                           End If
                           If rsaux2.State = 1 Then
                              rsaux2.Close
                           End If
                           rsaux2.Open "delete from nota_env", var_tabla, adOpenDynamic, adLockOptimistic
                           var_eliminar = DeleteFile(App.Path & "\temp_" + Trim(var_nombre_archivo) + ".dbf")
                           var_eliminar = DeleteFile(App.Path & "\" + Trim(var_nombre_archivo) + ".dbf")
                           var_copia = CopyFile(App.Path & "\nota_env.dbf", App.Path & "\t_" + Trim(var_nombre_archivo) + ".dbf", 1)
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
                                    Cadena = "insert into " + App.Path + "\T_" + Trim(var_nombre_archivo) + ".dbf (cvenota, cvecliente, clapr, canp1, canp2, canp3, canp4, canp5, canp6, prepr, cvepedido, anocosto, tallas, talla1, talla2, talla3, talla4, talla5, talla6) values ('" + Trim(Str(var_numero_folio)) + "', '" + var_clave_cliente + "', '" + Mid(Trim(rs!VCHA_ART_ARTICULO_ID), 7, 5) + "', " + Trim(CStr(IIf(IsNull(rs!floa_Sal_Cantidad), 0, rs!floa_Sal_Cantidad))) + ", 0, 0, 0, 0, 0, " + Trim(CStr(Round(rs!floa_Sal_costo, 4))) + ", " + CStr(var_numero_pedido_cliente) + ", '" + Trim(CStr(rs!INTE_sAL_AÑO)) + "',0,0,0,0,0,0,0)"
                                    'MsgBox var_tabla.ConnectionString
                                    rsaux2.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
                                     rs.MoveNext
                              Wend
                              rs.Close
                              var_tabla.Close
                              var_copia = CopyFile(App.Path & "\t_" + Trim(var_nombre_archivo) + ".dbf", App.Path & "\" + Trim(var_nombre_archivo) + ".dbf", 1)
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
                                 MAPISession1.NewSession = True
                                  
                                 MAPIMessages1.SessionID = MAPISession1.SessionID
                                 MAPIMessages1.Compose
                                 
                                 MAPIMessages1.RecipDisplayName = var_correo_electronico
                                 MAPIMessages1.MsgSubject = "Nota de envio " + Str(var_numero_folio)
                                 MAPIMessages1.MsgNoteText = "Se adjunta nota de envio número " + Str(var_numero_folio)
                                 MAPIMessages1.RecipAddress = var_correo_electronico
                                 
                                 MAPIMessages1.AttachmentType = mapData
                                 MAPIMessages1.AttachmentIndex = 0
                                 MAPIMessages1.AttachmentPosition = 0
                                 MAPIMessages1.AttachmentPathName = App.Path + "\" + Trim(var_nombre_archivo) + ".dbf"
                                 MAPIMessages1.AttachmentIndex = 1
                                 MAPIMessages1.AttachmentPosition = 1
                                 MAPIMessages1.AttachmentPathName = "c:\reportessid\Nota_envio_" + CStr(var_numero_folio) + ".xls"
                                 
                                 MAPIMessages1.AddressResolveUI = True
                                 MAPIMessages1.ResolveName
                                 'MAPIMessages1.ResolveName
                                 MAPIMessages1.Send False
                                 If MAPISession1.SessionID > 0 Then
                                    MAPISession1.SignOff
                                 End If
                                 'If MAPISession1.SessionID = 0 Then
                                 '   MAPISession1.SignOn
                                 'End If
                                 'MAPIMessages1.SessionID = MAPISession1.SessionID
                                 'MAPIMessages1.Compose
                                 'MAPIMessages1.RecipDisplayName = var_correo_electronico
                                 'MAPIMessages1.RecipAddress = var_correo_electronico
                                 'MAPIMessages1.AddressResolveUI = True
                                 'MAPIMessages1.ResolveName
                                 'MAPIMessages1.MsgSubject = "Nota de envio " + Str(var_numero_folio)
                                 'MAPIMessages1.MsgNoteText = "Se adjunta nota de envio número " + Str(var_numero_folio)
                                 'MAPIMessages1.AttachmentIndex = 0
                                 'MAPIMessages1.AttachmentPathName = App.Path + "\" + Trim(var_nombre_archivo) + ".dbf"
                                 'MAPIMessages1.AttachmentIndex = 1
                                 'MAPIMessages1.AttachmentPathName = "c:\reportessid\Nota_envio_" + CStr(var_numero_folio) + ".xls"
                                 
                                 
                                 'MAPIMessages1.Send True
                                 'If MAPISession1.SessionID > 0 Then
                                 '   MAPISession1.SignOff
                                 'End If
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
   Else
      MsgBox "Embarque incorrecto", vbOKOnly, "ATENCION"
   End If
End If
If KeyAscii = 27 Then
   Me.frm_envio_informacion.Visible = False
End If

End Sub

Private Sub txt_embarque_correo_clientes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_embarque_correo_clientes) Then
         var_cadena = "SELECT DISTINCT dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_EMAIL, dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID , dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE FROM  dbo.TB_CLIENTES INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_DETALLE_EMBARQUES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO = dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO "
         var_cadena = var_cadena + " WHERE (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque_correo_clientes + ") AND (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_CLIENTES.VCHA_CLI_EMAIL IS NOT NULL) AND (dbo.TB_CLIENTES.VCHA_CLI_EMAIL <> '')"
         If rs.State = 1 Then
            rs.Close
         End If
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            cnn.BeginTrans
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_consecutivo = IIf(IsNull(rsaux!NUMERO), 0, rsaux!NUMERO)
            Else
               var_consecutivo = 0
            End If
            rsaux.Close
            var_consecutivo = var_consecutivo + 1
            rsaux.Open "insert into TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES (INTE_TEM_CONSECUTIVO) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            cnn.CommandTimeout = 360
            
            
            var_cadena = "insert into TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES  SELECT TOP 100 PERCENT " + CStr(var_consecutivo) + ", dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_EMAIL, dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID,  dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, dbo.TB_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2, ((dbo.TB_SALIDAS.FLOA_SAL_PRECIO * (1 - dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1 / 100)) * (1 - dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2 / 100)) * (1 + dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_PORCENTAJE_IVA / 100) AS precio, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_PORCENTAJE_IVA, dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_PEDIDA, "
            var_cadena = var_cadena + " dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR , dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIDA, INTE_ORS_ORDEN_SURTIDO, tb_encabezado_Cartera.dtim_Car_Fecha "
            var_cadena = var_cadena + " FROM dbo.TB_CLIENTES INNER JOIN dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID ON dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_ARTICULOS INNER JOIN dbo.TB_SALIDAS ON dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID = dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND "
            var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = dbo.TB_SALIDAS.VCHA_SER_SERIE_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO AND dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO = dbo.TB_SALIDAS.INTE_CAR_NUMERO INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND  dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_SALIDAS.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO LEFT OUTER JOIN  dbo.TB_DET_ORDEN_SURTIDO ON "
            var_cadena = var_cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN = dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO AND dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.TB_DET_ORDEN_SURTIDO.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque_correo_clientes + ") AND (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') ORDER BY dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO "
            'MsgBox var_cadena
            rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            var_cadena = " SELECT DISTINCT dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN, dbo.TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES.INTE_TEM_CONSECUTIVO, dbo.TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES.INTE_EMB_EMBARQUE, dbo.TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES.VCHA_EMP_EMPRESA_ID, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE , dbo.TB_CLIENTES.VCHA_CLI_EMAIL FROM dbo.TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND "
            var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN dbo.TB_CLIENTES ON dbo.TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID WHERE     (dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES.INTE_EMB_EMBARQUE = " + Me.txt_embarque_correo_clientes + ") AND (dbo.TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES.INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + ")"
            rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  var_cadena = "SELECT *, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL AS Expr1, dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID AS Expr2 FROM dbo.TB_DET_ORDEN_SURTIDO INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_DET_ORDEN_SURTIDO.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = " + CStr(rsaux!inte_emo_numero_origen) + ") AND (dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_PEDIDA > dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIDA)"
                  rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux3.Open "insert into TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES (INTE_TEM_CONSECUTIVO, VCHA_ART_ARTICULO_ID, FLOA_SAL_cANTIDAD, vcha_cli_clave_id, vcha_cli_nombre, inte_emb_embarque, inte_ors_orden_surtido, vcha_cli_email, vcha_Art_nombre) VALUES (" + CStr(var_consecutivo) + ",'" + rsaux1!VCHA_ART_ARTICULO_ID + "'," + CStr(rsaux1!FLOA_ORS_CANTIDAD_SURTIDA - rsaux1!floa_ors_cantidad_pedida) + ", '" + rsaux!vcha_cli_clave_id + "','" + rsaux!VCHA_CLI_NOMBRE + "'," + Me.txt_embarque_correo_clientes + "," + CStr(rsaux!inte_emo_numero_origen) + ",'" + rsaux!vcha_cli_email + "','" + rsaux1!vcha_Art_nombre_español + "')", cnn, adOpenDynamic, adLockOptimistic
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux.MoveNext
            Wend
            rsaux.Close
            If rsaux2.State = 1 Then
               rsaux2.Close
            End If
            
            'var_cadena = "SELECT TB_ENCABEZADO_MOVIMIENTOS.vcha_cli_clave_id, dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR, dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIDA, dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO, dbo.VW_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES.INTE_TEM_CONSECUTIVO FROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_DET_ORDEN_SURTIDO INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO ON dbo.TB_DET_ORDEN_SURTIDO.VCHA_EMP_EMPRESA_ID = dbo.TB_ENC_ORDEN_SURTIDO.VCHA_EMP_EMPRESA_ID AND dbo.TB_DET_ORDEN_SURTIDO.VCHA_UOR_UNIDAD_ID = dbo.TB_ENC_ORDEN_SURTIDO.VCHA_UOR_UNIDAD_ID AND dbo.TB_DET_ORDEN_SURTIDO.VCHA_ALM_ALMACEN_ID = dbo.TB_ENC_ORDEN_SURTIDO.VCHA_ALM_ALMACEN_ID AND dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_DET_ORDEN_SURTIDO.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND "
            'var_cadena = var_cadena + " dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_DET_ORDEN_SURTIDO.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.VW_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.VW_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES.VCHA_EMP_EMPRESA_ID AND "
            'var_cadena = var_cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = dbo.VW_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES.INTE_EMB_EMBARQUE Where (dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR > dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIDA) And INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo)
            'rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            'While Not rsaux2.EOF
            '      rsaux3.Open "select * from tb_clientes where vcha_cli_clave_id = '" + rsaux2!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
            '      var_nombre_cliente = IIf(IsNull(rsaux3!vcha_Cli_nombre), "", rsaux3!vcha_Cli_nombre)
            '      rsaux3.Close
            '      rsaux3.Open "insert into TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES (inte_tem_consecutivo, vcha_Cli_clave_id, vcha_cli_nombre, inte_Car_numero, vcha_Art_articulo_id, vcha_Art_nombre, floa_sal_cantidad) values (" + CStr(var_consecutivo) + ",'" + rsaux2!vcha_cli_clave_id + "','" + var_nombre_cliente + "',10000000,'" + rsaux2!VCHA_ART_ARTICULO_ID + "', '" + rsaux2!vcha_Art_nombre_Español + "'," + CStr(rsaux2!FLOA_ORS_cANTIDAD_SURTIDA - rsaux2!FLOA_ORS_CANTIDAD_SURTIR) + ")", cnn, adOpenDynamic, adLockOptimistic
            '      rsaux2.MoveNext
            'Wend
            'rsaux2.Close
            
            While Not rs.EOF
                  var_correo_electronico = IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
                  var_correo_electronico = "fserna@vianney.com.mx"
                  
                  
                  Set reporte = appl.OpenReport(App.Path + "\rep_informacion_facturas_clientes.rpt")
                  reporte.RecordSelectionFormula = "{TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES.VCHA_CLI_CLAVE_ID} = '" + rs!vcha_cli_clave_id + "'"
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\informacion_facturas_cliente_" + rs!VCHA_CLI_NOMBRE + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  
                  
                  If MAPISession1.SessionID = 0 Then
                     MAPISession1.SignOn
                  End If
                  MAPIMessages1.SessionID = MAPISession1.SessionID
                  MAPIMessages1.Compose
                  MAPIMessages1.RecipDisplayName = var_correo_electronico
                  MAPIMessages1.RecipAddress = var_correo_electronico
                  MAPIMessages1.AddressResolveUI = True
                  MAPIMessages1.ResolveName
                  MAPIMessages1.MsgSubject = "Información de facturación de vianney"
                  MAPIMessages1.MsgNoteText = "Información de facturación de vianney del cliente " + rs!VCHA_CLI_NOMBRE
                  MAPIMessages1.AttachmentPathName = archivo
                  MAPIMessages1.Send False
    
                  If MAPISession1.SessionID > 0 Then
                     MAPISession1.SignOff
                  End If
                  
                  
                  
                  rs.MoveNext
            Wend
            Me.frm_correo_clientes.Visible = False
         Else
            MsgBox "El embarque no contiene clientes con direcciones electrona", vbOKOnly, "ATENCION"
            Me.frm_correo_clientes.Visible = False
         End If
         rs.Close
         rs.Open "DELETE FROM TB_TEMP_INFORMACION_FACTURAS_CORREO_CLIENTES WHERE INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      Else
         MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
         Me.frm_correo_clientes.Visible = False
      End If
      Me.frm_correo_clientes.Visible = False
   End If
   If KeyAscii = 27 Then
      Me.frm_correo_clientes.Visible = False
   End If
End Sub

Private Sub txt_embarque_correo_clientes_LostFocus()
   Me.frm_correo_clientes.Visible = False
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
               If var_cliente_coppel = "C000006202" Then
                  Print #1, ""
               Else
                  Print #1, "   Colonia:   " + IIf(IsNull(rsaux8!vcha_col_nombre), "", rsaux8!vcha_col_nombre)
               End If
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
                                 Cadena = "insert into " + App.Path + "\temp_" + Trim(var_nombre_archivo) + ".dbf (cvenota, cvecliente, clapr, canp1, canp2, canp3, canp4, canp5, canp6, prepr, cvepedido, anocosto) values ('" + Trim(Str(rs!inte_Car_numero)) + "', '" + var_clave_cliente + "', '" + Mid(Trim(rs!VCHA_ART_ARTICULO_ID), 7, 5) + "', " + Trim(CStr(IIf(IsNull(rs!floa_Sal_Cantidad), 0, rs!floa_Sal_Cantidad))) + ", 0, 0, 0, 0, 0, " + Trim(CStr(Round(var_precio_articulo, 4))) + ", 0, '" + Trim(CStr(rs!INTE_sAL_AÑO)) + "')"
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

Private Sub txt_embarque_reimprimir_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
   Dim var_nombre_unidad As String
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
   Dim var_pedido_tienada As Double
   Dim var_establecimiento_comercial As String
   
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
   Dim ndo As New aClsNodoArbolTrazabilidad
      If Trim(Me.txt_embarque_reimprimir) <> "" Then
         If IsNumeric(Me.txt_embarque_reimprimir) Then
            var_si = MsgBox("¿Desea reimprimir las facturas del embarque " + Me.txt_embarque_reimprimir + "?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_si = MsgBox("Confirmar la reimpresion de las facturas del embarque " + Me.txt_embarque_reimprimir, vbYesNo, "ATENCION")
               If var_si = 6 Then
                  var_cadena = "SELECT MIN(dbo.TB_SALIDAS.INTE_CAR_NUMERO) AS FACTURA_MENOR, MAX(dbo.TB_SALIDAS.INTE_CAR_NUMERO) AS FACTURA_MAYOR, dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID , dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE fROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_SALIDAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO WHERE dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_EMB_EMBARQUE = " + Me.txt_embarque_reimprimir + " GROUP BY dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID, dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE"
                  rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  var_posible = 0
                  If Not rsaux10.EOF Then
                     var_posible = 1
                     var_numero_menor = IIf(IsNull(rsaux10!factura_menor), 0, rsaux10!factura_menor)
                     var_numero_mayor = IIf(IsNull(rsaux10!factura_mayor), 0, rsaux10!factura_mayor)
                  Else
                     var_posible = 0
                  End If
                  If var_posible = 1 Then
                     If var_numero_mayor = 0 Or var_numero_menor = 0 Then
                        MsgBox "El embarque no se a facturado de manera correcta, lo debera de cancelar y volverlo a facturar", vbOKOnly, "ATENCION"
                     Else
                        var_si = MsgBox("Se va a imprimir de la factura " + CStr(var_numero_menor) + " a la factura " + CStr(var_numero_mayor), vbYesNo, "ATENCION")
                        If var_si = 6 Then
                           var_si = MsgBox("¿Confirmar la impresion de la factura " + CStr(var_numero_menor) + " a la factura " + CStr(var_numero_mayor), vbYesNo, "ATENCION")
                           If var_si = 6 Then
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
                                 Cadena = "EXEC SP_CREA_TABLA_FACTURAS " + CStr(var_consecutivo) + ",'" + var_empresa + "'," + Me.txt_embarque_reimprimir
                              Else
                                 Cadena = "EXEC SP_CREA_TABLA_FACTURAS_nuevo " + CStr(var_consecutivo) + ",'" + var_empresa + "'," + Me.txt_embarque_reimprimir
                              End If
                              If rsaux3.State = 1 Then
                                 rsaux3.Close
                              End If
                              rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                              If var_empresa = "02" Or var_empresa = "18" Or var_empresa = "31" Or var_empresa = "15" Or var_empresa = "06" Or var_empresa = "16" Or var_empresa = "32" Or var_empresa = "33" Or var_empresa = "34" Or var_empresa = "35" Or var_empresa = "36" Or var_empresa = "37" Or var_empresa = "38" Or var_empresa = "39" Or var_empresa = "40" Or var_empresa = "41" Or var_empresa = "42" Or var_empresa = "03" Or var_empresa = "43" Or var_empresa = "44" Or var_empresa = "17" Or var_empresa = "29" Or var_empresa = "30" Then
                                 rsaux3.Open "select distinct inte_car_numero from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque_reimprimir + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'  ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux3.EOF Then
                                    If var_empresa = "15" Or var_empresa = "16" Or var_empresa = "31" Or var_empresa = "32" Or var_empresa = "33" Or var_empresa = "34" Or var_empresa = "35" Or var_empresa = "36" Or var_empresa = "37" Or var_empresa = "38" Or var_empresa = "39" Or var_empresa = "40" Or var_empresa = "41" Or var_empresa = "42" Or var_empresa = "02" Or var_empresa = "03" Or var_empresa = "18" Or var_empresa = "06" Or var_empresa = "43" Or var_empresa = "44" Or var_empresa = "17" Or var_empresa = "29" Or var_empresa = "30" Then
                                    '''' empieza factura electronica
                                       txt_numero_embarque = Me.txt_embarque_reimprimir
                                       While Not rsaux3.EOF
                                             Call crea_factura_electronica
                                             rsaux3.MoveNext
                                       Wend
                                    ''''finaliza factura electronica
                                    Else
                                       var_Archivo = App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat"
                                       Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat") For Output As #2
                                       While Not rsaux3.EOF
                                             If rs.State = 1 Then
                                                rs.Close
                                             End If
                                             If var_empresa <> "03" Then
                                                rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_embarque_reimprimir + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' ORDER BY INTE_SAL_CONSECUTIVO_FACTURA", cnn, adOpenDynamic, adLockOptimistic
                                             Else
                                                rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_embarque_reimprimir + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' ORDER BY vcha_sal_descripcion_factura", cnn, adOpenDynamic, adLockOptimistic
                                             End If
                                             If Not rs.EOF Then
                                                'AQUI EMPIEZA LA FACTURA
                                                Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".txt") For Output As #1
                                                'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                                                'Print #1, Chr(27) + Chr(67) + Chr(44) + Chr(13)
                                                'Print #1, ""
                                                Print #1, Chr(15) + Chr(27) + Chr(64)
                                                If var_empresa = "18" Then
                                                   Print #1, ""
                                                End If
                                                Print #1, Spc(105); Str(rsaux3!inte_Car_numero)
                                                Print #1, ""
                                                Print #1, ""
                                                Print #1, Spc(105); Str(rs!INTE_CAR_PLAZO) + " DIAS DE VENCIMIENTO" + "                  " + Format(rs!dtim_Car_fecha, "Short Date")
                                                Print #1, ""
                                                'Print #1, Spc(92); Str(rs!inte_car_PLAZO) + " DIAS DE VENCIMIENTO"
                                                var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                                                var_cliente_coppel = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                                                var_cliente_sigo = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                                                For var_j = 1 + Len(Trim(var_cliente)) To 83
                                                    var_cliente = var_cliente + " "
                                                Next var_j
                                                If var_unidad_organizacional = "21" Then
                                                   var_cliente = var_cliente + "               MEXICO, D.F."
                                                Else
                                                   var_cliente = var_cliente + "               AGUASCALIENTES, AGS."
                                                End If
                                                Print #1, Spc(10); var_cliente
                                                ''' CAMBIO PARA AGREGAR COLONIA
                                                'var_domicilio = IIf(IsNull(rs!vcha_cli_direccion), "", rs!vcha_cli_direccion) + " C.P. " + IIf(IsNull(rs!vcha_cli_cp), "", rs!vcha_cli_cp)
                                                If var_cliente_coppel = "C000006202" Then
                                                   var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + ""
                                                Else
                                                   var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " COLONIA: " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
                                                End If
                                    
                                                'For var_j = 1 + Len(Trim(var_domicilio)) To 83
                                                '    var_domicilio = var_domicilio + " "
                                                'Next var_j
                                                ''' FIN DE CAMBIO PARAA AGREGAR COLONIA
                                    
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
                                       
                                                VAR_EMBARQUE = "EMB.: " + txt_embarque_reimprimir
                                                var_ordern_surtido = x
                                                Print #1, Spc(10); var_ciudad
                                                var_rfc = "RFC:  " + var_rfc
                                                var_rfc = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                                                var_establecimiento_comercial = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
                                                For var_j = 1 + Len(Trim(var_rfc)) To 89
                                                    var_rfc = var_rfc + " "
                                                Next var_j
                                                If Trim(var_cliente_coppel) = "C000002947" Or Trim(var_cliente_coppel) = "C000005566" Or Trim(var_cliente_coppel) = "C000005831" Then
                                                   rsaux5.Open "select * from tb_encabezado_pedidos where inte_ped_numero = " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))), cnn, adOpenDynamic, adLockOptimistic
                                                   If Not rsaux5.EOF Then
                                                      If Trim(var_cliente_coppel) = "C000005831" Then
                                                         var_rfc = var_rfc + "               O.C.: " + Trim(CStr(IIf(IsNull(rsaux5!VCHA_PED_PEDIDO_EXTERNO), "", rsaux5!VCHA_PED_PEDIDO_EXTERNO))) + " "
                                                      Else
                                                         var_rfc = var_rfc + "               PED.: " + Trim(CStr(IIf(IsNull(rsaux5!VCHA_PED_PEDIDO_EXTERNO), "", rsaux5!VCHA_PED_PEDIDO_EXTERNO))) + " "
                                                      End If
                                                   End If
                                                   rsaux5.Close
                                                Else
                                                   var_rfc = var_rfc + "               PED.: " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))) + " "
                                                End If
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
                                                       If var_empresa = "15" Then
                                                          var_linea = var_linea + "MAQUILA DE " + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                                       Else
                                                          var_linea = var_linea + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                                       End If
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
                                                If var_empresa = "02" Then
                                                   var_descuento_leyenda = 0
                                                   var_descuento_leyenda = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
                                                   If Trim(var_cliente_coppel) = "C000005566" Then
                                                      rsaux11.Open "select * from vw_establecimientos_direcciones where vcha_esb_establecimiento_id = '" + var_establecimiento_comercial + "'", cnn, adOpenDynamic, adLockOptimistic
                                                      var_linea = "DIRECCION DE ENTREGA: " + IIf(IsNull(rsaux11!vcha_esb_domicilio), "", rsaux11!vcha_esb_domicilio) + " COLONIA: " + IIf(IsNull(rsaux11!vcha_col_nombre), "", rsaux11!vcha_col_nombre)
                                                      Print #1, var_linea
                                                      var_linea = IIf(IsNull(rsaux11!vcha_ciu_nombre), "", rsaux11!vcha_ciu_nombre) + ", " + IIf(IsNull(rsaux11!vcha_est_nombre), "", rsaux11!vcha_est_nombre) + ", " + IIf(IsNull(rsaux11!vcha_pai_nombre), "", rsaux11!vcha_pai_nombre) + " C.P. " + IIf(IsNull(rsaux11!vcha_esb_cp), "", rsaux11!vcha_esb_cp) + " Tel: " + IIf(IsNull(rsaux11!vcha_esb_telefono), "", rsaux11!vcha_esb_telefono)
                                                      Print #1, var_linea
                                                      var_linea = ""
                                                      rsaux11.Close
                                                   Else
                                                      If var_descuento_leyenda >= 13 Then
                                                         If Trim(var_cliente_coppel) = "C000001636" Then
                                                            var_linea = var_solicitud_sigo
                                                         Else
                                                            var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                                                         End If
                                                         If Len(Trim(var_linea)) < 145 Then
                                                            For var_j = 1 + Len(Trim(var_linea)) To 145
                                                                var_linea = var_linea + " "
                                                            Next var_j
                                                         End If
                                                         Print #1, var_linea + var_importe_descuento_1_str
                                                         If var_empresa = "18" Then
                                                            var_linea = ""
                                                         Else
                                                            If Trim(var_cliente_coppel) = "C000002947" Or Trim(var_cliente_coppel) = "C000001636" Then
                                                               If Trim(var_cliente_coppel) = "C000002947" Then
                                                                  rsaux11.Open "select * from TB_eSTABLECIMIENTOS where vcha_esb_establecimiento_id = '" + var_establecimiento_comercial + "'", cnn, adOpenDynamic, adLockOptimistic
                                                                  var_linea = "ESTABLECIMIENTO: " + IIf(IsNull(rsaux11!VCHA_ESB_NOMBRE), "", rsaux11!VCHA_ESB_NOMBRE)
                                                                  rsaux11.Close
                                                               Else
                                                                  var_linea = ""
                                                               End If
                                                            Else
                                                               var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%" + " " + var_leyenda_sorteo
                                                            End If
                                                         End If
                                                         If Len(Trim(var_linea)) < 145 Then
                                                            For var_j = 1 + Len(Trim(var_linea)) To 145
                                                                var_linea = var_linea + " "
                                                            Next var_j
                                                         End If
                                                      Else
                                                         If Trim(var_cliente_coppel) = "C000001636" Then
                                                            var_linea = var_solicitud_sigo
                                                         Else
                                                            var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                                                         End If
                                                         If Len(Trim(var_linea)) < 145 Then
                                                            For var_j = 1 + Len(Trim(var_linea)) To 145
                                                                var_linea = var_linea + " "
                                                            Next var_j
                                                         End If
                                                         Print #1, var_linea + var_importe_descuento_1_str
                                                         If var_empresa = "18" Then
                                                            var_linea = ""
                                                         Else
                                                            If Trim(var_cliente_coppel) = "C000002947" Or Trim(var_cliente_coppel) = "C000001636" Then
                                                               If Trim(var_cliente_coppel) = "C000002947" Then
                                                                  rsaux11.Open "select * from TB_eSTABLECIMIENTOS where vcha_esb_establecimiento_id = '" + var_establecimiento_comercial + "'", cnn, adOpenDynamic, adLockOptimistic
                                                                  var_linea = "ESTABLECIMIENTO: " + IIf(IsNull(rsaux11!VCHA_ESB_NOMBRE), "", rsaux11!VCHA_ESB_NOMBRE)
                                                                  rsaux11.Close
                                                               Else
                                                                  var_linea = ""
                                                               End If
                                                            Else
                                                               var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%" + " " + var_leyenda_sorteo
                                                            End If
                                                         End If
                                                         If Len(Trim(var_linea)) < 145 Then
                                                            For var_j = 1 + Len(Trim(var_linea)) To 145
                                                                var_linea = var_linea + " "
                                                            Next var_j
                                                         End If
                                                      End If
                                                   End If ' comercial
                                                Else
                                                   If Trim(var_cliente_coppel) = "C000001636" Then
                                                      var_linea = var_solicitud_sigo
                                                   Else
                                                      var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                                                   End If
                                                   If Len(Trim(var_linea)) < 145 Then
                                                      For var_j = 1 + Len(Trim(var_linea)) To 145
                                                          var_linea = var_linea + " "
                                                      Next var_j
                                                   End If
                                                   Print #1, var_linea + var_importe_descuento_1_str
                                                   If var_empresa = "18" Then
                                                      var_linea = ""
                                                   Else
                                                      If Trim(var_cliente_coppel) = "C000002947" Or Trim(var_cliente_coppel) = "C000001636" Then
                                                         If Trim(var_cliente_coppel) = "C000002947" Then
                                                            rsaux11.Open "select * from TB_eSTABLECIMIENTOS where vcha_esb_establecimiento_id = '" + var_establecimiento_comercial + "'", cnn, adOpenDynamic, adLockOptimistic
                                                            var_linea = "ESTABLECIMIENTO: " + IIf(IsNull(rsaux11!VCHA_ESB_NOMBRE), "", rsaux11!VCHA_ESB_NOMBRE)
                                                            rsaux11.Close
                                                         Else
                                                            var_linea = ""
                                                         End If
                                                      Else
                                                         var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%" + " " + var_leyenda_sorteo
                                                      End If
                                                   End If
                                                   If Len(Trim(var_linea)) < 145 Then
                                                      For var_j = 1 + Len(Trim(var_linea)) To 145
                                                          var_linea = var_linea + " "
                                                      Next var_j
                                                   End If
                                                End If
                                                If var_empresa = "02" Then
                                                   If Trim(var_cliente_coppel) <> "C000002947" Then
                                                      var_linea = var_linea + var_importe_descuento_2_str
                                                   End If
                                                   If Trim(var_cliente_coppel) = "C000005566" Then
                                                   Else
                                                      Print #1, var_linea
                                                   End If
                                                   'var_contador_promociones = 1 ' se pone para poder poner la leyenda del IVA del 16%
                                                   If var_contador_promociones > 0 Then
                                                      If var_cliente_sigo = "C000001636" Then
                                                         'Print #1, "Descuento adicional del 2%"
                                                         Print #1, ""
                                                      Else
                                                         Print #1, var_cadena_promocion_171209
                                                      End If
                                                   Else
                                                      If var_cliente_sigo = "C000001636" Then
                                                         'Print #1, "Descuento adicional del 2%"
                                                         Print #1, ""
                                                      Else
                                                         Print #1, ""
                                                      End If
                                                   End If
                                                Else
                                                   var_linea = var_linea + var_importe_descuento_2_str
                                                   Print #1, var_linea
                                                   'var_contador_promociones = 1 ' se pone para poder poner la leyenda del IVA del 16%
                                                   If var_contador_promociones > 0 Then
                                                      If var_cliente_sigo = "C000001636" Then
                                                         'Print #1, "Descuento adicional del 2%"
                                                         Print #1, ""
                                                      Else
                                                         Print #1, var_cadena_promocion_171209
                                                      End If
                                                   Else
                                                      If var_cliente_sigo = "C000001636" Then
                                                         'Print #1, "Descuento adicional del 2%"
                                                         Print #1, ""
                                                      Else
                                                         Print #1, ""
                                                      End If
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
                                                If var_empresa = "31" Then
                                                   var_linea = "                                                       " + CStr(var_dia) + "     " + CStr(var_mes)
                                                Else
                                                   var_linea = "                                                             " + CStr(var_dia) + "     " + CStr(var_mes)
                                                End If
                                    
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
                                                If var_empresa = "31" Then
                                                   Print #1, Spc(10); Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                                                   Print #1, Spc(10); Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre))
                                                   Print #1, Spc(10); Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre))
                                                Else
                                                   Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                                                   Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre))
                                                   Print #1, Spc(5); Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre))
                                                End If
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
                                 End If
                                 rsaux3.Close
                                 'Aqui se termina de imprimir la factura
                                 '''' AQUI DEBE DE IR EL CORREO DE LAS VENTAS DE TIENDAS
                              If var_si_correo_ft = 1 Then
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
                                          If var_cliente_coppel = "C000006202" Then
                                             Print #1, ""
                                          Else
                                             Print #1, "   Colonia:   " + IIf(IsNull(rsaux8!vcha_col_nombre), "", rsaux8!vcha_col_nombre)
                                          End If
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
                              End If
                              ''''' hasta aqui termina el correo de ventas de tiendas
                           End If
                       
                           If var_empresa = "03000" Then
                              rsaux3.Open "select distinct inte_car_numero from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_embarque_reimprimir + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 var_Archivo = App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat"
                                 Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_Car_numero)) + ".bat") For Output As #2
                                 While Not rsaux3.EOF
                                       If rs.State = 1 Then
                                          rs.Close
                                       End If
                                       If var_empresa <> "03" Then
                                          rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_embarque_reimprimir + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY INTE_SAL_CONSECUTIVO_FACTURA", cnn, adOpenDynamic, adLockOptimistic
                                       Else
                                          rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_embarque_reimprimir + " and inte_Car_numero = " + Str(rsaux3!inte_Car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY vcha_sal_descripcion_factura", cnn, adOpenDynamic, adLockOptimistic
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
                                          var_cliente_sigo = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                                          var_cliente_coppel = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
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
                              
                                          VAR_EMBARQUE = "EMB.: " + txt_embarque_reimprimir
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
                                                 If var_empresa = "15" Then
                                                    var_linea = var_linea + "MAQUILA DE " + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                                 Else
                                                    var_linea = var_linea + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                                 End If
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
                                          If Trim(var_cliente_coppel) = "C000001636" Then
                                             var_linea = var_solicitud_sigo
                                          Else
                                             var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                                          End If
                                          If Len(Trim(var_linea)) < 145 Then
                                             For var_j = 1 + Len(Trim(var_linea)) To 145
                                                 var_linea = var_linea + " "
                                              Next var_j
                                          End If
                                          Print #1, var_linea + var_importe_descuento_1_str
                                          If var_empresa = "18" Then
                                             var_linea = ""
                                          Else
                                             If Trim(var_cliente_coppel) = "C000002947" Or Trim(var_cliente_coppel) = "C000001636" Then
                                                If Trim(var_cliente_coppel) = "C000002947" Then
                                                   rsaux11.Open "select * from TB_eSTABLECIMIENTOS where vcha_esb_establecimiento_id = '" + var_establecimiento_comercial + "'", cnn, adOpenDynamic, adLockOptimistic
                                                   var_linea = "ESTABLECIMIENTO: " + IIf(IsNull(rsaux11!VCHA_ESB_NOMBRE), "", rsaux11!VCHA_ESB_NOMBRE)
                                                   rsaux11.Close
                                                Else
                                                   var_linea = ""
                                                End If
                                             Else
                                                var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%" + " " + var_leyenda_sorteo
                                          End If
                                       End If
                                       If Len(Trim(var_linea)) < 145 Then
                                          For var_j = 1 + Len(Trim(var_linea)) To 145
                                              var_linea = var_linea + " "
                                          Next var_j
                                       End If
                                       var_linea = var_linea + var_importe_descuento_2_str
                                       Print #1, var_linea
                                       'var_contador_promociones = 1 ' se pone para poder poner la leyenda del IVA del 16%
                                       If var_contador_promociones > 0 Then
                                          If var_cliente_sigo = "C000001636" Then
                                             'Print #1, "Descuento adicional del 2%"
                                             Print #1, ""
                                          Else
                                             Print #1, var_cadena_promocion_171209
                                          End If
                                       Else
                                          If var_cliente_sigo = "C000001636" Then
                                             'Print #1, "Descuento adicional del 2%"
                                             Print #1, ""
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
                          rsaux3.Open "DELETE FROM TB_TEMP_SALIDAS_FACTURACION WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                          MsgBox "Se a terminado el proceso de facturación", vbOKOnly, "ATENCION"
                          If var_empresa = "02" Then
                             var_activa_forma_informacion_pedido_sugerido = Me.Name
                             frminformacion_pedido_sugerido_rutas.Show
                             Me.Enabled = False
                          End If
                          If var_trazabilidad = 1000 Then
                             If cnn_trazabilidad.State = 0 Then
                                cnn_trazabilidad.Open
                             End If
                             rsaux10.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                             var_nombre_unidad = ""
                             If Not rsaux10.EOF Then
                                var_nombre_unidad = IIf(IsNull(rsaux10!VCHA_UOR_NOMBRE), "", rsaux10!VCHA_UOR_NOMBRE)
                             End If
                             rsaux10.Close
                         
                        
                             var_cadena = "SELECT     dbo.TB_SALIDAS.VCHA_SER_SERIE_ID, dbo.TB_SALIDAS.INTE_CAR_NUMERO, dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE FROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_SALIDAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND"
                             var_cadena = var_cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO WHERE     (dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque_reimprimir + ")"
                             rsaux9.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                             While Not rsaux9.EOF
                                   If Not ndo.generarIdentificadorInformacionNodo(cnn_trazabilidad) Then
                                      MsgBox "No se pudo generar la trazabilidad", vbOKOnly, "ATENCION"
                                   Else
                                      ndo.nodoTipo = "E"
                                      ndo.organizacion = var_nombre_unidad
                                      ndo.eventoTipo = "FACTURACIÓN"
                                      ndo.eventoNumero = IIf(IsNull(rsaux9!inte_Car_numero), 0, rsaux9!inte_Car_numero)
                                      If Not ndo.registrarInformacionNodo(cnn_trazabilidad) Then
                                         MsgBox "No se pudo registrar la información del nodo de trazabilidad", vbOKOnly, "ATENCION"
                                      Else
                                         ndo.nodoTipo = "E"
                                         ndo.nodoPadreTipo = "N"
                                         ndo.informacionNodoPadreIdentificador = "0"
                                         If Not ndo.registrarNodo(cnn_trazabilidad) Then
                                            MsgBox "No se pudo registrar nodo de trazabilidad", vbOKOnly, "ATENCION"
                                         Else
                                            ndo.nodoPadreTipo = "E"
                                            ndo.informacionNodoPadreIdentificador = ndo.informacionNodoIdentificador
                                            ndo.nodoTipo = "L"
                                            var_cadena = "SELECT DISTINCT dbo.TB_SALIDAS.VCHA_SER_SERIE_ID, dbo.TB_SALIDAS.INTE_CAR_NUMERO, dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE,dbo.TB_TRAZABILIDAD_CODIGOS.INTE_TRA_NODO_IDENTIFICADOR , dbo.TB_TRAZABILIDAD_CODIGOS.INTE_TRA_LOTE FROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_SALIDAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND"
                                            var_cadena = var_cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO INNER JOIN dbo.TB_TRAZABILIDAD_CODIGOS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_TRAZABILIDAD_CODIGOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = dbo.TB_TRAZABILIDAD_CODIGOS.INTE_EMB_EMBARQUE AND dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.TB_TRAZABILIDAD_CODIGOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque_reimprimir + ") and (dbo.tb_salidas.inte_car_numero = " + CStr(IIf(IsNull(rsaux9!inte_Car_numero), 0, rsaux9!inte_Car_numero)) + ")"
                                            rsaux7.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                          
                                            While Not rsaux7.EOF
                                                  ndo.informacionNodoIdentificador = rsaux7!inte_tra_nodo_identificador
                                                  If Not ndo.registrarNodo(cnn_trazabilidad) Then
                                                     MsgBox "No se pudo registrar nodo de trazabilidad", vbOKOnly, "ATENCION"
                                                  End If
                                                  rsaux7.MoveNext
                                            Wend
                                            rsaux7.Close
                                         End If
                                      End If
                                   End If
                                   rsaux9.MoveNext
                             Wend
                             rsaux9.Close
                             cnn_trazabilidad.Close
                             
                          End If
                          End If
                          End If
                       End If
                  End If
               End If
            End If
         End If
      End If
      Me.frm_embarque_reimprimir.Visible = False
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

Private Sub txt_embarque_reimprimir_LostFocus()
   Me.frm_embarque_reimprimir.Visible = False
End Sub

Private Sub txt_embarque_relacion_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
   
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
        
      cnn.CommandTimeout = 360
      If var_empresa = "03" Then
         var_cadena = "SELECT dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID, dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID , dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO FROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON  dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_ALM_ALMACEN_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO WHERE  (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque_relacion + ") AND (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C') OR "
         var_cadena = var_cadena + " (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque_relacion + ") AND (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND  (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL)"
         var_facturas_cadena = ""
         rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            While Not rsaux.EOF
                  var_serie_cadena = rsaux!vcha_Ser_Serie_id
                  If var_facturas_cadena = "" Then
                     var_facturas_cadena = CStr(rsaux!inte_Car_numero)
                 Else
                     var_facturas_cadena = var_facturas_cadena + ", " + CStr(rsaux!inte_Car_numero)
                  End If
                  rsaux.MoveNext
            Wend
            rsaux.Close
            var_facturas_cadena = "(" + var_facturas_cadena + ")"
            rs.Open "select * from vw_facturas_distintas where VCHA_EMP_EMPRESA_ID ='" + var_empresa + "' and inte_Car_numero in " + var_facturas_cadena + " and vcha_Ser_serie_id = '" + var_serie_cadena + "'", cnn, adOpenDynamic, adLockOptimistic
         Else
            rs.Open "select * from vw_facturas_distintas where VCHA_EMP_EMPRESA_ID ='" + var_empresa + "' and inte_emb_embarque = " + txt_embarque_relacion, cnn, adOpenDynamic, adLockOptimistic
         End If
      Else
         var_cadena = "SELECT dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID, dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID , dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO FROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON  dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_ALM_ALMACEN_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO WHERE  (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque_relacion + ") AND (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C') OR "
         var_cadena = var_cadena + " (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque_relacion + ") AND (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND  (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL)"
         var_facturas_cadena = ""
         rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            While Not rsaux.EOF
                  var_serie_cadena = rsaux!vcha_Ser_Serie_id
                  If var_facturas_cadena = "" Then
                     var_facturas_cadena = CStr(rsaux!inte_Car_numero)
                 Else
                     var_facturas_cadena = var_facturas_cadena + ", " + CStr(rsaux!inte_Car_numero)
                  End If
                  rsaux.MoveNext
            Wend
            rsaux.Close
            var_facturas_cadena = "(" + var_facturas_cadena + ")"
            rs.Open "select * from vw_facturas_distintas where VCHA_EMP_EMPRESA_ID ='" + var_empresa + "' and inte_Car_numero in " + var_facturas_cadena + " and vcha_Ser_serie_id = '" + var_serie_cadena + "'", cnn, adOpenDynamic, adLockOptimistic
         Else
            rs.Open "select * from vw_facturas_distintas where VCHA_EMP_EMPRESA_ID ='" + var_empresa + "' and inte_emb_embarque = " + txt_embarque_relacion, cnn, adOpenDynamic, adLockOptimistic
         End If
      End If
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
         x = 1
         If x = 1 Then
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
            
            x = 0
            If x = 1 Then
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
            End If
            
            
            rsaux4.MoveNext
         Wend
         rsaux4.Close
         End If
      Else
         MsgBox "No existen facturas en el embarque indicado", vbOKOnly, "ATENCION"
      End If
      rs.Close
      frm_embarque_relacion.Visible = False
   End If
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
         rsaux4.Open "select * from VW_EMBARQUES_ACTIVOS_2 where vcha_emp_empresa_id = '" + var_empresa + "' AND (CHAR_MOV_DOCUMENTO = 'F' or CHAR_MOV_DOCUMENTO = 'V') and inte_emb_embarque = " + txt_numero_embarque, cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux4.EOF Then
            If rs.State = 1 Then
              rs.Close
            End If
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
                        If rsaux3.State = 1 Then
                           rsaux3.Close
                        End If
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
                        'se inivio
                        'Call facturas
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
   cnn.CommandTimeout = 360
   
   var_cadena = " SELECT     dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO,  dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_CLI_CLAVE_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_PRO_PROVEEDOR_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_ALMACEN_ORIGEN, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_ALMACEN_DESTINO, dbo.TB_ENCABEZADO_MOVIMIENTOS.CHAR_EMO_ESTATUS, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_AUD_USUARIO, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_AUD_MAQUINA, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_FACTURA, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_NOTA_CREDITO, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_MOVIMIENTO_ORIGEN, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA_FINALIZO, "
   var_cadena = var_cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.CHAR_EMO_BLOQUEADO, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_AGE_AGENTE_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_SUPERVISOR_CANCELO, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_SUPERVISOR_CANCELO_2, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REPORTE, dbo.TB_ENCABEZADO_MOVIMIENTOS.FLOA_EMO_DESCUENTO_1, dbo.TB_ENCABEZADO_MOVIMIENTOS.FLOA_EMO_DESCUENTO_2, dbo.TB_ENCABEZADO_MOVIMIENTOS.FLOA_EMO_DESCUENTO_3, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MON_MONEDA_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.FLOA_EMO_TIPO_CAMBIO, dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES.floa_sal_precio, dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES.FLOA_SAL_CANTIDAD, dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES.VCHA_AGR_AGRUPADOR_ID, dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES.VCHA_SAL_DESCRIPCION_FACTURA, "
   var_cadena = var_cadena + " dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES.VCHA_CAR_DOCUMENTO, dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES.VCHA_SER_SERIE_ID,"
   var_cadena = var_cadena + " dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES.INTE_CAR_NUMERO, dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES.FLOA_SAL_PROMOCION_1, dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES.FLOA_SAL_PROMOCION_2, dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES.VCHA_CAT_CATALOGO_ID, dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES.FLOA_SAL_DESCUENTO_1, dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES.FLOA_SAL_DESCUENTO_2 FROM  dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND "
   var_cadena = var_cadena + "  dbo.VW_SUMATORIA_SALIDAS_AGRUPADORES.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO where dbo.TB_ENCABEZADO_MOVIMIENTOS.vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and dbo.TB_ENCABEZADO_MOVIMIENTOS.vcha_mov_movimiento_id = '" + var_clave_mov + "' and dbo.TB_ENCABEZADO_MOVIMIENTOS.inte_Emo_numero = " + Str(var_numero_mov) + " and floa_Sal_Cantidad > 0"
   
   'rsaux2.Open "select * from vw_sumatoria_salidas_total where vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_mov + "' and inte_emo_numero = " + Str(var_numero_mov) + " and floa_sal_cantidad > 0", cnn, adOpenDynamic, adLockOptimistic
   rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
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
            Set list_item = lv_detalle.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
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
