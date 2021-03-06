VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmentradas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   Icon            =   "frmentradas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_crossdocking 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1170
      Picture         =   "frmentradas.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Crossdocking"
      Top             =   705
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame frmnumero_serie 
      Height          =   1065
      Left            =   2700
      TabIndex        =   56
      Top             =   3645
      Width           =   6570
      Begin VB.TextBox txt_numero_serie 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   75
         TabIndex        =   57
         Top             =   435
         Width           =   6405
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   " N?mero de Serie"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   58
         Top             =   120
         Width           =   6510
      End
   End
   Begin VB.Frame frm_articulo_caja 
      Height          =   1545
      Left            =   1980
      TabIndex        =   50
      Top             =   3540
      Width           =   8295
      Begin VB.TextBox txt_codigo_caja 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1590
         TabIndex        =   53
         Top             =   915
         Width           =   3375
      End
      Begin VB.TextBox txt_articulo_caja 
         Enabled         =   0   'False
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
         Left            =   90
         TabIndex        =   52
         Top             =   435
         Width           =   7965
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "C?digo del Art?culo:"
         Height          =   195
         Left            =   135
         TabIndex        =   54
         Top             =   1050
         Width           =   1395
      End
      Begin VB.Label lbl_articulo_caja 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   51
         Top             =   120
         Width           =   8220
      End
   End
   Begin VB.Frame frm_peso_caja 
      Height          =   885
      Left            =   4920
      TabIndex        =   47
      Top             =   4425
      Width           =   3165
      Begin VB.TextBox txt_peso_caja 
         Height          =   315
         Left            =   210
         MaxLength       =   10
         TabIndex        =   49
         Top             =   465
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Peso de la Caja"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   6
         Left            =   30
         TabIndex        =   48
         Top             =   120
         Width           =   3090
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11160
      Picture         =   "frmentradas.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Salir"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1170
      Picture         =   "frmentradas.frx":1006
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Cancelar Movimiento Alt + C"
      Top             =   705
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      Picture         =   "frmentradas.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Imprimir Movimiento Alt + I"
      Top             =   705
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   510
      Picture         =   "frmentradas.frx":120A
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Buscar Movimiento Alt + B"
      Top             =   705
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   180
      Picture         =   "frmentradas.frx":130C
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Nuevo Movimiento Alt + N"
      Top             =   705
      Width           =   330
   End
   Begin VB.TextBox txt_tipo_documento 
      Height          =   285
      Left            =   3135
      TabIndex        =   41
      Top             =   750
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox txt_clave_movimiento 
      Height          =   285
      Left            =   2250
      TabIndex        =   40
      Top             =   750
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   585
      TabIndex        =   23
      Top             =   1095
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         MaxLength       =   10
         TabIndex        =   25
         Top             =   510
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   2
         Left            =   30
         TabIndex        =   24
         Top             =   120
         Width           =   3075
      End
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   12405
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1785
      Width           =   2100
   End
   Begin MSComDlg.CommonDialog cmdentradas 
      Left            =   3105
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Busqueda de archivo"
   End
   Begin VB.Frame Frame3 
      Height          =   1350
      Index           =   4
      Left            =   9135
      TabIndex        =   17
      Top             =   2235
      Width           =   2355
      Begin VB.Label lbl_recibidos 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   105
         TabIndex        =   21
         Top             =   525
         Width           =   2115
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad Recibida"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   5
         Left            =   30
         TabIndex        =   18
         Top             =   120
         Width           =   2280
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1350
      Index           =   3
      Left            =   6615
      TabIndex        =   15
      Top             =   2235
      Width           =   2445
      Begin VB.Label lbl_enviados 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   135
         TabIndex        =   20
         Top             =   525
         Width           =   2205
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad Enviada"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   4
         Left            =   30
         TabIndex        =   16
         Top             =   120
         Width           =   2370
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1125
      Index           =   0
      Left            =   6615
      TabIndex        =   10
      Top             =   1095
      Width           =   4845
      Begin VB.TextBox txt_folio 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   1
         Left            =   30
         TabIndex        =   11
         Top             =   120
         Width           =   4755
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   585
      Width           =   11460
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   60
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
            Picture         =   "frmentradas.frx":140E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas.frx":1CE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas.frx":25C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas.frx":2B5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas.frx":343A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas.frx":3D14
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas.frx":45EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas.frx":4700
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas.frx":4812
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas.frx":4924
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas.frx":4A36
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas.frx":4B48
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   90
      TabIndex        =   19
      Top             =   960
      Width           =   11475
   End
   Begin VB.Frame Frame3 
      Height          =   2490
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   1095
      Width           =   6480
      Begin VB.TextBox txt_pedimento 
         Height          =   315
         Left            =   3540
         TabIndex        =   62
         Top             =   1725
         Width           =   2850
      End
      Begin VB.TextBox txt_destino 
         Enabled         =   0   'False
         Height          =   315
         Left            =   930
         TabIndex        =   39
         Top             =   420
         Width           =   5475
      End
      Begin VB.TextBox txt_factura 
         Height          =   315
         Left            =   930
         TabIndex        =   2
         Top             =   1740
         Width           =   1230
      End
      Begin VB.TextBox txt_transporto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   930
         TabIndex        =   37
         Top             =   2070
         Width           =   5475
      End
      Begin VB.TextBox txt_referencia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   930
         MaxLength       =   50
         TabIndex        =   35
         Top             =   1410
         Width           =   2625
      End
      Begin VB.TextBox txt_archivo 
         Height          =   315
         Left            =   930
         TabIndex        =   1
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txt_origen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   930
         TabIndex        =   32
         Top             =   750
         Width           =   5475
      End
      Begin VB.Label lbl_transito 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3765
         TabIndex        =   65
         Top             =   1110
         Width           =   2535
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Transito:"
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
         Left            =   2790
         TabIndex        =   64
         Top             =   1117
         Width           =   930
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pedimento:"
         Height          =   195
         Left            =   2685
         TabIndex        =   63
         Top             =   1785
         Width           =   795
      End
      Begin VB.Label lbl_factura 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Factura:"
         Height          =   195
         Left            =   90
         TabIndex        =   38
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Transporto:"
         Height          =   195
         Left            =   90
         TabIndex        =   36
         Top             =   2130
         Width           =   810
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Referencia:"
         Height          =   195
         Left            =   90
         TabIndex        =   34
         Top             =   1470
         Width           =   825
      End
      Begin VB.Label lbl_archivo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Archivo:"
         Height          =   195
         Left            =   90
         TabIndex        =   33
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label lbl_origen 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor:"
         Height          =   195
         Left            =   90
         TabIndex        =   31
         Top             =   810
         Width           =   780
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   26
         Top             =   480
         Width           =   585
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   13
         Top             =   120
         Width           =   6405
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3690
      Left            =   135
      TabIndex        =   3
      Top             =   3525
      Width           =   11370
      Begin VB.CommandButton cmd_tipo_lectura 
         Height          =   375
         Left            =   9780
         Picture         =   "frmentradas.frx":4C5A
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   435
         Width           =   585
      End
      Begin VB.CommandButton cmd_pasar_todos 
         Height          =   390
         Left            =   10395
         Picture         =   "frmentradas.frx":4D5C
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   450
         Visible         =   0   'False
         Width           =   585
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
         Left            =   5505
         TabIndex        =   5
         Top             =   450
         Width           =   1890
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   7050
         TabIndex        =   27
         Top             =   1680
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   75
            TabIndex        =   29
            Top             =   375
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   "Cantidad a eliminar"
            ForeColor       =   &H8000000E&
            Height          =   225
            Index           =   3
            Left            =   0
            TabIndex        =   28
            Top             =   15
            Width           =   2895
         End
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
         TabIndex        =   4
         Top             =   390
         Width           =   3195
      End
      Begin MSComctlLib.ListView lv_entradas 
         Height          =   2655
         Left            =   90
         TabIndex        =   22
         Top             =   1050
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   4683
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
         NumItems        =   17
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "C?digo"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripci?n"
            Object.Width           =   8696
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Enviados"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Recibidos"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Movimiento"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Faltan"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Costo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "precio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "lote"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "consecutivo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "a?o"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "CAJA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "PESO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "P_RC_LINEA_ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "P_RC_NUMERO_LINEA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "vcha_com_tipo_lectura"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "PrecioEI"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_cancelado 
         Alignment       =   2  'Center
         Caption         =   "MOVIMIENTO CANCELADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   4965
         TabIndex        =   55
         Top             =   390
         Width           =   6045
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   4800
         TabIndex        =   30
         Top             =   570
         Width           =   675
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de Art?culos"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   8
         Top             =   120
         Width           =   11280
      End
      Begin VB.Label lbl_tipo 
         AutoSize        =   -1  'True
         Caption         =   "C?digo del Art?culo:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   570
         Width           =   1395
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   0
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
      Left            =   810
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label lblnombremovimiento 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   90
      TabIndex        =   9
      Top             =   75
      Width           =   11445
   End
End
Attribute VB_Name = "frmentradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_pedimento As String
Dim var_cantidad_multibondeados As Double
Dim var_kanban As String
Dim var_descripcion_etiqueta As String
Dim var_numero_serie As Integer
Dim var_txt_archivo As String
Dim var_clave_almacen_seleccionado As String
Dim var_peso_correcto As Boolean
Dim var_cajas As Boolean
Dim var_codigo_caja As String
Dim var_peso_caja As Double
Dim var_cantidad_caja_peso As Double
Dim var_tolerancia_peso_caja As Double
Dim var_a?o As Integer
Dim var_origen As String
Dim var_lote As Double
Dim var_consecutivo As Integer
Dim var_transporto As String
Dim var_tipo_proveedor As String
Dim var_primera_vez As Boolean
Dim var_numero_folio As Double
Dim VAR_TABLA_NOMBRE_ORIGEN As String
Dim VAR_RUTA_TABLA_ORIGEN As String
Dim VAR_CAMPO_CODIGO_ORIGEN As String
Dim VAR_CAMPO_DESCRIPCION_ORIGEN As String
Dim VAR_CAMPO_COSTO_ORIGEN As String
Dim VAR_CAMPO_CANTIDAD_ORIGEN As String
Dim VAR_CAMPO_CANTIDAD_ENTRADA As String
Dim VAR_TABLA_DESTINO As String
Dim VAR_CAMPO_CODIGO_DESTINO As String
Dim VAR_CAMPO_DESCRIPCION_DESTINO As String
Dim VAR_CAMPO_COSTO_DESTINO As String
Dim VAR_CAMPO_CANTIDAD_DESTINO  As String
Dim VAR_CAMPO_NUMERO  As String
Dim var_cantidad_enviada As Double
Dim var_cantidad_recibida As Double
Dim var_articulo_enviado As String
Dim var_costo_enviado As Double
Dim var_almacen_Destino As String
Dim var_almacen_origen As String
Dim var_proveedor As String
Dim var_estatus_movimiento As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_modifica As Boolean
Dim var_factura As String
Dim var_cantidad_leida As Double
Dim var_tabla As ADODB.Connection
Dim var_ruta As String
Dim var_folio_enviado As Double
Dim var_referencia As String
Dim var_suma_cantidad_enviada As Double
Dim var_suma_cantidad_recibida As Double
Dim var_numero_causa As Integer
Dim ntablas As Integer
Dim var_fecha_movimiento As Date
Dim var_solo_lectura As Boolean
Dim ok As Boolean
Dim var_entrada_calidad As Boolean
Dim var_almacen_costeo As String
Dim var_ventana As Integer
Dim var_tipo_Cambio As Double
Dim var_moneda_local As Integer
Dim var_clave_moneda As String
Dim var_renglon As Double


Private Sub cmd_buscar_Click()
   Me.cmd_tipo_lectura.Visible = False
   var_ventana = 1
   frm_busqueda.Visible = True
   txt_busqueda_folio.SetFocus
End Sub


Private Sub cmd_cancelar_Click()
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_ENTRADAS_I = New TB_ENTRADAS_I
   Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_ENTRADAS_VISTAS_I = New TB_ENTRADAS_VISTAS_I
   If var_numero_folio > 0 Then
      If var_estatus_movimiento = "C" Then
         MsgBox "El Movimiento ya fue cancelado", vbOKOnly, "ATENCION"
      Else
         If var_estatus_movimiento = "I" Then
            If var_fecha_movimiento <> Date Then
               var_posible_accion = False
               frmsupervisor1.Show 1
               If var_posible_accion = True Then
                  si = MsgBox("?Desea cancelar el movimiento?", vbYesNo, "ATENCION")
                  If si = 6 Then
                     si = MsgBox("Confirmar la cancelaci?n del movimiento", vbYesNo, "ATENCION")
                     If si = 6 Then
                        Set TB_ENC_MOV_CANCELACION = New TB_ENC_MOV_CANCELACION
                        var_actualizar = False
                        var_actualizar = TB_ENC_MOV_CANCELACION.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "C", var_global_supervisor_1, var_global_supervisor_2)
                        rs.Open "select * from tb_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                        While Not rs.EOF
                              rsaux4.Open "update tb_archivo_comparacion set floa_com_cantidad_recibida = floa_com_cantidad_recibida - " + CStr(rs!floa_ent_cantidaD) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_com_referencia = '" + txt_archivo + "' and vcha_art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "' and inte_com_consecutivo = " + CStr(rs!inte_Ent_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                              rsaux4.Open "update tb_archivo_comparacion set floa_com_cantidad_recibida = floa_com_cantidad_recibida - " + CStr(rs!floa_ent_cantidaD) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_com_referencia = '" + txt_archivo + "' and vcha_art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "' and inte_com_consecutivo = " + CStr(rs!inte_Ent_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                              rs.MoveNext
                        Wend
                        rs.Close
                        lbl_cancelado = "MOVIMIENTO CANCELADO"
                        Me.cmd_imprimir.Enabled = False
                        Me.cmd_cancelar.Enabled = False
                        MsgBox "El movimiento a sido cancelado", vbOKOnly, "ATENCION"
                        var_estatus_movimiento = "C"
                     End If
                  End If
               End If
            Else
               var_posible_accion = False
               frmsupervisor1.Show
               If var_posible_accion = True Then
                  si = MsgBox("?Desea cancelar el movimiento?", vbYesNo, "ATENCION")
                  If si = 6 Then
                     si = MsgBox("Confirmar la cancelaci?n del movimiento", vbYesNo, "ATENCION")
                     If si = 6 Then
                        Set TB_ENC_MOV_CANCELACION = New TB_ENC_MOV_CANCELACION
                        var_actualizar = False
                        var_actualizar = TB_ENC_MOV_CANCELACION.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "C", var_global_supervisor_1, var_global_supervisor_2)
                        var_estatus_movimiento = "C"
                        MsgBox "El movimiento a sido cancelado", vbOKOnly, "ATENCION"
                     End If
                  End If
               End If
            End If
         Else
            MsgBox "El Movimiento no a sido cerrado aun", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "No se a seleccionado un movimiento", vbOKOnly, "ATENCION"
   End If
End Sub


Private Sub cmd_crossdocking_Click()
   If var_estatus_movimiento = "I" Then
      If Me.txt_folio <> "" Then
         If var_empresa = "31" And (var_clave_movimiento = "EP" Or var_clave_movimiento = "EI") Then
            If rsaux10.State = 1 Then
               rsaux10.Close
            End If
            rsaux10.Open "select * from tb_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' AND vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + Me.txt_folio, cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux10.EOF Then
               rs.Open "select * from tb_TEMPORAL_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' AND vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + Me.txt_folio + " AND FLOA_ENT_CANTIDAD - ISNULL(FLOA_ENT_cANTIDAD_CROSSDOCKING,0) > 0", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  frmsalidas_crossdocking_cantia.lbl_numero_entrada = Me.txt_folio
                  frmsalidas_crossdocking_cantia.lbl_movimiento_entrada = var_clave_movimiento
                  While Not rs.EOF
                        Set list_item = frmsalidas_crossdocking_cantia.lv_articulos.ListItems.Add(, , Trim(rs!VCHA_ART_ARTICULO_ID))
                        rsaux.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_Art_nombre_espa?ol), "", rsaux!vcha_Art_nombre_espa?ol)
                        End If
                        rsaux.Close
                        list_item.SubItems(2) = Format(rs!floa_ent_cantidaD - IIf(IsNull(rs!floa_ent_Cantidad_crossdockinG), 0, rs!floa_ent_Cantidad_crossdockinG), "###,###,##0.00")
                        list_item.SubItems(3) = Format(IIf(IsNull(rs!floa_ent_costo), 0, rs!floa_ent_costo), "###,###,##0.00")
                        list_item.SubItems(4) = Format(IIf(IsNull(rs!floa_ent_precio), 0, rs!floa_ent_precio), "###,###,##0.00")
                        list_item.SubItems(7) = var_clave_movimiento
                        list_item.SubItems(8) = var_numero_folio
                        rs.MoveNext
                  Wend
               End If
               rs.Close
               rsaux10.Close
               frmsalidas_crossdocking_cantia.Show 1
            Else
               MsgBox "El movimiento no existe o no a sido impreso", vbOKOnly, "ATENCION"
               rsaux10.Close
            End If
            
         Else
            rs.Open "select * from tb_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' AND vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + Me.txt_folio, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               frmsalidas_crossdocking.lbl_numero_entrada = Me.txt_folio
               While Not rs.EOF
                     Set list_item = frmsalidas_crossdocking.lv_articulos.ListItems.Add(, , Trim(rs!VCHA_ART_ARTICULO_ID))
                     rsaux.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_Art_nombre_espa?ol), "", rsaux!vcha_Art_nombre_espa?ol)
                     End If
                     rsaux.Close
                     list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_ent_cantidaD), 0, rs!floa_ent_cantidaD), "###,###,##0.00")
                     list_item.SubItems(3) = Format(IIf(IsNull(rs!floa_ent_costo), 0, rs!floa_ent_costo), "###,###,##0.00")
                     list_item.SubItems(4) = Format(IIf(IsNull(rs!floa_ent_precio), 0, rs!floa_ent_precio), "###,###,##0.00")
                     rs.MoveNext
               Wend
               rs.Close
               frmsalidas_crossdocking.Show 1
            Else
               MsgBox "El movimiento no existe o no a sido impreso", vbOKOnly, "ATENCION"
               rs.Close
            End If
         End If
      End If
   Else
      If Me.txt_folio = "" Then
         MsgBox "No se a seleccionado un movimiento", vbOKOnly, "ATENCION"
      Else
         MsgBox "El movimiento no a sido impreso aun", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub cmd_imprimir_Click()
   
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_ENTRADAS_I = New TB_ENTRADAS_I
   Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_ENTRADAS_VISTAS_I = New TB_ENTRADAS_VISTAS_I
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
   
   Dim var_posible_cerrar_movimiento As Integer
   var_posible_cerrar_movimiento = 1
   
   Dim dl As Long                                 ' Valor devuelto por la funci?n API
   Dim sAttributes As String                  ' Aributos
   Dim sDriver As String                       ' Nombre del controlador
   Dim sDescription As String                ' Descripci?n del DSN
   Dim sDsnName As String                  ' Nombre del DSN

   Const ODBC_ADD_SYS_DSN As Long = 4         ' Se crear? un DSN de sistema
   Const vbAPINull As Long = 0&                         ' Puntero NULL

   ' se elimina
   Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminar? un DSN de sistema
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
   
   
   If Me.txt_folio <> "" Then
   If var_posible_cerrar_movimiento = 1 Then
      If var_numero_folio > 0 Then
         'var_estatus_movimiento = ""
         If var_empresa = "31" Then
            'rs.Open "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_ent_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            'While Not rs.EOF
            '      rsaux.Open "select * from TB_EXISTENCIAS_UBICACIONES where vcha_Art_articulo_id = '" + rs!VCHA_aRT_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
            '      If Not rsaux.EOF Then
            '         If rsaux1.State = 1 Then
            '            rsaux1.Close
            '         End If
            '         rsaux1.Open "select *  from tb_ubicaciones_almacen where vcha_alm_almacen_id = '" + var_almacen_Destino + "'  and vcha_Art_articulo_id = '" + rs!VCHA_aRT_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
            '         If rsaux1.EOF Then
            '            rsaux2.Open "insert into tb_ubicaciones_almacen (vcha_alm_almacen_id, vcha_Art_Articulo_id, vcha_ubi_ubicacion_1) values ('" + var_almacen_Destino + "','" + rs!VCHA_aRT_ARTICULO_ID + "','" + rsaux!VCHA_UBI_UBICACION + "')", cnn, adOpenDynamic, adLockOptimistic
            '         End If
            '         rsaux1.Close
            '      End If
            '      rsaux.Close
            '      rs.MoveNext
            'Wend
            'rs.Close
         End If
         If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
            If Trim(var_reporte_imprimir) <> "" Then
               If var_clave_movimiento = "EI" And var_almacen_Destino = "RETEX" Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_reporte_movimientos_entrada_intercompa?ia_segundas.rpt")
                  frmvistasprevias.cr.ReportSource = reporte
                  reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_INTERCOMPA?IAS_SEGUNDAS.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_REPORTE_ENTRADAS_INTERCOMPA?IAS_SEGUNDAS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "'"
                  For ntablas = 1 To reporte.Database.Tables.Count
                     reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Movimientos"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
               Else
                  'MsgBox var_reporte_imprimir
                  If var_clave_movimiento = "EC" Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_entradas_comparacion_compras.rpt")
                  Else
                     Set reporte = appl.OpenReport(App.Path + "\" + Trim(var_reporte_imprimir) + ".rpt")
                  End If
                  frmvistasprevias.cr.ReportSource = reporte
                  If var_empresa = "06" Then
                     'reporte.RecordSelectionFormula = "{VW_ENTRADAS_COMPARACION.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_ENTRADAS_COMPARACION.VCHA_EMO_ALMACEN_DESTINO} = '" + var_almacen_Destino + "' AND {VW_ENTRADAS_COMPARACION.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and {VW_ENTRADAS_COMPARACION.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' and {VW_ENTRADAS_COMPARACION.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {@cantidad_enviada} > 0"
                     reporte.RecordSelectionFormula = "{VW_ENTRADAS_COMPARACION.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_ENTRADAS_COMPARACION.vcha_uor_unidad_id} = '" + var_unidad_organizacional + "' AND {VW_ENTRADAS_COMPARACION.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' and {VW_ENTRADAS_COMPARACION.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "'"
                  Else
                     reporte.RecordSelectionFormula = "{VW_ENTRADAS_COMPARACION.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_ENTRADAS_COMPARACION.vcha_uor_unidad_id} = '" + var_unidad_organizacional + "' AND {VW_ENTRADAS_COMPARACION.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' and {VW_ENTRADAS_COMPARACION.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "'"
                  End If
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Movimientos"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  rsaux4.Open "select * from TB_EXISTENCIAS_SERIES WHERE vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_eMO_numero = " + Str(var_numero_folio)
                  If Not rsaux4.EOF Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_entradas_numero_serie.rpt")
                     frmvistasprevias.cr.ReportSource = reporte
                     reporte.RecordSelectionFormula = "{VW_EXISTENCIAS_SERIES.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and {VW_EXISTENCIAS_SERIES.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' AND {VW_EXISTENCIAS_SERIES.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_Destino + "' and {VW_EXISTENCIAS_SERIES.vcha_mov_movimiento_id} = '" + var_clave_movimiento + "' and {VW_EXISTENCIAS_SERIES.INTE_EMO_NUMERO} = " + CStr(var_numero_folio)
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     frmvistasprevias.cr.ViewReport
                     frmvistasprevias.Caption = "Reporte de Movimientos"
                     frmvistasprevias.Show
                     Set reporte = Nothing
                  End If
                  rsaux4.Close
               End If
               var_m = 0
               If var_empresa = "180000" Then
                  If var_clave_movimiento = "DT" Then
                     rsaux9.Open "select vcha_age_agente_id from tb_clientes where vcha_cli_clave_id = '" + var_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux9.EOF Then
                        var_clave_agente_correo = IIf(IsNull(rsaux9!VCHA_AGE_AGENTE_ID), "", rsaux9!VCHA_AGE_AGENTE_ID)
                     End If
                     rsaux9.Close
                     rsaux9.Open "select * from tb_agentes where vcha_age_agente_id =  '" + var_clave_agente_correo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux9.EOF Then
                        var_correo_electronico = IIf(IsNull(rsaux9!VCHA_AGE_EMAIL), "", rsaux9!VCHA_AGE_EMAIL)
                        var_nombre_agente_correo = IIf(IsNull(rsaux9!VCHA_AGE_NOMBRE), "", rsaux9!VCHA_AGE_NOMBRE)
                     End If
                     rsaux9.Close
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
                        MAPIMessages1.MsgSubject = "Devoluci?n de tienda " + Me.txt_archivo
                        MAPIMessages1.MsgNoteText = "Se anexa informaci?n de la devoluci?n  " + Me.txt_archivo
                        var_Archivo = App.Path & "\Devolucion_" + Trim(Me.txt_archivo) + ".txt"
                        Open (App.Path & "\Devolucion_" + Trim(Me.txt_archivo) + ".txt") For Output As #1
                        Print #1, "Se genero la devoluci?n " + Trim(Me.txt_archivo) + " con los siguientes datos"
                        Print #1, ""
                        Print #1, "Cliente: " + Trim(var_nombre_agente_correo)
                        Print #1, ""
                        rsaux9.Open "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_Alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                        Print #1, "CODIGO       DESCRIPCION                                      CANTIDAD"
                        Print #1, "======================================================================"
                        var_total_correo = 0
                        While Not rsaux9.EOF
                              var_linea = ""
                              var_cantidad_correo = Format(IIf(IsNull(rsaux9!floa_ent_cantidaD), 0, rsaux9!floa_ent_cantidaD), "###,###,##0.00")
                              var_total_correo = var_total_correo + IIf(IsNull(rsaux9!floa_ent_cantidaD), 0, rsaux9!floa_ent_cantidaD)
                              rsaux8.Open "select vcha_Art_nombre_espa?ol from tb_articulos where vcha_Art_articulo_id = '" + rsaux9!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux8.EOF Then
                                 var_nombre_articulo_correo = Mid(IIf(IsNull(rsaux8!vcha_Art_nombre_espa?ol), "", rsaux8!vcha_Art_nombre_espa?ol), 1, 40)
                              End If
                              rsaux8.Close
                              For var_j = Len(var_cantidad_correo) To 14
                                  var_cantidad_correo = " " + var_cantidad_correo
                              Next var_j
                              For var_j = Len(var_nombre_articulo_correo) To 40
                                  var_nombre_articulo_correo = var_nombre_articulo_correo + " "
                              Next var_j
                              var_linea = rsaux9!VCHA_ART_ARTICULO_ID + " " + var_nombre_articulo_correo + " " + var_cantidad_correo
                              Print #1, var_linea
                              rsaux9.MoveNext
                        Wend
                        Print #1, "======================================================================"
                        rsaux9.Close
                        var_total_correo_str = Format(var_total_correo, "###,###,##0.00")
                        For var_j = Len(var_total_correo_str) To 14
                            var_total_correo_str = " " + var_total_correo_str
                        Next var_j
                        var_linea = "                                       POR UN TOTAL DE:" + var_total_correo_str
                        Print #1, var_linea
                        Print #1, ""
                        Close #1
                        MAPIMessages1.AttachmentPathName = var_Archivo
                        MAPIMessages1.Send True
                        If MAPISession1.SessionID > 0 Then
                           MAPISession1.SignOff
                        End If
                     Else
                     End If
                  End If
               End If
               '''' fin del correo
            Else
               MsgBox "El movimiento no tiene un reporte asociado", vbOKOnly, "ATENCION"
            End If
         Else
             var_si = MsgBox("?Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
            If var_si = 1 Then
               'If var_empresa = "31" And var_clave_movimiento = "EC" Then
               '   rsaux9.Open "select * from tb_Temporal_Entradas where vcha_Emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
               '   While Not rsaux9.EOF
               '         rsaux10.Open "update tb_articulos set inte_art_detenido = 1 where vcha_Art_articulo_id = '" + rsaux9!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
               '         rsaux9.MoveNext
               '   Wend
               '   rsaux9.Close
               'End If
               'On Error GoTo salir:
               var_posible_cerrar_movimiento = 1
               x = 1 ' aqui se cambia lo del apartado
               If x = 1 Then
                  If var_empresa = "31" And var_clave_movimiento = "EI" Then
                     If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
                        var_posible_cerrar_movimiento = 1
                     Else
                        var_consecutivo = 1
                        rsaux9.Open "select * from tb_Temporal_Entradas where vcha_Emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux9.EOF
                              rs.Open "select * from TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA where  vcha_Emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(var_numero_folio) + " and vcha_Art_articulo_id = '" + rsaux9!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                              If rs.EOF Then
                                 var_cadena = "insert into TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA (inte_tem_consecutivo, vcha_tem_almacen_apartado, vcha_tem_almacen_apartado_entregar, vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, floa_ent_cantidad_almacen_apartado, floa_ent_cantidad_almacen_apartado_entregar) "
                                 var_cadena = var_cadena + " values (" + CStr(var_consecutivo) + ", 'ALAPP', 'ALAPPE','" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_Destino + "','" + var_clave_movimiento + "','" + CStr(var_numero_folio) + "','" + rsaux9!VCHA_ART_ARTICULO_ID + "', " + CStr(rsaux9!floa_ent_cantidaD) + ", " + CStr(rsaux9!floa_ent_costo) + "," + CStr(rsaux9!floa_ent_precio) + ",0,0)"
                                 rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 rsaux.Open "update  TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA set floa_Ent_Cantidad =  " + CStr(rsaux9!floa_ent_cantidaD) + " where vcha_Emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(var_numero_folio) + " and vcha_Art_articulo_id = '" + rsaux9!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rs.Close
                              rsaux9.MoveNext
                        Wend
                        rsaux9.Close
                        rs.Open "SELECT * FROM TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA WHERE  vcha_Emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                        While Not rs.EOF
                              rsaux.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ART_aRTICULO_ID = '" + IIf(IsNull(rs!VCHA_ART_ARTICULO_ID), "", rs!VCHA_ART_ARTICULO_ID) + "' AND VCHA_ALM_ALMACEN_ID = 'ALAPP'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux.EOF Then
                                 var_cantidad_apartada = IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad)
                              Else
                                 var_cantidad_apartada = 0
                              End If
                              rsaux.Close
                              rsaux.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ART_aRTICULO_ID = '" + IIf(IsNull(rs!VCHA_ART_ARTICULO_ID), "", rs!VCHA_ART_ARTICULO_ID) + "' AND VCHA_ALM_ALMACEN_ID = 'ALAPPE'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux.EOF Then
                                 var_cantidad_apartada_entregar = IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad)
                              Else
                                 var_cantidad_apartada_entregar = 0
                              End If
                              rsaux.Close
                              If var_cantidad_apartada_entregar > 0 Then
                                 var_cantidad_apartada_entregar = var_cantidad_apartada_entregar * (0 - 1)
                              End If
                              rsaux.Open "update TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA set floa_ent_cantidad_almacen_apartado = " + CStr(var_cantidad_apartada) + ", floa_ent_cantidad_almacen_apartado_entregar = " + CStr(var_cantidad_apartada_entregar) + " where vcha_Emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(var_numero_folio) + " and vcha_Art_Articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                              rs.MoveNext
                        Wend
                        rs.Close
                        rs.Open "select * from TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA where  vcha_Emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(var_numero_folio) + " and (floa_ent_cantidad_almacen_apartado < 0 or floa_ent_cantidad_almacen_apartado_entregar < 0)", cnn, adOpenDynamic, adLockOptimistic
                        'rs.Open "select * from TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA where  vcha_Emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(var_numero_folio) + " and (floa_ent_cantidad_almacen_apartado < 0)", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_consecutivo_apartados_Cantia = var_numero_folio
                           var_clave_movimiento_apartados_Cantia = var_clave_movimiento
                           frmsalidas_apartados_cantia.Show 1
                           rsaux10.Open "select * from TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA where  vcha_Emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(var_numero_folio) + " and vcha_ent_marca_traspaso = 'T'", cnn, adOpenDynamic, adLockOptimistic
                           If rsaux10.EOF Then
                              If var_almacen_Destino = "CC_1" Then
                                 var_posible_cerrar_movimiento = 1
                              Else
                                 var_posible_cerrar_movimiento = 0
                              End If
                           Else
                              var_posible_cerrar_movimiento = 1
                           End If
                           rsaux10.Close
                        Else
                           var_posible_cerrar_movimiento = 1
                        End If
                        rs.Close
                     End If
                  Else
                     var_posible_cerrar_movimiento = 1
                  End If
               End If 'fin del x
               If var_posible_cerrar_movimiento = 1 Then
                  x = 0
                  Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
                  rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  VAR_ZZ = rs.RecordCount
                  If var_clave_movimiento = "EC" Then
                     var_posible_entrada = True
                     If x = 0 Then
                        var_si_paso = Obtener_Identificador_Recepcion(cnnoracle, 0)
                     End If
                     
                    
                     If var_clave_movimiento = "EC" Then
                        While Not rs.EOF
                              If x = 0 Then
                                 If IsNumeric(txt_archivo) Then
                                    VAR_PASO = Insertar_Recepcion(CDbl(var_txt_archivo), CDbl(rs!p_rc_numero_linea), CDbl(rs!floa_ent_cantidaD), CDbl(rs!P_RC_LINEA_ID), CDbl(var_numero_folio), CDbl(var_unidad_OC), "0", txt_factura)
                                 End If
                              End If
                              rs.MoveNext
                        Wend
                     End If
                     'MsgBox var_clave_movimiento
                  Else
                     var_posible_entrada = True
                  End If
                  If rs.RecordCount > 0 Then
                    rs.MoveFirst
                  End If
       'GoTo salir:r
                  If var_posible_entrada = True Then
                     var_posible = 1
                        'GoTo salir:
                     cnn.BeginTrans
                     If Not rs.EOF Then
                        var_inserta = False
                        x = 1
                        If x = 1 Then
                        'If var_posible_kanban = 1 Then
                           If var_clave_movimiento = "EP" And var_empresa = "02" Then
                              var_cadena = "SELECT dbo.TB_TEMPORAL_ENTRADAS.VCHA_EMP_EMPRESA_ID, dbo.TB_TEMPORAL_ENTRADAS.VCHA_UOR_UNIDAD_ID, dbo.TB_TEMPORAL_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_TEMPORAL_ENTRADAS.INTE_ENT_NUMERO, dbo.TB_TEMPORAL_ENTRADAS.VCHA_ART_ARTICULO_ID , dbo.TB_Articulos.INTE_ART_KANBAN FROM dbo.TB_TEMPORAL_ENTRADAS INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_TEMPORAL_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_ARTICULOS.INTE_ART_KANBAN = 1) AND (dbo.TB_TEMPORAL_ENTRADAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND  (dbo.TB_TEMPORAL_ENTRADAS.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') AND (dbo.TB_TEMPORAL_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "') AND  (dbo.TB_TEMPORAL_ENTRADAS.INTE_ENT_NUMERO = " + CStr(var_numero_folio) + ")"
                              rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux10.EOF Then
                                 Set TB_PROC_KANBANS_EN_MOV_ENTRADA = New TB_PROC_KANBANS_EN_MOV_ENTRADA
                                 var_inserta = TB_PROC_KANBANS_EN_MOV_ENTRADA.Anadir(var_almacen_Destino, var_clave_movimiento, CDbl(Me.txt_folio), "", "")
                                 If var_kanban_exito = "N" Then
                                    var_posible = 0
                                 End If
                              Else
                                 var_posible = 1
                              End If
                           Else
                              var_posible = 1
                           End If
                        'End If
                        'cnn.RollbackTrans
                        End If
                        If var_posible = 1 Then
                           If var_empresa = "18" And var_clave_movimiento = "EI" And var_almacen_Destino = "RETEX" Then
                              rsaux10.Open "select * from tb_temporal_entradas where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'EI' and inte_ent_numero = " + Me.txt_folio, cnn, adOpenDynamic, adLockOptimistic
                              While Not rsaux10.EOF
                                    rsaux11.Open "update tb_temporal_entradas set floa_ent_costo =  (floa_ent_costo * floa_sal_Cantidad_metros)/ floa_ent_cantidad where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'EI' and inte_ent_numero = " + Me.txt_folio + " and vcha_Art_articulo_id = '" + rsaux10!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                    rsaux10.MoveNext
                              Wend
                              rsaux10.Close
                           End If
                           var_inserta = TB_ENTRADAS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, var_tipo_documento, var_almacen_origen, var_folio_enviado)
                        End If
                     End If
                     rs.Close
                     If var_tipo_documento = "V" Then
                        var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, "", Now, var_tipo_Cambio)
                     Else
                        var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, var_tipo_Cambio)
                     End If
                     var_estatus_movimiento = "I"
                     If var_tipo_documento = "V" Then
                        var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, "I", Now, var_tipo_Cambio)
                     Else
                        var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, var_tipo_Cambio)
                     End If
                     'cnn.RollbackTrans
                     '################################################################
                     cnn.CommitTrans
                     'If var_clave_movimiento = "EI" Or var_clave_movimiento = "ETA" Or var_clave_movimiento = "EP" Then
                     ZZ = 0
                     If ZZ = 0 Then
                     If var_clave_movimiento = "EI" Or var_clave_movimiento = "ETA" Then
                        If var_clave_movimiento = "EP" And var_empresa = "16" Then
                        Else
                           If var_almacen_Destino = "RETEX" Then
                              var_cadena = "SELECT dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ENTRADAS.INTE_ENT_NUMERO, Sum(dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD) AS CANTIDAD, SUM(dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_CANTIDAD_ENVIADA * dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_COSTO / dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD * dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD) AS COSTO, SUM(dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_CANTIDAD_ENVIADA * dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_PRECIO / dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD * dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD) AS Precio FROM  dbo.TB_ENTRADAS INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND "
                              var_cadena = var_cadena + " dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENTRADAS.INTE_ENT_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN dbo.TB_ARCHIVO_COMPARACION ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA = dbo.TB_ARCHIVO_COMPARACION.VCHA_COM_REFERENCIA AND dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') "
                              var_cadena = var_cadena + " AND (dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') AND"
                              var_cadena = var_cadena + " (dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = 'EI') AND (dbo.TB_ENTRADAS.INTE_ENT_NUMERO = " + CStr(var_numero_folio) + ") GROUP BY dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ENTRADAS.INTE_ENT_NUMERO "
                           Else
                              If var_empresa = "31" Then
                                 var_cadena = " SELECT dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ENTRADAS.INTE_ENT_NUMERO, SUM(dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD) AS CANTIDAD, SUM(dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD * dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_COSTO) AS COSTO, SUM(dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD * dbo.TB_ENTRADAS.FLOA_ENT_PRECIO) As Precio FROM dbo.TB_ENTRADAS INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENTRADAS.INTE_ENT_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN "
                                 var_cadena = var_cadena + " dbo.TB_ARCHIVO_COMPARACION ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA = dbo.TB_ARCHIVO_COMPARACION.VCHA_COM_REFERENCIA AND dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') AND (dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = 'EI') AND (dbo.TB_ENTRADAS.INTE_ENT_NUMERO = " + CStr(var_numero_folio) + ") "
                                 var_cadena = var_cadena + " GROUP BY dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ENTRADAS.INTE_ENT_NUMERO                "
                              Else
                                 'var_cadena = "SELECT  dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ENTRADAS.INTE_ENT_NUMERO, SUM(dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD) AS CANTIDAD, SUM(dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD * dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_COSTO) AS COSTO, SUM(dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD * dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_PRECIO) As Precio FROM dbo.TB_ENTRADAS INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND Dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENTRADAS.INTE_ENT_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN "
                                 'var_cadena = var_cadena + " dbo.TB_ARCHIVO_COMPARACION ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ALM_ALMACEN_ID AND  dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA = dbo.TB_ARCHIVO_COMPARACION.VCHA_COM_REFERENCIA AND dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') AND"
                                 'var_cadena = var_cadena + " (dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "') AND (dbo.TB_ENTRADAS.INTE_ENT_NUMERO = " + CStr(var_numero_folio) + ") GROUP BY dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ENTRADAS.INTE_ENT_NUMERO"
                                 var_cadena = "SELECT     dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ENTRADAS.INTE_ENT_NUMERO, SUM(dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD) AS CANTIDAD, SUM(dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD * dbo.TB_ENTRADAS.FLOA_ENT_PRECIO) AS PRECIO, SUM(dbo.TB_ENTRADAS.FLOA_ENT_COSTO * dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD) As Costo FROM dbo.TB_ENTRADAS INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENTRADAS.INTE_ENT_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO"
                                 var_cadena = var_cadena + " WHERE     (dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') AND (dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "') AND (dbo.TB_ENTRADAS.INTE_ENT_NUMERO = " + CStr(var_numero_folio) + ") GROUP BY dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ENTRADAS.INTE_ENT_NUMERO "
                              End If
                           End If
                           'MsgBox var_cadena
                           If rsaux10.State = 1 Then
                              rsaux10.Close
                           End If
                           rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                           'MsgBox cnnoracle.ConnectionString
                           'MsgBox "select * from tb_generador_polizas where empresa_id = '" + var_empresa + "'"
                           If var_clave_movimiento = "ETA" Then
                              concepto_1 = "ENTRADA TRASPASO " + CStr(var_numero_folio) + " "
                              CONCEPTO_2 = "ORIGEN " + Me.txt_origen + " " + Me.txt_archivo
                              CONCEPTO_3 = "RECEPCION POR TRASPASO " + Me.txt_destino
                              
                              If var_empresa = "06" Or var_empresa = "17" Or var_empresa = "02" Or var_empresa = "15" Or var_empresa = "16" Then
                                
                                 If var_almacen_Destino = "ABPT" Then
                                    rsaux11.Open "select * from tb_generador_polizas where poliza_id = '59'", cnnoracle, adOpenDynamic, adLockOptimistic
                                 End If
                                 If var_almacen_Destino = "28" Then
                                    rsaux11.Open "select * from tb_generador_polizas where poliza_id = '89'", cnnoracle, adOpenDynamic, adLockOptimistic
                                 End If
                                 If var_almacen_Destino = "Q0Z" Then
                                    rsaux11.Open "select * from tb_generador_polizas where poliza_id = '19'", cnnoracle, adOpenDynamic, adLockOptimistic
                                 End If
                                 If var_almacen_Destino = "MPCOL" Then
                                    rsaux11.Open "select * from tb_generador_polizas where poliza_id = '24'", cnnoracle, adOpenDynamic, adLockOptimistic
                                 End If
                                 If var_almacen_Destino = "MPCOC" Then
                                    rsaux11.Open "select * from tb_generador_polizas where poliza_id = '28'", cnnoracle, adOpenDynamic, adLockOptimistic
                                 End If
                                 If var_almacen_Destino = "MPEDR" Then
                                    rsaux11.Open "select * from tb_generador_polizas where poliza_id = '23'", cnnoracle, adOpenDynamic, adLockOptimistic
                                 End If
                                 If var_almacen_Destino = "PTMU" Or var_almacen_Destino = "CMU" Or var_almacen_Destino = "PMU" Then
                                    rsaux11.Open "select * from tb_generador_polizas where poliza_id = '60'", cnnoracle, adOpenDynamic, adLockOptimistic
                                 End If
                                 If var_almacen_Destino = "BORPT" Or var_almacen_Destino = "BORSEG" Then
                                    rsaux11.Open "select * from tb_generador_polizas where poliza_id = '66'", cnnoracle, adOpenDynamic, adLockOptimistic
                                 End If
                                 If var_almacen_Destino = "8" Then
                                    rsaux11.Open "select * from tb_generador_polizas where poliza_id = '59'", cnnoracle, adOpenDynamic, adLockOptimistic
                                 End If
                                 If var_almacen_Destino = "CAEE" Then
                                    rsaux11.Open "select * from tb_generador_polizas where poliza_id = '59'", cnnoracle, adOpenDynamic, adLockOptimistic
                                 End If
                              End If
                              If var_empresa = "18" Then
                                 rsaux11.Open "select * from tb_generador_polizas where empresa_id = '" + var_empresa + "' and poliza_id = '17'", cnnoracle, adOpenDynamic, adLockOptimistic
                              End If
                           End If
                           If var_clave_movimiento = "EP" Then
                               ok = pro_DatosTransito(CStr(var_numero_folio), var_almacen_Destino, "")
                           End If
                           If var_clave_movimiento = "EI" Then
                              cnnoracle.Close
                              cnnoracle.Open "Provider=OraOLEDB.Oracle.1;User ID=INTERFACE;Data Source=ap;Extended Properties=;Persist Security Info=True;Password=INTERFACE"
                              cnnoracle.CursorLocation = adUseClient
                            '################
                            'REVISAR CON FERNANDO
                            '################
                              rsaux11.Open "select * from tb_generador_polizas where empresa_id = '" + var_empresa + "' AND POLIZA_ID = 2", cnnoracle, adOpenDynamic, adLockOptimistic
                              concepto_1 = "EI:" + Me.txt_folio.Text + ";" + "Fac:" + Me.txt_archivo.Text
                              CONCEPTO_2 = "FACTURA NUM: " + Me.txt_archivo
                              CONCEPTO_3 = "POLIZA FACTURAS INTERCOMPA?IA"
                           End If
                           If var_clave_movimiento = "EP" Then
                              'If var_empresa = "18" Then
                              '   rsaux11.Open "select * from tb_generador_polizas where poliza_id = '61'", cnnoracle, adOpenDynamic, adLockOptimistic
                              'End If
                              If var_empresa = "02" Or var_empresa = "18" Or var_empresa = "06" Or var_empresa = "17" Then
                                 rsaux11.Open "select * from tb_generador_polizas where poliza_id = '62'", cnnoracle, adOpenDynamic, adLockOptimistic
                              End If
                              concepto_1 = "ENTRADA DE PRODUCCION " + Me.txt_folio
                              CONCEPTO_2 = "NOTA ENVIO NUM: " + Me.txt_archivo + " " + Me.txt_origen
                              CONCEPTO_3 = "POLIZA ENTRADA DE PRODUCCION"
                           End If
                           
                           While Not rsaux11.EOF
                                 var_tipo_poliza = rsaux11!tipo
                                 var_origen_poliza = rsaux11!Origen
                                 var_categoria_poliza = rsaux11!categoria
                                 var_moneda_poliza = rsaux11!moneda
                                 var_segmento1_poliza = rsaux11!segmento1
                                 var_segmento2_poliza = rsaux11!segmento2
                                 var_segmento3_poliza = rsaux11!segmento3
                                 var_segmento4_poliza = rsaux11!SEGMENTO4
                                 var_segmento5_poliza = rsaux11!segmento5
                                 var_segmento6_poliza = rsaux11!segmento6
                                 var_segmento7_poliza = rsaux11!segmento7
                                 var_juego_libros_poliza = rsaux11!juego_libros
                                 var_descripcion_poliza = rsaux11!descripcion
                                 var_cargo_poliza = rsaux11!cargo
                                 var_abono_poliza = rsaux11!abono
                                 var_precio = rsaux11!Precio
                                 If var_precio = 1 Then
                                    If rsaux10.EOF Then
                                       var_importe_precio = 0
                                    Else
                                       var_importe_precio = rsaux10!Precio
                                    End If
                                 Else
                                    If rsaux10.EOF Then
                                       var_importe_precio = 0
                                    Else
                                       var_importe_precio = IIf(IsNull(rsaux10!Costo), 0, rsaux10!Costo)
                                       'MsgBox CStr(var_importe_precio)
                                    End If
                                 End If
                                 var_cadena = "InsERT INTO IN_TB_POLIZAS_INT (STATUS, SET_OF_BOOKS_ID, USER_JE_SOURCE_NAME, USER_JE_CATEGORY_NAME, ACCOUNTING_DATE, CURRENCY_CODE, DATE_CREATED, ACTUAL_FLAG,  SEGMENT1, SEGMENT2, SEGMENT3, SEGMENT4, SEGMENT5, SEGMENT6, SEGMENT7, ENTERED_DR, ENTERED_CR, ACCOUNTED_DR, ACCOUNTED_CR, GROUP_ID, REFERENCE4, REFERENCE5, REFERENCe10, REFERENCE1, REFERENCE2, CREATED_BY)"
                                 If var_cargo_poliza = 1 Then
                                    var_cadena = var_cadena & " VALUES ('NEW', " & CStr(var_juego_libros_poliza) & ",'" & var_origen_poliza & "','" & var_categoria_poliza & "',TO_DATE('" & CStr(Date) & "','DD/MM/YYYY'),'" & var_moneda_poliza & "',TO_DATE('" & CStr(Date) & "','DD/MM/YYYY'),'A','" & var_segmento1_poliza & "','" & var_segmento2_poliza & "','" & var_segmento3_poliza & "','" & var_segmento4_poliza & "','" & var_segmento5_poliza & "','" & var_segmento6_poliza & "','" & var_segmento7_poliza & "'," & CStr(IIf(IsNull(var_importe_precio), 0, var_importe_precio)) & ",0," & CStr(IIf(IsNull(var_importe_precio), 0, var_importe_precio)) & ",0,1,'" & concepto_1 & "','" & CONCEPTO_2 & "','" & var_descripcion_poliza & "','" & CONCEPTO_3 & "','" & CONCEPTO_3 & "',1143)"
                                 Else
                                    var_cadena = var_cadena & " VALUES ('NEW', " & CStr(var_juego_libros_poliza) & ",'" & var_origen_poliza & "','" & var_categoria_poliza & "',TO_DATE('" & CStr(Date) & "','DD/MM/YYYY'),'" & var_moneda_poliza & "',TO_DATE('" & CStr(Date) & "','DD/MM/YYYY'),'A','" & var_segmento1_poliza & "','" & var_segmento2_poliza & "','" & var_segmento3_poliza & "','" & var_segmento4_poliza & "','" & var_segmento5_poliza & "','" & var_segmento6_poliza & "','" & var_segmento7_poliza & "',0," & CStr(IIf(IsNull(var_importe_precio), 0, var_importe_precio)) & ",0," & CStr(IIf(IsNull(var_importe_precio), 0, var_importe_precio)) & ",1,'" & concepto_1 & "','" & CONCEPTO_2 & "','" & var_descripcion_poliza & "','" & CONCEPTO_3 & "','" & CONCEPTO_3 & "',1143)"
                                 End If
                                 'MsgBox var_cadena
                                 rsaux9.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                                 rsaux11.MoveNext
                           Wend
                           rsaux11.Close
                           If var_clave_movimiento = "EI" Then
                               rsaux11.Open "select sq_id_facturas.nextval from dual", cnnoracle, adOpenDynamic, adLockOptimistic
                               
                               var_consecutivo = rsaux11(0).Value
                               rsaux11.Close
                               
                               rsaux11.Open "SELECT TOP 1 * FROM tb_Encabezado_movimientos WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND VCHA_emo_REFERENCIA = '" + Me.txt_archivo + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                               var_proveedor_oracle = rsaux11!VCHA_PRO_PROVEEDOR_ID
                               rsaux11.Close
                               'MsgBox var_proveedor_oracle
                               'MsgBox "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_unidad_id = '" + var_proveedor_oracle + "'"
                               rsaux11.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_unidad_id = '" + var_proveedor_oracle + "'", cnn, adOpenDynamic, adLockOptimistic
                               var_empresa_emite = rsaux11!VCHA_EMP_EMPRESA_ID
                               var_proveedor_oracle_2 = rsaux11!vcha_uor_proveedor_oracle
                               rsaux11.Close
                               'MsgBox "select * from tb_empresas_cruzadas_oracle where vcha_emp_Empresa_emite = '" + var_empresa_emite + "' and vcha_emp_empresa_recibe = '" + var_empresa + "'"
                               rsaux11.Open "select * from tb_empresas_cruzadas_oracle where vcha_emp_Empresa_emite = '" + var_empresa_emite + "' and vcha_emp_empresa_recibe = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                               'MsgBox var_empresa_emite
                               var_unidad_oracle = rsaux11!vcha_emp_organizacion
                               rsaux11.Close
                               
                               'MsgBox var_proveedor_oracle
                               'MsgBox "SELECT vendor_site_id FROM po_vendor_sites_all@perpvia.vianney.com.mx Where vendor_id = '" + var_proveedor_oracle_2 + "'  AND vendor_site_id in (4070,2803,1125,1200,1202,1126,1519,1327,1520,3545,1127,1674,4383,2668,1529,1326,2669,2925,3755,1268,1324,3768,2392,9016,3332,1462, 1473,1737,1992, 1991,4618,5454,6407,4380,10626) AND ORG_ID = '" + var_unidad_oracle + "'"
                               rsaux11.Open "SELECT vendor_site_id FROM po_vendor_sites_all@perpvia.vianney.com.mx Where vendor_id = '" + var_proveedor_oracle_2 + "'  AND vendor_site_id in (4070,2803,1125,1200,1202,1126,1519,1327,1520,3545,1127,7435,4383,2668,1529,1326,2669,2925,3755,1268,1324,3768,2392,9016,3332,1462, 1473,1737,1992, 1991,4618,5454,6407,4380,10626,9899,7244,7367) AND ORG_ID = '" + var_unidad_oracle + "'", cnnoracle, adOpenDynamic, adLockOptimistic
                               'MsgBox var_proveedor_oracle_2
                               'MsgBox var_unidad_oracle
                               var_clave_proveedor_oracle = rsaux11!vendor_site_id
                               rsaux11.Close
                            
                               'rsaux11.Open "select sum(FLOA_ENT_CANTIDAD * floa_ENT_costo) FROM  tb_temporal_entradas where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + Me.txt_folio, cnn, adOpenDynamic, adLockOptimistic
                               var_cadena = "SELECT     SUM(dbo.TB_TEMPORAL_ENTRADAS.FLOA_ENT_CANTIDAD * ((dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_COSTO*FLOA_COM_CANTIDAD_ENVIADA + dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_PRECIO *FLOA_COM_CANTIDAD_ENVIADA) /FLOA_ENT_CANTIDAD )) AS Expr1 "
                               var_cadena = var_cadena + " FROM         dbo.TB_TEMPORAL_ENTRADAS INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_TEMPORAL_ENTRADAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_TEMPORAL_ENTRADAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_TEMPORAL_ENTRADAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_TEMPORAL_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND Dbo.TB_TEMPORAL_ENTRADAS.INTE_ENT_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN dbo.TB_ARCHIVO_COMPARACION ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ALM_ALMACEN_ID AND"
                               var_cadena = var_cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA = dbo.TB_ARCHIVO_COMPARACION.VCHA_COM_REFERENCIA AND dbo.TB_TEMPORAL_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ART_ARTICULO_ID WHERE  (dbo.TB_TEMPORAL_ENTRADAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_TEMPORAL_ENTRADAS.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') AND (dbo.TB_TEMPORAL_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "') AND (dbo.TB_TEMPORAL_ENTRADAS.INTE_ENT_NUMERO = " + Me.txt_folio + ") and FLOA_ENT_CANTIDAD >0"
                               rsaux11.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                               var_importe_total = rsaux11(0).Value
                               rsaux11.Close
                               var_importe_total = var_importe_total * 1.16
                            
                               var_cadena = "insert into IN_TB_FACTURAS_INT (INVOICE_ID,INVOICE_NUM,INVOICE_TYPE_LOOKUP_CODE,VENDOR_ID,VENDOR_SITE_ID,INVOICE_AMOUNT,INVOICE_CURRENCY_CODE,EXCHANGE_RATE_TYPE,EXCHANGE_DATE,EXCHANGE_RATE,Description,Source,GL_DATE,INVOICE_DATE,ORG_ID) values (" + CStr(var_consecutivo) + ",'" + Me.txt_archivo + "','STANDARD'," + CStr(var_proveedor_oracle_2) + "," + CStr(var_clave_proveedor_oracle) + "," + CStr(IIf(IsNull(var_importe_total), 0, var_importe_total)) + ",'MXP',null,null,null,'FACTURA DE RECEPCION NUM: " + Me.txt_archivo + "','FACTURA INTERCOMPA?IAS',TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),TO_DATE('" + CStr(Date) + "','DD/MM/YYYY')," + var_unidad_oracle + ")"
                               'MsgBox var_cadena
                               rsaux11.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                            
                               rsaux11.Open "select sq_id_lineas_factura.nextval from dual", cnnoracle, adOpenDynamic, adLockOptimistic
                               var_consecutivo_linea = rsaux11(0).Value
                               rsaux11.Close
                               var_subimporte = var_importe_total / 1.16
                               var_importe_iva = var_importe_total - var_subimporte
                               rsaux11.Open "select amount_includes_tax_flag, vat_code from po_vendor_sites_all@perpvia.vianney.com.mx Where vendor_id = " + CStr(var_proveedor_oracle_2) + " and vendor_site_id = " + CStr(var_clave_proveedor_oracle) + " and org_id = " + CStr(var_unidad_oracle), cnnoracle, adOpenDynamic, adLockOptimistic
                               amount_includes_tax_flag = rsaux11!amount_includes_tax_flag
                               TAX_CODE = IIf(IsNull(rsaux11!vat_code), 0, rsaux11!vat_code)
                               rsaux11.Close
                               rsaux.Open "select awt_group_id from po_vendors@perpvia.vianney.com.mx Where vendor_id = " + CStr(var_proveedor_oracle), cnnoracle, adOpenDynamic, adLockOptimistic
                               AWT_GROUP_ID = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                               rsaux.Close
                               'MsgBox CStr(AWT_GROUP_ID)
                               If TAX_CODE = 0 Then
                                  If AWT_GROUP_ID = 0 Then
                                     var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",1,'ITEM'," + CStr(var_subimporte) + ", TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                                  Else
                                     var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",1,'ITEM'," + CStr(var_subimporte) + ", TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "',NULL,NULL," + CStr(AWT_GROUP_ID) + ",NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                                  End If
                               Else
                                  If AWT_GROUP_ID = 0 Then
                                     var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",1,'ITEM'," + CStr(IIf(IsNull(var_subimporte), 0, var_subimporte)) + ", TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "','" + CStr(TAX_CODE) + "',NULL,NULL,NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                                  Else
                                     var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",1,'ITEM'," + CStr(var_subimporte) + ", TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "','" + CStr(TAX_CODE) + "',NULL," + CStr(AWT_GROUP_ID) + ",NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                                  End If
                               End If
                            
                               'MsgBox var_cadena
                               rsaux11.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                               rsaux11.Open "select sq_id_lineas_factura.nextval from dual", cnnoracle, adOpenDynamic, adLockOptimistic
                               var_consecutivo_linea = rsaux11(0).Value
                               rsaux11.Close
                            
                               var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description,AMOUNT_INCLUDES_TAX_FLAG,TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",2,'TAX'," + CStr(IIf(IsNull(var_importe_iva), 0, var_importe_iva)) + ", TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'IMPUESTO',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                               rsaux11.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                               
                               If var_empresa = "18" Then
                               End If
                               'rsaux10.MoveNext
                               rsaux10.Close
                           End If
                        End If
                    
                     End If
                     End If
                     rsaux2.Open "select * from tb_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                     VAR_ZZ = CStr(rsaux2.RecordCount)
                     'MsgBox (var_zz)
                     If Not rsaux2.EOF Then
                        If var_posible = 1 Then
                           If Trim(var_reporte_imprimir) <> "" Then
                              If var_clave_movimiento = "EI" And var_almacen_Destino = "RETEX" Then
                                 Set reporte = appl.OpenReport(App.Path + "\rep_reporte_movimientos_entrada_intercompa?ia_segundas.rpt")
                                 frmvistasprevias.cr.ReportSource = reporte
                                 reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_INTERCOMPA?IAS_SEGUNDAS.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_REPORTE_ENTRADAS_INTERCOMPA?IAS_SEGUNDAS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "'"
                                 For ntablas = 1 To reporte.Database.Tables.Count
                                    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                 Next ntablas
                                 frmvistasprevias.cr.ViewReport
                                 frmvistasprevias.Caption = "Reporte de Movimientos"
                                 frmvistasprevias.Show 1
                                 Set reporte = Nothing
                              Else
                                 If var_clave_movimiento = "EC" Then
                                    Set reporte = appl.OpenReport(App.Path + "\rep_entradas_comparacion_compras.rpt")
                                 Else
                                    Set reporte = appl.OpenReport(App.Path + "\" + Trim(var_reporte_imprimir) + ".rpt")
                                 End If
                                 If var_empresa = "06" Then
                                    'reporte.RecordSelectionFormula = "{VW_ENTRADAS_COMPARACION.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_ENTRADAS_COMPARACION.VCHA_EMO_ALMACEN_DESTINO} = '" + var_almacen_Destino + "' AND {VW_ENTRADAS_COMPARACION.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and {VW_ENTRADAS_COMPARACION.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' and {VW_ENTRADAS_COMPARACION.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {@cantidad_enviada} > 0"
                                    reporte.RecordSelectionFormula = "{VW_ENTRADAS_COMPARACION.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_ENTRADAS_COMPARACION.VCHA_EMO_ALMACEN_DESTINO} = '" + var_almacen_Destino + "' AND {VW_ENTRADAS_COMPARACION.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and {VW_ENTRADAS_COMPARACION.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' and {VW_ENTRADAS_COMPARACION.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "'"
                                 Else
                                    reporte.RecordSelectionFormula = "{VW_ENTRADAS_COMPARACION.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_ENTRADAS_COMPARACION.VCHA_EMO_ALMACEN_DESTINO} = '" + var_almacen_Destino + "' AND {VW_ENTRADAS_COMPARACION.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and {VW_ENTRADAS_COMPARACION.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' and {VW_ENTRADAS_COMPARACION.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "'"
                                 End If
                                 frmvistasprevias.cr.ReportSource = reporte
                                 For ntablas = 1 To reporte.Database.Tables.Count
                                     reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                 Next ntablas
                                 frmvistasprevias.cr.ViewReport
                                 frmvistasprevias.Caption = "Reporte de Movimientos"
                                 frmvistasprevias.Show 1
                                 Set reporte = Nothing
                                 rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                              
                                 rsaux4.Open "select * from TB_EXISTENCIAS_SERIES WHERE vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_eMO_numero = " + Str(var_numero_folio)
                                 If Not rsaux4.EOF Then
                                    Set reporte = appl.OpenReport(App.Path + "\rep_entradas_numero_serie.rpt")
                                    frmvistasprevias.cr.ReportSource = reporte
                                    reporte.RecordSelectionFormula = "{VW_EXISTENCIAS_SERIES.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and {VW_EXISTENCIAS_SERIES.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' AND {VW_EXISTENCIAS_SERIES.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_Destino + "' and {VW_EXISTENCIAS_SERIES.vcha_mov_movimiento_id} = '" + var_clave_movimiento + "' and {VW_EXISTENCIAS_SERIES.INTE_EMO_NUMERO} = " + CStr(var_numero_folio)
                                    For ntablas = 1 To reporte.Database.Tables.Count
                                        reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                    Next ntablas
                                    frmvistasprevias.cr.ViewReport
                                    frmvistasprevias.Caption = "Reporte de Movimientos"
                                    frmvistasprevias.Show
                                    Set reporte = Nothing
                                 End If
                                 rsaux4.Close
                           
                              End If
                           
                              var_m = 0
                              If var_empresa = "18000" Then
                                 If var_clave_movimiento = "DT" Then
                                    rsaux9.Open "select vcha_age_agente_id from tb_clientes where vcha_cli_clave_id = '" + var_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux9.EOF Then
                                       var_clave_agente_correo = IIf(IsNull(rsaux9!VCHA_AGE_AGENTE_ID), "", rsaux9!VCHA_AGE_AGENTE_ID)
                                    End If
                                    rsaux9.Close
                                    rsaux9.Open "select * from tb_agentes where vcha_age_agente_id =  '" + var_clave_agente_correo + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux9.EOF Then
                                       var_correo_electronico = IIf(IsNull(rsaux9!VCHA_AGE_EMAIL), "", rsaux9!VCHA_AGE_EMAIL)
                                       var_nombre_agente_correo = IIf(IsNull(rsaux9!VCHA_AGE_NOMBRE), "", rsaux9!VCHA_AGE_NOMBRE)
                                    End If
                                    rsaux9.Close
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
                                       MAPIMessages1.MsgSubject = "Devoluci?n de tienda " + Me.txt_archivo
                                       MAPIMessages1.MsgNoteText = "Se anexa informaci?n de la devoluci?n  " + Me.txt_archivo
                                       var_Archivo = App.Path & "\Devolucion_" + Trim(Me.txt_archivo) + ".txt"
                                       Open (App.Path & "\Devolucion_" + Trim(Me.txt_archivo) + ".txt") For Output As #1
                                       Print #1, "Se genero la devoluci?n " + Trim(Me.txt_archivo) + " con los siguientes datos"
                                       Print #1, ""
                                       Print #1, "Cliente: " + Trim(var_nombre_agente_correo)
                                       Print #1, ""
                                       rsaux9.Open "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_Alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                                       Print #1, "CODIGO       DESCRIPCION                                      CANTIDAD"
                                       Print #1, "======================================================================"
                                       var_total_correo = 0
                                       While Not rsaux9.EOF
                                             var_linea = ""
                                             var_cantidad_correo = Format(IIf(IsNull(rsaux9!floa_ent_cantidaD), 0, rsaux9!floa_ent_cantidaD), "###,###,##0.00")
                                             var_total_correo = var_total_correo + IIf(IsNull(rsaux9!floa_ent_cantidaD), 0, rsaux9!floa_ent_cantidaD)
                                             rsaux8.Open "select vcha_Art_nombre_espa?ol from tb_articulos where vcha_Art_articulo_id = '" + rsaux9!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                             If Not rsaux8.EOF Then
                                                var_nombre_articulo_correo = Mid(IIf(IsNull(rsaux8!vcha_Art_nombre_espa?ol), "", rsaux8!vcha_Art_nombre_espa?ol), 1, 40)
                                             End If
                                             rsaux8.Close
                                             For var_j = Len(var_cantidad_correo) To 14
                                                 var_cantidad_correo = " " + var_cantidad_correo
                                             Next var_j
                                             For var_j = Len(var_nombre_articulo_correo) To 40
                                                 var_nombre_articulo_correo = var_nombre_articulo_correo + " "
                                             Next var_j
                                             var_linea = rsaux9!VCHA_ART_ARTICULO_ID + " " + var_nombre_articulo_correo + " " + var_cantidad_correo
                                             Print #1, var_linea
                                             rsaux9.MoveNext
                                       Wend
                                       Print #1, "======================================================================"
                                       rsaux9.Close
                                       var_total_correo_str = Format(var_total_correo, "###,###,##0.00")
                                       For var_j = Len(var_total_correo_str) To 14
                                           var_total_correo_str = " " + var_total_correo_str
                                       Next var_j
                                       var_linea = "                                       POR UN TOTAL DE:" + var_total_correo_str
                                       Print #1, var_linea
                                       Print #1, ""
                                       Close #1
                                       MAPIMessages1.AttachmentPathName = var_Archivo
                                       MAPIMessages1.Send True
                                       If MAPISession1.SessionID > 0 Then
                                          MAPISession1.SignOff
                                       End If
                                    Else
                                    End If ' fin del correo
                                 End If ' fin del movimiento dt
                              End If 'fin empresa 18
                           Else
                              MsgBox "El movimiento no tiene un reporte asociado", vbOKOnly, "ATENCION"
                           End If
                        Else
                           MsgBox "No se pudo cerrar el movimiento kanban", vbOKOnly, "ATENCION"
                        End If 'vaR_posible
                     Else
                        If var_empresa <> "18" Then
                           MsgBox "El movimiento no a afectado el inventario intentelo nuevamente o consulte a sistemas", vbOKOnly, "ATENCION"
                           rsaux4.Open "update tb_encabezado_movimientos set char_emo_estatus = '' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                           var_estatus_movimiento = ""
                           Me.txt_codigo.Enabled = True
                        Else
                           MsgBox "El movimiento a sido cerrado", vbOKOnly, "ATENCION"
                        End If
                     End If
                     rsaux2.Close
                  Else
                     MsgBox "A surgido un problema al afectar ORACLE vuelva a intentar de nuevo", vbOKOnly, "ATENCION"
                  End If
                  txt_codigo.Enabled = False
                  txt_foco.Enabled = False
               Else
                  MsgBox "No se puede cerrar el movimiento ya que hay mercancia negativa en los almacenes de apartado", vbOKOnly, "ATENCION"
               End If
            End If
         End If
      Else
         MsgBox "No se a seleccionado ning?n movimiento", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se puede cerrar el movimiento ya que hay mercancia negativa en los almacenes de apartado", vbOKOnly, "ATENCION"
   End If
   End If
   If var_clave_movimiento = "EP" Then
        ok = pro_DatosTransito(CStr(var_numero_folio), var_almacen_Destino, "")
   End If
   'cnn.RollbackTrans
   Exit Sub
salir:
   If rs.State = 1 Then
      rs.Close
   End If
   'MsgBox Err.Description, vbCritical, "sid"
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
   'If rsaux12.State = 1 Then
   '  rsaux12.Close
   'End If
    cnn.RollbackTrans
    'cnn.Open var_conexion_string_distribucion
    rs.Open "select * from tb_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_Ent_numero = " + Me.txt_folio, cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          rsaux.Open "delete from tb_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_Ent_numero = " + Me.txt_folio + " and vcha_Art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "' and inte_Ent_consecutivo = " + CStr(rs!inte_Ent_consecutivo), cnn, adOpenDynamic, adLockOptimistic
          rs.MoveNext
    Wend
    rs.Close
    rs.Open "update tb_encabezado_movimientos set char_emo_estatus = '' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_Emo_numero = " + Me.txt_folio, cnn, adOpenDynamic, adLockOptimistic
    var_estatus_movimiento = ""
    MsgBox "A surgido un error al imprimir el documento consulte a sistemas", vbOKOnly, "ATENCION"
End Sub


Private Sub cmd_nuevo_Click()
   Me.lbl_transito = ""
   Me.cmd_tipo_lectura.Visible = False
   If var_numero_folio > 0 Then
     rs.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
   End If
   lbl_cancelado = ""
   cmd_imprimir.Enabled = True
   cmd_cancelar.Enabled = True
   If var_solo_lectura = False Then
      Set TB_BLOQUEOS = New TB_BLOQUEOS
      var_global_bloqueado = 0
      ok = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, txt_archivo, Now, "", "")
   End If
   lv_entradas.ListItems.Clear
   Me.lbl_transito = ""
            var_primera_vez = True
            Me.txt_pedimento = ""
            txt_destino = ""
            txt_origen = ""
            txt_archivo = ""
            txt_factura = ""
            txt_transporto = ""
            txt_archivo.Enabled = True
            var_cantidad_enviada = 0
            var_cantidad_recibida = 0
            var_numero_folio = 0
            var_factura = ""
            txt_factura = ""
            txt_proveedor = ""
            txt_numero = ""
            lbl_recibidos = ""
            lbl_enviados = ""
            txt_folio = ""
            txt_codigo = ""
            var_estatus_movimiento = ""
            txt_archivo.SetFocus
            If var_clave_movimiento = "EC" Then
               txt_factura.Enabled = True
            End If

End Sub



Private Sub cmd_pasar_todos_Click()
   Dim pError As ADODB.Error
   Dim var_codigo_barras_caja As String
   Dim var_actualiza As Boolean
   Dim var_inserta As Boolean
   Dim bandera_suma As Boolean
   Dim var_cantidad As Variant
   Dim var_costo As Variant
   Dim var_precio As Variant
   Dim var_consecutivo_serie  As Double
   Dim var_posible As Boolean
   Dim var_P_RC_LINEA_ID As Double
   Dim var_P_RC_NUMERO_LINEA As Double
   Set TB_ARCH_COMPARACION_M = New TB_ARCH_COMPARACION_M
   Set TB_ARCH_COMPARACION_I = New TB_ARCH_COMPARACION_I
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   Dim var_posible_lectura_kanban As Boolean
   'On Error GoTo salir:
   If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
      MsgBox "El movimiento ya fue impreso", vbOKOnly, "ATENCION"
   Else
   For var_jj = 1 To Me.lv_entradas.ListItems.Count
       Me.lv_entradas.ListItems.item(var_jj).Selected = True
       var_cantidad_leida = CDbl(Me.lv_entradas.selectedItem.SubItems(5))
       Me.txt_codigo = lv_entradas.selectedItem
       cnn.CommandTimeout = 360
       If var_posible_kanban = 1 Then
          var_global_aceptar_demas = 0
       End If
       If var_empresa <> "18" Then
          If var_empresa <> "06" Then
             var_costo_tela = 0
          End If
       End If
       If var_clave_movimiento <> "EC" Then
          var_costo_tela = 0
       End If
       If Trim(txt_codigo.Text) <> "" Then
          lv_entradas.Font.Bold = False
          bandera_suma = False
          If var_primera_vez = True Then
             If var_tipo_documento = "V" Then
                var_inserta = False
                var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, var_folio_enviado, "", var_proveedor, var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, var_factura, "", txt_archivo, "", "B", "", "", 0, 0, 0, var_clave_moneda, var_tipo_Cambio)
                var_numero_folio = var_numero_folio_regreso
             Else
                var_inserta = False
                var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, var_folio_enviado, "", var_proveedor, var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, var_factura, "", txt_archivo, "", "B", "", "", 0, 0, 0, var_clave_moneda, var_tipo_Cambio)
                var_numero_folio = var_numero_folio_regreso
             End If
             txt_folio = var_numero_folio
             var_primera_vez = False
             var_fecha_movimiento = Date
          End If
          var_posible = True
          If var_cajas = True Then
             var_posible = True
          Else
             rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
             If Not rsaux.EOF Then
                var_posible = True
             Else
                var_posible = False
             End If
             rsaux.Close
          End If
          If var_posible = True Then
             If var_cajas = True Then
                Cadena = "select * from tb_archivo_comparacion where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_com_referencia = '"
                Cadena = Cadena + txt_archivo + "' and vcha_com_caja = '" + txt_codigo + "'"
             Else
                Cadena = "select * from tb_archivo_comparacion where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_com_referencia = '"
                Cadena = Cadena + txt_archivo + "' and vcha_art_articulo_id = '" + txt_codigo + "'"
             End If
             If rs.State = 1 Then
                rs.Close
             End If
             rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
             If Not rs.EOF Then
                valor = txt_codigo
                var_n = lv_entradas.ListItems.Count
                var_encontro = 0
                var_i = 1
                If var_cajas = True Then
                   While (var_i <= var_n)
                         lv_entradas.ListItems.item(var_i).Selected = True
                         valor = Trim(lv_entradas.selectedItem.SubItems(11))
                         If txt_codigo = valor Then
                            var_encontro = 1
                            var_i = var_n + 1
                         End If
                         var_i = var_i + 1
                   Wend
                Else
                   'While (var_i <= var_n)
                   '      lv_entradas.ListItems.Item(var_i).Selected = True
                   '      valor = Trim(lv_entradas.selectedItem)
                   '      If txt_codigo = valor Then
                   '         var_cantidad_posible = lv_entradas.selectedItem.SubItems(2)
                   '         If var_cantidad_posible < lv_entradas.selectedItem.SubItems(3) + var_cantidad_leida Then
                   '            var_encontro = 0
                   '         Else
                   '            var_encontro = 1
                   '            var_i = var_n + 1
                   '         End If
                   '      End If
                   '      var_i = var_i + 1
                   'Wend
                End If
                var_encontro = 1
                If var_encontro = 1 Then
                   If var_cajas = True Then
                      var_codigo_barras_caja = txt_codigo
                      txt_codigo = var_codigo_caja
                   End If
                   bandera_suma = True
                   var_posible = True
                   If var_tipo_documento = "V" Then
                      If (lv_entradas.selectedItem.SubItems(2) * 1) < ((lv_entradas.selectedItem.SubItems(3) * 1) + var_cantidad_leida) Then
                         var_posible = False
                      End If
                   Else
                       var_posible = True
                   End If
                   If var_posible = True Then
                      'If var_posible_kanban = 1 Then
                      '   var_global_aceptar_demas = 0
                      'Else
                      '   var_global_aceptar_demas = 1
                      'End If
                      If var_global_aceptar_demas = 0 Then
                         If (lv_entradas.selectedItem.SubItems(2) * 1) < ((lv_entradas.selectedItem.SubItems(3) * 1) + var_cantidad_leida) Then
                            var_posible = False
                         End If
                      End If
                      If var_posible = True Then
                         var_posible_lectura_kanban = True
                         If var_posible_kanban = 1 Then
                            Set TB_RESERVAR_KANBAN_ENTRADA = New TB_RESERVAR_KANBAN_ENTRADA
                            If var_kanban_es_un_kanban = "S" Then
                               var_inserta = TB_RESERVAR_KANBAN_ENTRADA.Anadir(var_kanban, var_clave_movimiento, var_numero_folio, var_almacen_Destino, Me.txt_codigo, Me.txt_archivo, "", "")
                               If var_kanban_exito = "S" Then
                                  var_posible_lectura_kanban = True
                               Else
                                  var_posible_lectura_kanban = False
                               End If
                            Else
                                'Set TB_RESERVAR_FUERA_KANBAN_ENT = New TB_RESERVAR_FUERA_KANBAN_ENT
                               'var_inserta = TB_RESERVAR_FUERA_KANBAN_ENT.Anadir(Me.txt_archivo, var_clave_movimiento, var_numero_folio, var_almacen_Destino, Me.txt_codigo, "", "")
                               var_posible_lectura_kanban = True
                            End If
                         Else
                            var_kanban_mensaje = ""
                            var_posible_lectura_kanban = True
                         End If
                         If var_posible_lectura_kanban = True Then
                            'lv_entradas.selectedItem.Selected = True
                            lv_entradas.selectedItem.EnsureVisible
                            lv_entradas.selectedItem.SubItems(3) = Format(lv_entradas.selectedItem.SubItems(3) + var_cantidad_leida, "###,###,##0.00")
                            lv_entradas.selectedItem.SubItems(4) = Format(lv_entradas.selectedItem.SubItems(4) + var_cantidad_leida, "###,###,##0.00")
                            lv_entradas.selectedItem.SubItems(5) = Format(lv_entradas.selectedItem.SubItems(2) - lv_entradas.selectedItem.SubItems(3), "###,###,##0.00")
                            var_consecutivo = lv_entradas.selectedItem.SubItems(9)
                            var_costo = lv_entradas.selectedItem.SubItems(6)
                            var_precio = lv_entradas.selectedItem.SubItems(7)
                            var_cantidad = lv_entradas.selectedItem.SubItems(4)
                            var_a?o = 2005
                         
                            If IsNumeric(lv_entradas.selectedItem.SubItems(14)) Then
                               VAR_RC_NL = lv_entradas.selectedItem.SubItems(14) * 1
                            Else
                               VAR_RC_NL = 0
                            End If
                            If IsNumeric(lv_entradas.selectedItem.SubItems(13)) Then
                               VAR_RC_LINEA_ID = lv_entradas.selectedItem.SubItems(13) * 1
                            Else
                               VAR_RC_LINEA_ID = 0
                            End If
                            var_P_RC_NUMERO_LINEA = VAR_RC_NL
                            var_P_RC_LINEA_ID = VAR_RC_LINEA_ID
                            lbl_recibidos = Format(CDbl(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                            var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                            var_actualiza = TB_ARCH_COMPARACION_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), var_tipo_proveedor, var_origen, txt_codigo, var_cantidad_leida, txt_archivo, var_consecutivo)
                         Else
                            txt_codigo = ""
                            If var_kanban_mensaje = "" Then
                               frmmensaje.lbl_mensaje = "Debe de leer kanbans"
                            Else
                               frmmensaje.lbl_mensaje = var_kanban_mensaje
                            End If
                            frmmensaje.Show 1
                            GoTo salir:
                         End If
                      Else
                         txt_codigo = ""
                         frmmensaje.lbl_mensaje = "La cantidad exede a la cantidad en la relaci?n"
                         frmmensaje.Show 1
                         'MsgBox "La cantidad exede a la cantidad en la relaci?n", vbOKOnly, "ATENCION"
                      End If
                   Else
                      txt_codigo = ""
                      frmmensaje.lbl_mensaje = "La cantidad exede a la incluida en la salida a vistas"
                      frmmensaje.Show 1
                      'MsgBox "La cantidad exede a la incluida en la salida a vistas", vbOKOnly, "ATENCION"
                   End If
                Else
                   valor = txt_codigo
                   'Set itmfound = lv_entradas.findItem(valor, lvwText, , lvwPartial)
                   'itmfound.EnsureVisible
                   'itmfound.Selected = True
                   bandera_suma = True
                   var_posible = True
                   If var_tipo_documento = "V" Then
                      If (lv_entradas.selectedItem.SubItems(2) * 1) < ((lv_entradas.selectedItem.SubItems(3) * 1) + var_cantidad_leida) Then
                         var_posible = False
                      End If
                   Else
                      var_posible = True
                   End If
                   If var_posible = True Then
                      If var_global_aceptar_demas = 0 Then
                         If (lv_entradas.selectedItem.SubItems(2) * 1) < ((lv_entradas.selectedItem.SubItems(3) * 1) + var_cantidad_leida) Then
                            var_posible = False
                         End If
                      End If
                      If var_posible_kanban = 1 Then
                         If var_kanban_es_un_kanban = "S" Then
                            Set TB_RESERVAR_KANBAN_ENTRADA = New TB_RESERVAR_KANBAN_ENTRADA
                            If var_kanban_es_un_kanban = "S" Then
                               var_inserta = TB_RESERVAR_KANBAN_ENTRADA.Anadir(var_kanban, var_clave_movimiento, var_numero_folio, var_almacen_Destino, Me.txt_codigo, Me.txt_archivo, "", "")
                               If var_kanban_exito = "S" Then
                                  var_posible_lectura_kanban = True
                               Else
                                  
                                  var_posible_lectura_kanban = False
                                  var_posible = False
                                  txt_codigo = ""
                                  frmmensaje.lbl_mensaje = var_kanban_mensaje
                                  frmmensaje.Show 1
                           
                               End If
                            Else
                               'Set TB_RESERVAR_FUERA_KANBAN_ENT = New TB_RESERVAR_FUERA_KANBAN_ENT
                               'var_inserta = TB_RESERVAR_FUERA_KANBAN_ENT.Anadir(Me.txt_archivo, var_clave_movimiento, var_numero_folio, var_almacen_Destino, Me.txt_codigo, "", "")
                               var_posible_lectura_kanban = True
                            End If
                         Else
                            txt_codigo = ""
                            frmmensaje.lbl_mensaje = "Se deben de leer etiquetas Kanban"
                            frmmensaje.Show 1
                         End If
                      Else
                         var_posible_lectura_kanban = True
                      End If
                      If var_posible_lectura_kanban = True Then
                         If var_posible = True Then
                            lv_entradas.selectedItem.Selected = True
                            lv_entradas.selectedItem.EnsureVisible
                            lv_entradas.selectedItem.SubItems(3) = Format(lv_entradas.selectedItem.SubItems(3) + var_cantidad_leida, "###,###,##0.00")
                            lv_entradas.selectedItem.SubItems(4) = Format(lv_entradas.selectedItem.SubItems(4) + var_cantidad_leida, "###,###,##0.00")
                            lv_entradas.selectedItem.SubItems(5) = Format(lv_entradas.selectedItem.SubItems(2) - lv_entradas.selectedItem.SubItems(3), "###,###,##0.00")
                            var_consecutivo = lv_entradas.selectedItem.SubItems(9)
                            var_costo = lv_entradas.selectedItem.SubItems(6)
                            var_precio = lv_entradas.selectedItem.SubItems(7)
                            var_cantidad = lv_entradas.selectedItem.SubItems(4)
                            lbl_recibidos = Format(CDbl(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                            var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                            var_actualiza = TB_ARCH_COMPARACION_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), var_tipo_proveedor, var_origen, txt_codigo, var_cantidad_leida, txt_archivo, var_consecutivo)
                            var_renglon = var_n
                         Else
                            txt_codigo = ""
                            frmmensaje.lbl_mensaje = "La cantidad exede a la cantidad en la relaci?n"
                            frmmensaje.Show 1
                            'MsgBox "La cantidad exede a la cantidad en la relaci?n", vbOKOnly, "ATENCION"
                         End If
                      End If
                   Else
                      txt_codigo = ""
                      frmmensaje.lbl_mensaje = "La cantidad exede a la incluida en la salida a vistas"
                      frmmensaje.Show 1
                      'MsgBox "La cantidad exede a la incluida en la salida a vistas", vbOKOnly, "ATENCION"
                   End If
                End If
             Else
                var_posible = True
                If var_tipo_documento = "V" Then
                   txt_codigo = ""
                   frmmensaje.lbl_mensaje = "El art?culo no existe dentro de la salida a vistas"
                   frmmensaje.Show 1
                   'MsgBox "El art?culo no existe dentro de la salida a vistas", vbOKOnly, "ATENCION"
                   var_posible = False
                Else
                   If var_global_aceptar_demas = 1 Then
                      rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                      If Not rsaux.EOF Then
                         rsaux3.Open "select max(inte_com_consecutivo) as maximo from tb_Archivo_comparacion where vcha_emp_empresa_id = '" + var_empresa + "' and  vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "'  and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "'  and INTE_COM_NUMERO = " + CStr(CDbl(var_folio_enviado)) + " and VCHA_COM_REFERENCIA = '" + txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
                         If Not rsaux3.EOF Then
                            var_consecutivo = IIf(IsNull(rsaux3!maximo), 0, rsaux3!maximo) + 1
                         Else
                            var_consecutivo = 1
                         End If
                         rsaux3.Close
                         Set list_item = lv_entradas.ListItems.Add(, , rsaux(0).Value)
                         list_item.SubItems(1) = IIf(IsNull(rsaux(1).Value), "", rsaux(1).Value)
                         list_item.SubItems(2) = Format(0, "###,###,##0.00")
                         list_item.SubItems(3) = Format(var_cantidad_leida, "###,###,##0.00")
                         list_item.SubItems(4) = Format(var_cantidad_leida, "###,###,##0.00")
                         list_item.SubItems(5) = Format(list_item.SubItems(2) - list_item.SubItems(3), "###,###,##0.00")
                         list_item.SubItems(6) = IIf(IsNull(rsaux(3).Value), "", rsaux(3).Value)
                         list_item.SubItems(7) = IIf(IsNull(rsaux(2).Value), "", rsaux(2).Value)
                         list_item.SubItems(8) = 0
                         list_item.SubItems(9) = var_consecutivo
                         list_item.SubItems(10) = 2005
                         var_n = lv_entradas.ListItems.Count
                         lv_entradas.ListItems.item(var_n).Selected = True
                         var_precio = rsaux(2).Value
                         var_costo = rsaux!mone_Art_costo_estandar
                     
                         If var_entrada_calidad = True Then
                            rsaux2.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_costeo + "' and vcha_art_articulo_id = '" + Trim(txt_codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
                            If Not rsaux2.EOF Then
                               var_costo = rsaux2!FLOA_eXI_COSTO
                            End If
                            rsaux2.Close
                         End If
                         bandera_suma = True
                         lbl_recibidos = Format(CDbl(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                         var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                         If var_cajas = True Then
                            ok = TB_ARCH_COMPARACION_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), Date, var_tipo_proveedor, var_origen, txt_codigo, var_costo, 0, var_cantidad_leida, var_transporto, txt_archivo, 0, var_consecutivo, 2005, txt_codigo, var_peso_caja)
                            var_posible_lectura_kanban = True
                         Else
                            ok = TB_ARCH_COMPARACION_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), Date, var_tipo_proveedor, var_origen, txt_codigo, var_costo, 0, var_cantidad_leida, var_transporto, txt_archivo, 0, var_consecutivo, 2005, "", 0)
                            var_posible_lectura_kanban = True
                         End If
                         var_renglon = lv_entradas.ListItems.Count
                      Else
                         txt_codigo = ""
                         frmmensaje.lbl_mensaje = "El art?culo no existe"
                         frmmensaje.Show 1
                         'MsgBox "El art?culo no existe", vbOKOnly, "ATENCION"
                         bandera_suma = False
                      End If
                      rsaux.Close
                   Else
                      txt_codigo = ""
                      frmmensaje.lbl_mensaje = "El art?culo no se encuentra dentro de la relaci?n"
                      frmmensaje.Show 1
                      'MsgBox "El art?culo no se encuentra dentro de la relaci?n", vbOKOnly, "ATENCION"
                   End If
                End If
                If var_global_aceptar_demas = 0 Then
                   var_posible = False
                Else
                   var_posible = True
                End If
             End If
             If rs.State = 1 Then
                rs.Close
             End If
          Else
             txt_codigo = ""
             frmmensaje.lbl_mensaje = "El art?culo no existe"
             frmmensaje.Show 1
             'MsgBox "El art?culo no existe", vbOKOnly, "ATENCION"
             rsaux.Close
          End If
          If bandera_suma = True Then
             If var_tipo_documento = "V" Then
                If var_posible = True Then
               
                   Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "' and inte_ent_consecutivo = " + CStr(var_consecutivo)
                   rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                   If Not rs.EOF Then
                      var_inserta = False
                      rsaux3.Open "UPDATE TB_TEMPORAL_ENTRADAS SET VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "', FLOA_ENT_CANTIDAD = FLOA_ENT_CANTIDAD + " + CStr(var_cantidad_leida) + " Where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_ENT_NUMERO= '" + CStr(var_numero_folio) + "' AND VCHA_ART_ARTICULO_ID  = '" + txt_codigo + "' and inte_ent_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                      rs.Close
                   Else
                      var_inserta = False
                      If var_empresa = "18" And var_proveedor = "2458" Then
                         rsaux3.Open "INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN,INTE_ENT_CONSECUTIVO, INTE_ENT_A?O, P_RC_NUMERO_LINEA, P_RC_LINEA_ID) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo_tela) + ", " + CStr(var_precio) + ", 0,'" + var_almacen_origen + "', " + CStr(var_consecutivo) + ", " + CStr(var_a?o) + ", " + CStr(var_P_RC_LINEA_ID) + "," + CStr(var_P_RC_NUMERO_LINEA) + ")", cnn, adOpenDynamic, adLockOptimistic
                      Else
                         rsaux3.Open "INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN,INTE_ENT_CONSECUTIVO, INTE_ENT_A?O, P_RC_NUMERO_LINEA, P_RC_LINEA_ID) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo + var_costo_tela) + ", " + CStr(var_precio) + ", 0,'" + var_almacen_origen + "', " + CStr(var_consecutivo) + ", " + CStr(var_a?o) + ", " + CStr(var_P_RC_LINEA_ID) + "," + CStr(var_P_RC_NUMERO_LINEA) + ")", cnn, adOpenDynamic, adLockOptimistic
                      End If
                      rs.Close
                   End If
                   If var_numero_serie = 1 Then
                      Cadena = "select MAX(INTE_EXI_CONSECUTIVO) from TB_EXISTENCIAS_SERIES WHERE vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_eMO_numero = " + Str(var_numero_folio)
                      var_consecutivo_serie = 0
                      rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                      If Not rs.EOF Then
                         var_consecutivo_serie = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                      Else
                         var_consecutivo_serie = 1
                      End If
                      rs.Close
                      rsaux.Open "insert into tb_exiStencias_series INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, VCHA_ART_ARTICULO_ID, VCHA_ART_NUMERO_SERIE, INTE_EXI_CONSECUTIVO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', '" + Me.txt_numero_serie + "'," + CStr(var_consecutivo_serie) + ")", cnn, adOpenDynamic, adLockOptimistic
                   End If
                End If
             Else
                If var_posible = True Then
                   var_posible_lectura_kanban = True
                   If var_posible_lectura_kanban = True Then
                      If var_posible_kanban = 1 Then
                         If var_kanban_es_un_kanban = "N" Or var_kanban_es_un_kanban = "" Then
                            Set TB_RESERVAR_FUERA_KANBAN_ENT = New TB_RESERVAR_FUERA_KANBAN_ENT
                            var_inserta = TB_RESERVAR_FUERA_KANBAN_ENT.Anadir(Me.txt_archivo, var_clave_movimiento, var_numero_folio, var_almacen_Destino, Me.txt_codigo, "", "")
                            If var_kanban_exito = "N" Then
                               var_posible = False
                            End If
                         End If
                      End If
                      If var_posible = True Then
                         Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "' and inte_ent_consecutivo = " + CStr(var_consecutivo)
                         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                         If Not rs.EOF Then
                            var_inserta = False
                            rsaux3.Open "UPDATE TB_TEMPORAL_ENTRADAS SET VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "', FLOA_ENT_CANTIDAD = FLOA_ENT_CANTIDAD + " + CStr(var_cantidad_leida) + " Where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_ENT_NUMERO= '" + CStr(var_numero_folio) + "' AND VCHA_ART_ARTICULO_ID  = '" + txt_codigo + "' and inte_ent_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                            rs.Close
                         Else
                            var_inserta = False
                            var_costo = IIf(IsNull(var_costo), 0, var_costo)
                            If var_costo = 0 Or var_clave_movimiento = "DT" Then
                               rsaux4.Open "select floa_exi_costo_2005, floa_Exi_costo from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                               If Not rsaux4.EOF Then
                                  var_costo = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                               Else
                                  var_costo = 0
                               End If
                               rsaux4.Close
                               If var_costo = 0 And var_almacen_Destino = "14" Then
                                  rsaux4.Open "select floa_exi_costo_2005, floa_Exi_costo from tb_existencias where vcha_alm_almacen_id = '8' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                  If Not rsaux4.EOF Then
                                     var_costo = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                  Else
                                     var_costo = 0
                                  End If
                                  rsaux4.Close
                               End If
                               If var_costo = 0 And var_almacen_Destino = "11" Then
                                  rsaux4.Open "select floa_exi_costo_2005, floa_Exi_costo from tb_existencias where vcha_alm_almacen_id = '8' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                  If Not rsaux4.EOF Then
                                     var_costo = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                  Else
                                     var_costo = 0
                                  End If
                                  rsaux4.Close
                               End If
                               If var_costo = 0 Then
                                  rsaux4.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                  If Not rsaux4.EOF Then
                                     var_costo = IIf(IsNull(rsaux4!mone_Art_costo_estandar), 0, rsaux4!mone_Art_costo_estandar)
                                  End If
                                  rsaux4.Close
                               End If
                            End If
                            If var_empresa = "18" And var_proveedor = "2458" Then
                               rsaux3.Open "INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN, INTE_ENT_CONSECUTIVO, INTE_ENT_A?O, P_RC_LINEA_ID, P_RC_NUMERO_LINEA) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo_tela) + ", " + CStr(var_precio) + ", 0,'" + var_almacen_origen + "', " + CStr(var_consecutivo) + ", " + CStr(var_a?o) + "," + CStr(var_P_RC_LINEA_ID) + "," + CStr(var_P_RC_NUMERO_LINEA) + ")", cnn, adOpenDynamic, adLockOptimistic
                            Else
                               If var_empresa = "06" And var_proveedor = "2458" Then
                                  rsaux3.Open "INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN, INTE_ENT_CONSECUTIVO, INTE_ENT_A?O, P_RC_LINEA_ID, P_RC_NUMERO_LINEA) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo_tela) + ", " + CStr(var_precio) + ", 0,'" + var_almacen_origen + "', " + CStr(var_consecutivo) + ", " + CStr(var_a?o) + "," + CStr(var_P_RC_LINEA_ID) + "," + CStr(var_P_RC_NUMERO_LINEA) + ")", cnn, adOpenDynamic, adLockOptimistic
                               Else
                                  If var_precio = "" Then
                                     var_precio = 0
                                  End If
    
                                  rsaux3.Open "INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN, INTE_ENT_CONSECUTIVO, INTE_ENT_A?O, P_RC_LINEA_ID, P_RC_NUMERO_LINEA) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo + var_costo_tela) + ", " + CStr(var_precio) + ", 0,'" + var_almacen_origen + "', " + CStr(var_consecutivo) + ", " + CStr(var_a?o) + "," + CStr(var_P_RC_LINEA_ID) + "," + CStr(var_P_RC_NUMERO_LINEA) + ")", cnn, adOpenDynamic, adLockOptimistic
                               End If
                            End If
                            rs.Close
                         End If
                        If var_clave_movimiento = "ETA" Or var_clave_movimiento = "EI" Or var_clave_movimiento = "EP" Then
                           '2
                           rsaux3.Open "update tb_transito set floa_tra_cantidad_recibida = isnull(floa_Tra_cantidad_recibida,0) + " + CStr(var_cantidad_leida) + " where vcha_tra_nota_envio = '" + Me.lbl_transito + "' and vcha_Art_articulo_recivo = '" + Me.txt_codigo + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
                        End If
                      Else
                      'cuando no es kanban
                         lv_entradas.selectedItem.Selected = True
                         lv_entradas.selectedItem.EnsureVisible
                         lv_entradas.selectedItem.SubItems(3) = Format(lv_entradas.selectedItem.SubItems(3) - var_cantidad_leida, "###,###,##0.00")
                         lv_entradas.selectedItem.SubItems(4) = Format(lv_entradas.selectedItem.SubItems(4) - var_cantidad_leida, "###,###,##0.00")
                         lv_entradas.selectedItem.SubItems(5) = Format(lv_entradas.selectedItem.SubItems(5) + var_cantidad_leida, "###,###,##0.00")
                         var_consecutivo = lv_entradas.selectedItem.SubItems(9)
                         var_costo = lv_entradas.selectedItem.SubItems(6)
                         var_precio = lv_entradas.selectedItem.SubItems(7)
                         var_cantidad = lv_entradas.selectedItem.SubItems(4)
                         lbl_recibidos = Format(CDbl(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                         var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                         var_actualiza = TB_ARCH_COMPARACION_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), var_tipo_proveedor, var_origen, txt_codigo, 0 - var_cantidad_leida, txt_archivo, var_consecutivo)
                         var_renglon = var_n
                         'fin de cuando no es un kanban
                         frmmensaje.lbl_mensaje = var_kanban_mensaje
                         frmmensaje.Show 1
                      End If
                      
                   End If 'kanban
                   If var_numero_serie = 1 Then
                      'Cadena = "select MAX(INTE_EXI_CONSECUTIVO) from TB_EXISTENCIAS_SERIES WHERE vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_eMO_numero = " + Str(var_numero_folio)
                      Cadena = "select MAX(INTE_EXI_CONSECUTIVO) from TB_EXISTENCIAS_SERIES WHERE VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'"
                      var_consecutivo_serie = 0
                      rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                      If Not rs.EOF Then
                         var_consecutivo_serie = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                      Else
                         var_consecutivo_serie = 1
                      End If
                      rs.Close
                      var_consecutivo_serie_str = CStr(var_consecutivo_serie)
                      If Len(var_consecutivo_serie_str) = 1 Then
                         var_consecutivo_serie_str = "000" + Trim(var_consecutivo_serie_str)
                      Else
                         If Len(var_consecutivo_serie_str) = 2 Then
                            var_consecutivo_serie_str = "00" + Trim(var_consecutivo_serie_str)
                         Else
                            If Len(var_consecutivo_serie_str) = 3 Then
                               var_consecutivo_serie_str = "0" + Trim(var_consecutivo_serie_str)
                            Else
                               If Len(var_consecutivo_serie_str) = 4 Then
                                  var_consecutivo_serie_str = Trim(var_consecutivo_serie_str)
                               Else
                               End If
                            End If
                         End If
                      End If
                      Me.txt_numero_serie = Me.txt_codigo + var_consecutivo_serie_str
                      rsaux.Open "insert into tb_exiStencias_series (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, VCHA_ART_ARTICULO_ID, VCHA_ART_NUMERO_SERIE, INTE_EXI_CONSECUTIVO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', '" + Me.txt_numero_serie + "'," + CStr(var_consecutivo_serie) + ")", cnn, adOpenDynamic, adLockOptimistic
                      rsaux9.Open "select substring(vcha_art_nombre_espa?ol,1,23) AS descripcion, substring(vcha_art_nombre_espa?ol,24,46) as descripcion_2 from tb_articulos where vcha_art_Articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                      If Not rsaux9.EOF Then
                         var_descripcion_etiqueta = IIf(IsNull(rsaux9!descripcion), "", rsaux9!descripcion)
                         var_descripcion_etiqueta2 = IIf(IsNull(rsaux9!descripcion_2), "", rsaux9!descripcion_2)
                      End If
                      rsaux9.Close
                      Open (App.Path & "\etiqueta.bat") For Output As #2
                      Print #2, "copy " + App.Path + "\etiqueta.txt lpt1"
                      Open (App.Path & "\etiqueta.txt") For Output As #1
                      Close #2
                      Print #1, "US"
                      Print #1, "q392"
                      Print #1, "Q256,24+0"
                      Print #1, "S2"
                      Print #1, "D8"
                      Print #1, "ZT"
                      Print #1, "TTh: m"
                      Print #1, "TDy2.mn.dd"
                      Print #1, "A39,20,0,3,1,1,N,""" + var_descripcion_etiqueta + """"
                      Print #1, "A39,40,0,3,1,1,N,""" + var_descripcion_etiqueta2 + """"
                      Print #1, "B39,90,0,3,2,4,101,B,""" + Trim(Me.txt_numero_serie) + """"
                      Print #1, "P1"
                      Close #1
                      x = Shell(App.Path & "\etiqueta.bat", vbHide)
                   End If
                End If
             End If
             bandera_suma = False
             var_renglon = lv_entradas.selectedItem.Index
             Call ilumina_grid
          End If
          If var_n > 11 Then
             lv_entradas.ColumnHeaders(2).Width = 4700.01
          Else
             lv_entradas.ColumnHeaders(2).Width = 4930.01
          End If
          If var_cajas = True Then
             If var_peso_correcto = True Then
                var_costo_tela = 0
                Me.txt_codigo.SetFocus
             Else
                txt_codigo = var_codigo_barras_caja
                Me.txt_codigo_caja.SetFocus
             End If
          Else
             var_costo_tela = 0
             Me.txt_codigo = ""
             'txt_codigo.SetFocus
          End If
       End If
    Next var_jj
    If Me.lv_entradas.ListItems.Count > 0 Then
       MsgBox "Se a terminado de cargar el movimiento", vbOKOnly, "ATENCION"
    End If
   End If
   Exit Sub
    
salir:
   If Err.Number = -2147217871 Then
      Resume
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

Private Sub cmd_salir_Click()
   If Me.txt_folio <> "" Then
      var_si = MsgBox("?Deseas cerrar el movimiento?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         Unload Me
      End If
   Else
      Unload Me
   End If
End Sub

Private Sub cmd_tipo_lectura_Click()
   Dim pError As ADODB.Error
   Dim var_codigo_barras_caja As String
   Dim var_actualiza As Boolean
   Dim var_inserta As Boolean
   Dim bandera_suma As Boolean
   Dim var_cantidad As Variant
   Dim var_costo As Variant
   Dim var_precio As Variant
   Dim var_consecutivo_serie  As Double
   Dim var_posible As Boolean
   Dim var_P_RC_LINEA_ID As Double
   Dim var_P_RC_NUMERO_LINEA As Double
   Set TB_ARCH_COMPARACION_M = New TB_ARCH_COMPARACION_M
   Set TB_ARCH_COMPARACION_I = New TB_ARCH_COMPARACION_I
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   Dim var_posible_lectura_kanban As Boolean
   'On Error GoTo salir:
   If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
      MsgBox "El movimiento ya fue impreso", vbOKOnly, "ATENCION"
   Else
      var_si = MsgBox("?Se va a generar el movimiento?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         For var_jj = 1 To Me.lv_entradas.ListItems.Count
             Me.lv_entradas.ListItems.item(var_jj).Selected = True
             var_cantidad_leida = CDbl(Me.lv_entradas.selectedItem.SubItems(3))
             Me.txt_codigo = lv_entradas.selectedItem
             cnn.CommandTimeout = 360
             If var_posible_kanban = 1 Then
                var_global_aceptar_demas = 0
             End If
             If var_empresa <> "18" Then
                If var_empresa <> "06" Then
                   var_costo_tela = 0
                End If
             End If
             If var_clave_movimiento <> "EC" Then
                var_costo_tela = 0
             End If
             If Trim(txt_codigo.Text) <> "" Then
                lv_entradas.Font.Bold = False
                bandera_suma = False
                If var_primera_vez = True Then
                   If var_tipo_documento = "V" Then
                      var_inserta = False
                      var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, var_folio_enviado, "", var_proveedor, var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, var_factura, "", txt_archivo, "", "B", "", "", 0, 0, 0, var_clave_moneda, var_tipo_Cambio)
                      var_numero_folio = var_numero_folio_regreso
                   Else
                      var_inserta = False
                      var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, var_folio_enviado, "", var_proveedor, var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, var_factura, "", txt_archivo, "", "B", "", "", 0, 0, 0, var_clave_moneda, var_tipo_Cambio)
                      var_numero_folio = var_numero_folio_regreso
                   End If
                   txt_folio = var_numero_folio
                   var_primera_vez = False
                   var_fecha_movimiento = Date
                End If
                var_posible = True
                If var_cajas = True Then
                   var_posible = True
                Else
                   rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                   If Not rsaux.EOF Then
                      var_posible = True
                   Else
                      var_posible = False
                   End If
                   rsaux.Close
                End If
                If var_posible = True Then
                   If var_cajas = True Then
                      Cadena = "select * from tb_archivo_comparacion where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_com_referencia = '"
                      Cadena = Cadena + txt_archivo + "' and vcha_com_caja = '" + txt_codigo + "'"
                   Else
                      Cadena = "select * from tb_archivo_comparacion where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_com_referencia = '"
                      Cadena = Cadena + txt_archivo + "' and vcha_art_articulo_id = '" + txt_codigo + "'"
                   End If
                   If rs.State = 1 Then
                      rs.Close
                   End If
                   rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                   If Not rs.EOF Then
                      valor = txt_codigo
                      var_n = lv_entradas.ListItems.Count
                      var_encontro = 0
                      var_i = 1
                      If var_cajas = True Then
                         While (var_i <= var_n)
                               lv_entradas.ListItems.item(var_i).Selected = True
                               valor = Trim(lv_entradas.selectedItem.SubItems(11))
                               If txt_codigo = valor Then
                                   var_encontro = 1
                                  var_i = var_n + 1
                               End If
                               var_i = var_i + 1
                         Wend
                      Else
                      End If
                      var_encontro = 1
                      If var_encontro = 1 Then
                         If var_cajas = True Then
                            var_codigo_barras_caja = txt_codigo
                            txt_codigo = var_codigo_caja
                         End If
                         bandera_suma = True
                         var_posible = True
                         If var_tipo_documento = "V" Then
                            If (lv_entradas.selectedItem.SubItems(2) * 1) < ((lv_entradas.selectedItem.SubItems(3) * 1) + var_cantidad_leida) Then
                               var_posible = False
                            End If
                         Else
                             var_posible = True
                         End If
                         If var_posible = True Then
                            If var_global_aceptar_demas = 0 Then
                               If (lv_entradas.selectedItem.SubItems(2) * 1) < ((lv_entradas.selectedItem.SubItems(3) * 1) + var_cantidad_leida) Then
                                  var_posible = False
                               End If
                            End If
                            If var_posible = True Then
                               var_posible_lectura_kanban = True
                               If var_posible_kanban = 1 Then
                                  Set TB_RESERVAR_KANBAN_ENTRADA = New TB_RESERVAR_KANBAN_ENTRADA
                                  If var_kanban_es_un_kanban = "S" Then
                                     var_inserta = TB_RESERVAR_KANBAN_ENTRADA.Anadir(var_kanban, var_clave_movimiento, var_numero_folio, var_almacen_Destino, Me.txt_codigo, Me.txt_archivo, "", "")
                                     If var_kanban_exito = "S" Then
                                        var_posible_lectura_kanban = True
                                     Else
                                        var_posible_lectura_kanban = False
                                     End If
                                  Else
                                     var_posible_lectura_kanban = True
                                  End If
                               Else
                                  var_kanban_mensaje = ""
                                  var_posible_lectura_kanban = True
                               End If
                               If var_posible_lectura_kanban = True Then
                                  lv_entradas.selectedItem.EnsureVisible
                                  lv_entradas.selectedItem.SubItems(3) = Format(lv_entradas.selectedItem.SubItems(3), "###,###,##0.00")
                                  lv_entradas.selectedItem.SubItems(4) = Format(lv_entradas.selectedItem.SubItems(4) + var_cantidad_leida, "###,###,##0.00")
                                  lv_entradas.selectedItem.SubItems(5) = Format(lv_entradas.selectedItem.SubItems(2) - lv_entradas.selectedItem.SubItems(3), "###,###,##0.00")
                                  var_consecutivo = lv_entradas.selectedItem.SubItems(9)
                                  var_costo = lv_entradas.selectedItem.SubItems(6)
                                  var_precio = lv_entradas.selectedItem.SubItems(7)
                                  var_cantidad = lv_entradas.selectedItem.SubItems(4)
                                  var_a?o = lv_entradas.selectedItem.SubItems(10)
                               
                                  If IsNumeric(lv_entradas.selectedItem.SubItems(14)) Then
                                     VAR_RC_NL = lv_entradas.selectedItem.SubItems(14) * 1
                                  Else
                                     VAR_RC_NL = 0
                                  End If
                                  If IsNumeric(lv_entradas.selectedItem.SubItems(13)) Then
                                     VAR_RC_LINEA_ID = lv_entradas.selectedItem.SubItems(13) * 1
                                  Else
                                     VAR_RC_LINEA_ID = 0
                                  End If
                                  var_P_RC_NUMERO_LINEA = VAR_RC_NL
                                  var_P_RC_LINEA_ID = VAR_RC_LINEA_ID
                                  'lbl_recibidos = Format(CDbl(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                                  'var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                                  'var_actualiza = TB_ARCH_COMPARACION_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), var_tipo_proveedor, var_origen, txt_codigo, var_cantidad_leida, txt_archivo, var_consecutivo)
                               Else
                                  txt_codigo = ""
                                  If var_kanban_mensaje = "" Then
                                     frmmensaje.lbl_mensaje = "Debe de leer kanbans"
                                  Else
                                     frmmensaje.lbl_mensaje = var_kanban_mensaje
                                  End If
                                  frmmensaje.Show 1
                                  GoTo salir:
                               End If
                            Else
                               txt_codigo = ""
                               frmmensaje.lbl_mensaje = "La cantidad exede a la cantidad en la relaci?n"
                               frmmensaje.Show 1
                            End If
                         Else
                            txt_codigo = ""
                            frmmensaje.lbl_mensaje = "La cantidad exede a la incluida en la salida a vistas"
                            frmmensaje.Show 1
                         End If
                      Else
                         valor = txt_codigo
                         bandera_suma = True
                         var_posible = True
                         If var_tipo_documento = "V" Then
                            If (lv_entradas.selectedItem.SubItems(2) * 1) < ((lv_entradas.selectedItem.SubItems(3) * 1) + var_cantidad_leida) Then
                               var_posible = False
                            End If
                         Else
                            var_posible = True
                         End If
                         If var_posible = True Then
                            If var_global_aceptar_demas = 0 Then
                               If (lv_entradas.selectedItem.SubItems(2) * 1) < ((lv_entradas.selectedItem.SubItems(3) * 1) + var_cantidad_leida) Then
                                  var_posible = False
                               End If
                            End If
                            If var_posible_kanban = 1 Then
                               If var_kanban_es_un_kanban = "S" Then
                                  Set TB_RESERVAR_KANBAN_ENTRADA = New TB_RESERVAR_KANBAN_ENTRADA
                                  If var_kanban_es_un_kanban = "S" Then
                                     var_inserta = TB_RESERVAR_KANBAN_ENTRADA.Anadir(var_kanban, var_clave_movimiento, var_numero_folio, var_almacen_Destino, Me.txt_codigo, Me.txt_archivo, "", "")
                                     If var_kanban_exito = "S" Then
                                        var_posible_lectura_kanban = True
                                     Else
                                        var_posible_lectura_kanban = False
                                        var_posible = False
                                        txt_codigo = ""
                                        frmmensaje.lbl_mensaje = var_kanban_mensaje
                                        frmmensaje.Show 1
                                     End If
                                  Else
                                     var_posible_lectura_kanban = True
                                  End If
                               Else
                                  txt_codigo = ""
                                  frmmensaje.lbl_mensaje = "Se deben de leer etiquetas Kanban"
                                  frmmensaje.Show 1
                               End If
                            Else
                               var_posible_lectura_kanban = True
                            End If
                            If var_posible_lectura_kanban = True Then
                               If var_posible = True Then
                                  lv_entradas.selectedItem.Selected = True
                                  lv_entradas.selectedItem.EnsureVisible
                                  lv_entradas.selectedItem.SubItems(3) = Format(lv_entradas.selectedItem.SubItems(3) + var_cantidad_leida, "###,###,##0.00")
                                  lv_entradas.selectedItem.SubItems(4) = Format(lv_entradas.selectedItem.SubItems(4) + var_cantidad_leida, "###,###,##0.00")
                                  lv_entradas.selectedItem.SubItems(5) = Format(lv_entradas.selectedItem.SubItems(2) - lv_entradas.selectedItem.SubItems(3), "###,###,##0.00")
                                  var_consecutivo = lv_entradas.selectedItem.SubItems(9)
                                  var_costo = lv_entradas.selectedItem.SubItems(6)
                                  var_precio = lv_entradas.selectedItem.SubItems(7)
                                  var_cantidad = lv_entradas.selectedItem.SubItems(4)
                                  lbl_recibidos = Format(CDbl(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                                  var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                                  var_actualiza = TB_ARCH_COMPARACION_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), var_tipo_proveedor, var_origen, txt_codigo, var_cantidad_leida, txt_archivo, var_consecutivo)
                                  var_renglon = var_n
                               Else
                                  txt_codigo = ""
                                  frmmensaje.lbl_mensaje = "La cantidad exede a la cantidad en la relaci?n"
                                  frmmensaje.Show 1
                               End If
                            End If
                         Else
                            txt_codigo = ""
                            frmmensaje.lbl_mensaje = "La cantidad exede a la incluida en la salida a vistas"
                            frmmensaje.Show 1
                         End If
                      End If
                   Else
                      var_posible = True
                      If var_tipo_documento = "V" Then
                         txt_codigo = ""
                         frmmensaje.lbl_mensaje = "El art?culo no existe dentro de la salida a vistas"
                         frmmensaje.Show 1
                         var_posible = False
                      Else
                         If var_global_aceptar_demas = 1 Then
                            rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                            If Not rsaux.EOF Then
                               rsaux3.Open "select max(inte_com_consecutivo) as maximo from tb_Archivo_comparacion where vcha_emp_empresa_id = '" + var_empresa + "' and  vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "'  and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "'  and INTE_COM_NUMERO = " + CStr(CDbl(var_folio_enviado)) + " and VCHA_COM_REFERENCIA = '" + txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
                               If Not rsaux3.EOF Then
                                  var_consecutivo = IIf(IsNull(rsaux3!maximo), 0, rsaux3!maximo) + 1
                               Else
                                  var_consecutivo = 1
                               End If
                               rsaux3.Close
                               Set list_item = lv_entradas.ListItems.Add(, , rsaux(0).Value)
                               list_item.SubItems(1) = IIf(IsNull(rsaux(1).Value), "", rsaux(1).Value)
                               list_item.SubItems(2) = Format(0, "###,###,##0.00")
                               list_item.SubItems(3) = Format(var_cantidad_leida, "###,###,##0.00")
                               list_item.SubItems(4) = Format(var_cantidad_leida, "###,###,##0.00")
                               list_item.SubItems(5) = Format(list_item.SubItems(2) - list_item.SubItems(3), "###,###,##0.00")
                               list_item.SubItems(6) = IIf(IsNull(rsaux(3).Value), "", rsaux(3).Value)
                               list_item.SubItems(7) = IIf(IsNull(rsaux(2).Value), "", rsaux(2).Value)
                               list_item.SubItems(8) = 0
                               list_item.SubItems(9) = var_consecutivo
                               list_item.SubItems(10) = 2005
                               var_n = lv_entradas.ListItems.Count
                               lv_entradas.ListItems.item(var_n).Selected = True
                               var_precio = rsaux(2).Value
                               var_costo = rsaux!mone_Art_costo_estandar
                           
                               If var_entrada_calidad = True Then
                                  rsaux2.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_costeo + "' and vcha_art_articulo_id = '" + Trim(txt_codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
                                  If Not rsaux2.EOF Then
                                     var_costo = rsaux2!FLOA_eXI_COSTO
                                  End If
                                  rsaux2.Close
                               End If
                               bandera_suma = True
                               lbl_recibidos = Format(CDbl(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                               var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                               If var_cajas = True Then
                                  ok = TB_ARCH_COMPARACION_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), Date, var_tipo_proveedor, var_origen, txt_codigo, var_costo, 0, var_cantidad_leida, var_transporto, txt_archivo, 0, var_consecutivo, 2005, txt_codigo, var_peso_caja)
                                  var_posible_lectura_kanban = True
                               Else
                                  ok = TB_ARCH_COMPARACION_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), Date, var_tipo_proveedor, var_origen, txt_codigo, var_costo, 0, var_cantidad_leida, var_transporto, txt_archivo, 0, var_consecutivo, 2005, "", 0)
                                  var_posible_lectura_kanban = True
                               End If
                               var_renglon = lv_entradas.ListItems.Count
                            Else
                               txt_codigo = ""
                               frmmensaje.lbl_mensaje = "El art?culo no existe"
                               frmmensaje.Show 1
                               bandera_suma = False
                            End If
                            rsaux.Close
                         Else
                            txt_codigo = ""
                            frmmensaje.lbl_mensaje = "El art?culo no se encuentra dentro de la relaci?n"
                            frmmensaje.Show 1
                         End If
                      End If
                      If var_global_aceptar_demas = 0 Then
                         var_posible = False
                      Else
                         var_posible = True
                      End If
                   End If
                   If rs.State = 1 Then
                      rs.Close
                   End If
                Else
                   txt_codigo = ""
                   frmmensaje.lbl_mensaje = "El art?culo no existe"
                   frmmensaje.Show 1
                   rsaux.Close
                End If
                If bandera_suma = True Then
                   If var_tipo_documento = "V" Then
                      If var_posible = True Then
                         Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "' and inte_ent_consecutivo = " + CStr(var_consecutivo)
                         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                         If Not rs.EOF Then
                            var_inserta = False
                            rsaux3.Open "UPDATE TB_TEMPORAL_ENTRADAS SET VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "', FLOA_ENT_CANTIDAD = FLOA_ENT_CANTIDAD + " + CStr(var_cantidad_leida) + " Where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_ENT_NUMERO= '" + CStr(var_numero_folio) + "' AND VCHA_ART_ARTICULO_ID  = '" + txt_codigo + "' and inte_ent_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                            rs.Close
                         Else
                            var_inserta = False
                            If var_empresa = "18" And var_proveedor = "2458" Then
                               rsaux3.Open "INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN,INTE_ENT_CONSECUTIVO, INTE_ENT_A?O, P_RC_NUMERO_LINEA, P_RC_LINEA_ID) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo_tela) + ", " + CStr(var_precio) + ", 0,'" + var_almacen_origen + "', " + CStr(var_consecutivo) + ", " + CStr(var_a?o) + ", " + CStr(var_P_RC_LINEA_ID) + "," + CStr(var_P_RC_NUMERO_LINEA) + ")", cnn, adOpenDynamic, adLockOptimistic
                            Else
                               rsaux3.Open "INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN,INTE_ENT_CONSECUTIVO, INTE_ENT_A?O, P_RC_NUMERO_LINEA, P_RC_LINEA_ID) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo + var_costo_tela) + ", " + CStr(var_precio) + ", 0,'" + var_almacen_origen + "', " + CStr(var_consecutivo) + ", " + CStr(var_a?o) + ", " + CStr(var_P_RC_LINEA_ID) + "," + CStr(var_P_RC_NUMERO_LINEA) + ")", cnn, adOpenDynamic, adLockOptimistic
                            End If
                            rs.Close
                         End If
                         If var_numero_serie = 1 Then
                            Cadena = "select MAX(INTE_EXI_CONSECUTIVO) from TB_EXISTENCIAS_SERIES WHERE vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_eMO_numero = " + Str(var_numero_folio)
                            var_consecutivo_serie = 0
                            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                            If Not rs.EOF Then
                               var_consecutivo_serie = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                            Else
                               var_consecutivo_serie = 1
                            End If
                            rs.Close
                            rsaux.Open "insert into tb_exiStencias_series INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, VCHA_ART_ARTICULO_ID, VCHA_ART_NUMERO_SERIE, INTE_EXI_CONSECUTIVO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', '" + Me.txt_numero_serie + "'," + CStr(var_consecutivo_serie) + ")", cnn, adOpenDynamic, adLockOptimistic
                         End If
                      End If
                   Else
                      If var_posible = True Then
                         var_posible_lectura_kanban = True
                         If var_posible_lectura_kanban = True Then
                            If var_posible_kanban = 1 Then
                               If var_kanban_es_un_kanban = "N" Or var_kanban_es_un_kanban = "" Then
                                  Set TB_RESERVAR_FUERA_KANBAN_ENT = New TB_RESERVAR_FUERA_KANBAN_ENT
                                  var_inserta = TB_RESERVAR_FUERA_KANBAN_ENT.Anadir(Me.txt_archivo, var_clave_movimiento, var_numero_folio, var_almacen_Destino, Me.txt_codigo, "", "")
                                  If var_kanban_exito = "N" Then
                                     var_posible = False
                                  End If
                               End If
                            End If
                            If var_posible = True Then
                               Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "' and inte_ent_consecutivo = " + CStr(var_consecutivo)
                               rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                               If Not rs.EOF Then
                                  var_inserta = False
                                  rsaux3.Open "UPDATE TB_TEMPORAL_ENTRADAS SET VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "', FLOA_ENT_CANTIDAD = FLOA_ENT_CANTIDAD + " + CStr(var_cantidad_leida) + " Where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_ENT_NUMERO= '" + CStr(var_numero_folio) + "' AND VCHA_ART_ARTICULO_ID  = '" + txt_codigo + "' and inte_ent_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                                  rs.Close
                               Else
                                  var_inserta = False
                                  var_costo = IIf(IsNull(var_costo), 0, var_costo)
                                  If var_costo = 0 Or var_clave_movimiento = "DT" Then
                                     rsaux4.Open "select floa_exi_costo_2005, floa_Exi_costo from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                     If Not rsaux4.EOF Then
                                        var_costo = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                     Else
                                        var_costo = 0
                                     End If
                                     rsaux4.Close
                                     If var_costo = 0 And var_almacen_Destino = "14" Then
                                        rsaux4.Open "select floa_exi_costo_2005, floa_Exi_costo from tb_existencias where vcha_alm_almacen_id = '8' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                        If Not rsaux4.EOF Then
                                           var_costo = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                        Else
                                           var_costo = 0
                                        End If
                                        rsaux4.Close
                                     End If
                                     If var_costo = 0 And var_almacen_Destino = "11" Then
                                        rsaux4.Open "select floa_exi_costo_2005, floa_Exi_costo from tb_existencias where vcha_alm_almacen_id = '8' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                        If Not rsaux4.EOF Then
                                           var_costo = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                        Else
                                           var_costo = 0
                                        End If
                                        rsaux4.Close
                                     End If
                                     If var_costo = 0 Then
                                        rsaux4.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                        If Not rsaux4.EOF Then
                                           var_costo = IIf(IsNull(rsaux4!mone_Art_costo_estandar), 0, rsaux4!mone_Art_costo_estandar)
                                        End If
                                        rsaux4.Close
                                     End If
                                  End If
                                  If var_empresa = "18" And var_proveedor = "2458" Then
                                     rsaux3.Open "INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN, INTE_ENT_CONSECUTIVO, INTE_ENT_A?O, P_RC_LINEA_ID, P_RC_NUMERO_LINEA) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo_tela) + ", " + CStr(var_precio) + ", 0,'" + var_almacen_origen + "', " + CStr(var_consecutivo) + ", " + CStr(var_a?o) + "," + CStr(var_P_RC_LINEA_ID) + "," + CStr(var_P_RC_NUMERO_LINEA) + ")", cnn, adOpenDynamic, adLockOptimistic
                                  Else
                                     If var_empresa = "06" And var_proveedor = "2458" Then
                                        rsaux3.Open "INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN, INTE_ENT_CONSECUTIVO, INTE_ENT_A?O, P_RC_LINEA_ID, P_RC_NUMERO_LINEA) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo_tela) + ", " + CStr(var_precio) + ", 0,'" + var_almacen_origen + "', " + CStr(var_consecutivo) + ", " + CStr(var_a?o) + "," + CStr(var_P_RC_LINEA_ID) + "," + CStr(var_P_RC_NUMERO_LINEA) + ")", cnn, adOpenDynamic, adLockOptimistic
                                     Else
                                        If var_precio = "" Then
                                           var_precio = 0
                                        End If
          
                                        rsaux3.Open "INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN, INTE_ENT_CONSECUTIVO, INTE_ENT_A?O, P_RC_LINEA_ID, P_RC_NUMERO_LINEA) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo + var_costo_tela) + ", " + CStr(var_precio) + ", 0,'" + var_almacen_origen + "', " + CStr(var_consecutivo) + ", " + CStr(var_a?o) + "," + CStr(var_P_RC_LINEA_ID) + "," + CStr(var_P_RC_NUMERO_LINEA) + ")", cnn, adOpenDynamic, adLockOptimistic
                                     End If
                                  End If
                                  rs.Close
                               End If
                            Else
                               lv_entradas.selectedItem.Selected = True
                               lv_entradas.selectedItem.EnsureVisible
                               lv_entradas.selectedItem.SubItems(3) = Format(lv_entradas.selectedItem.SubItems(3) - var_cantidad_leida, "###,###,##0.00")
                               lv_entradas.selectedItem.SubItems(4) = Format(lv_entradas.selectedItem.SubItems(4) - var_cantidad_leida, "###,###,##0.00")
                               lv_entradas.selectedItem.SubItems(5) = Format(lv_entradas.selectedItem.SubItems(5) + var_cantidad_leida, "###,###,##0.00")
                               var_consecutivo = lv_entradas.selectedItem.SubItems(9)
                               var_costo = lv_entradas.selectedItem.SubItems(6)
                               var_precio = lv_entradas.selectedItem.SubItems(7)
                               var_cantidad = lv_entradas.selectedItem.SubItems(4)
                               lbl_recibidos = Format(CDbl(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                               var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                               var_actualiza = TB_ARCH_COMPARACION_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), var_tipo_proveedor, var_origen, txt_codigo, 0 - var_cantidad_leida, txt_archivo, var_consecutivo)
                               var_renglon = var_n
                               frmmensaje.lbl_mensaje = var_kanban_mensaje
                               frmmensaje.Show 1
                            End If
                            
                         End If 'kanban
                         If var_numero_serie = 1 Then
                            Cadena = "select MAX(INTE_EXI_CONSECUTIVO) from TB_EXISTENCIAS_SERIES WHERE VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'"
                            var_consecutivo_serie = 0
                            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                            If Not rs.EOF Then
                               var_consecutivo_serie = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                            Else
                               var_consecutivo_serie = 1
                            End If
                            rs.Close
                            var_consecutivo_serie_str = CStr(var_consecutivo_serie)
                            If Len(var_consecutivo_serie_str) = 1 Then
                               var_consecutivo_serie_str = "000" + Trim(var_consecutivo_serie_str)
                            Else
                               If Len(var_consecutivo_serie_str) = 2 Then
                                  var_consecutivo_serie_str = "00" + Trim(var_consecutivo_serie_str)
                               Else
                                 If Len(var_consecutivo_serie_str) = 3 Then
                                    var_consecutivo_serie_str = "0" + Trim(var_consecutivo_serie_str)
                                 Else
                                     If Len(var_consecutivo_serie_str) = 4 Then
                                        var_consecutivo_serie_str = Trim(var_consecutivo_serie_str)
                                     Else
                                     End If
                                 End If
                              End If
                            End If
                            Me.txt_numero_serie = Me.txt_codigo + var_consecutivo_serie_str
                            rsaux.Open "insert into tb_exiStencias_series (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, VCHA_ART_ARTICULO_ID, VCHA_ART_NUMERO_SERIE, INTE_EXI_CONSECUTIVO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', '" + Me.txt_numero_serie + "'," + CStr(var_consecutivo_serie) + ")", cnn, adOpenDynamic, adLockOptimistic
                            rsaux9.Open "select substring(vcha_art_nombre_espa?ol,1,23) AS descripcion, substring(vcha_art_nombre_espa?ol,24,46) as descripcion_2 from tb_articulos where vcha_art_Articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                            If Not rsaux9.EOF Then
                               var_descripcion_etiqueta = IIf(IsNull(rsaux9!descripcion), "", rsaux9!descripcion)
                               var_descripcion_etiqueta2 = IIf(IsNull(rsaux9!descripcion_2), "", rsaux9!descripcion_2)
                            End If
                            rsaux9.Close
                            Open (App.Path & "\etiqueta.bat") For Output As #2
                            Print #2, "copy " + App.Path + "\etiqueta.txt lpt1"
                            Open (App.Path & "\etiqueta.txt") For Output As #1
                            Close #2
                            Print #1, "US"
                            Print #1, "q392"
                            Print #1, "Q256,24+0"
                            Print #1, "S2"
                            Print #1, "D8"
                            Print #1, "ZT"
                            Print #1, "TTh: m"
                            Print #1, "TDy2.mn.dd"
                            Print #1, "A39,20,0,3,1,1,N,""" + var_descripcion_etiqueta + """"
                            Print #1, "A39,40,0,3,1,1,N,""" + var_descripcion_etiqueta2 + """"
                            Print #1, "B39,90,0,3,2,4,101,B,""" + Trim(Me.txt_numero_serie) + """"
                            Print #1, "P1"
                            Close #1
                            x = Shell(App.Path & "\etiqueta.bat", vbHide)
                         End If
                      End If
                   End If
                   bandera_suma = False
                   var_renglon = lv_entradas.selectedItem.Index
                   Call ilumina_grid
                End If
                If var_n > 11 Then
                   lv_entradas.ColumnHeaders(2).Width = 4700.01
                Else
                   lv_entradas.ColumnHeaders(2).Width = 4930.01
                End If
                If var_cajas = True Then
                   If var_peso_correcto = True Then
                      var_costo_tela = 0
                      Me.txt_codigo.SetFocus
                   Else
                      txt_codigo = var_codigo_barras_caja
                      Me.txt_codigo_caja.SetFocus
                   End If
                Else
                   var_costo_tela = 0
                    Me.txt_codigo = ""
                End If
             End If
         Next var_jj
         If Me.lv_entradas.ListItems.Count > 0 Then
            MsgBox "Se a terminado de cargar el movimiento", vbOKOnly, "ATENCION"
         End If
      End If
      
      
'''' inicio de impresion
      
      
            
            
            
      Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
      Set TB_ENTRADAS_I = New TB_ENTRADAS_I
      Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
      Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
      Set TB_ENTRADAS_VISTAS_I = New TB_ENTRADAS_VISTAS_I
      'On Error GoTo salir:
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
      If var_numero_folio > 0 Then
         'var_estatus_movimiento = ""
         If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
            If Trim(var_reporte_imprimir) <> "" Then
               Set reporte = appl.OpenReport(App.Path + "\" + Trim(var_reporte_imprimir) + ".rpt")
               frmvistasprevias.cr.ReportSource = reporte
               reporte.RecordSelectionFormula = "{VW_ENTRADAS_COMPARACION.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_ENTRADAS_COMPARACION.VCHA_EMO_ALMACEN_DESTINO} = '" + var_almacen_Destino + "' AND {VW_ENTRADAS_COMPARACION.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and {VW_ENTRADAS_COMPARACION.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' and {VW_ENTRADAS_COMPARACION.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "'"
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte de Movimientos"
               frmvistasprevias.Show 1
               Set reporte = Nothing
               rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
               rsaux4.Open "select * from TB_EXISTENCIAS_SERIES WHERE vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_eMO_numero = " + Str(var_numero_folio)
               If Not rsaux4.EOF Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_entradas_numero_serie.rpt")
                  frmvistasprevias.cr.ReportSource = reporte
                  reporte.RecordSelectionFormula = "{VW_EXISTENCIAS_SERIES.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and {VW_EXISTENCIAS_SERIES.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' AND {VW_EXISTENCIAS_SERIES.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_Destino + "' and {VW_EXISTENCIAS_SERIES.vcha_mov_movimiento_id} = '" + var_clave_movimiento + "' and {VW_EXISTENCIAS_SERIES.INTE_EMO_NUMERO} = " + CStr(var_numero_folio)
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Movimientos"
                  frmvistasprevias.Show
                  Set reporte = Nothing
               End If
               rsaux4.Close
      
               var_m = 0
               If var_empresa = "18" Then
                  If var_clave_movimiento = "DT" Then
                     rsaux9.Open "select vcha_age_agente_id from tb_clientes where vcha_cli_clave_id = '" + var_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux9.EOF Then
                        var_clave_agente_correo = IIf(IsNull(rsaux9!VCHA_AGE_AGENTE_ID), "", rsaux9!VCHA_AGE_AGENTE_ID)
                     End If
                     rsaux9.Close
                     rsaux9.Open "select * from tb_agentes where vcha_age_agente_id =  '" + var_clave_agente_correo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux9.EOF Then
                        var_correo_electronico = IIf(IsNull(rsaux9!VCHA_AGE_EMAIL), "", rsaux9!VCHA_AGE_EMAIL)
                        var_nombre_agente_correo = IIf(IsNull(rsaux9!VCHA_AGE_NOMBRE), "", rsaux9!VCHA_AGE_NOMBRE)
                     End If
                     rsaux9.Close
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
                        MAPIMessages1.MsgSubject = "Devoluci?n de tienda " + Me.txt_archivo
                        MAPIMessages1.MsgNoteText = "Se anexa informaci?n de la devoluci?n  " + Me.txt_archivo
                        var_Archivo = App.Path & "\Devolucion_" + Trim(Me.txt_archivo) + ".txt"
                        Open (App.Path & "\Devolucion_" + Trim(Me.txt_archivo) + ".txt") For Output As #1
                        Print #1, "Se genero la devoluci?n " + Trim(Me.txt_archivo) + " con los siguientes datos"
                        Print #1, ""
                        Print #1, "Cliente: " + Trim(var_nombre_agente_correo)
                        Print #1, ""
                        rsaux9.Open "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_Alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                        Print #1, "CODIGO       DESCRIPCION                                      CANTIDAD"
                        Print #1, "======================================================================"
                        var_total_correo = 0
                        While Not rsaux9.EOF
                              var_linea = ""
                              var_cantidad_correo = Format(IIf(IsNull(rsaux9!floa_ent_cantidaD), 0, rsaux9!floa_ent_cantidaD), "###,###,##0.00")
                              var_total_correo = var_total_correo + IIf(IsNull(rsaux9!floa_ent_cantidaD), 0, rsaux9!floa_ent_cantidaD)
                              rsaux8.Open "select vcha_Art_nombre_espa?ol from tb_articulos where vcha_Art_articulo_id = '" + rsaux9!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux8.EOF Then
                                 var_nombre_articulo_correo = Mid(IIf(IsNull(rsaux8!vcha_Art_nombre_espa?ol), "", rsaux8!vcha_Art_nombre_espa?ol), 1, 40)
                              End If
                              rsaux8.Close
                              For var_j = Len(var_cantidad_correo) To 14
                                  var_cantidad_correo = " " + var_cantidad_correo
                              Next var_j
                              For var_j = Len(var_nombre_articulo_correo) To 40
                                  var_nombre_articulo_correo = var_nombre_articulo_correo + " "
                              Next var_j
                              var_linea = rsaux9!VCHA_ART_ARTICULO_ID + " " + var_nombre_articulo_correo + " " + var_cantidad_correo
                              Print #1, var_linea
                              rsaux9.MoveNext
                        Wend
                        Print #1, "======================================================================"
                        rsaux9.Close
                        var_total_correo_str = Format(var_total_correo, "###,###,##0.00")
                        For var_j = Len(var_total_correo_str) To 14
                            var_total_correo_str = " " + var_total_correo_str
                        Next var_j
                        var_linea = "                                       POR UN TOTAL DE:" + var_total_correo_str
                        Print #1, var_linea
                        Print #1, ""
                        Close #1
                        MAPIMessages1.AttachmentPathName = var_Archivo
                        MAPIMessages1.Send True
                        If MAPISession1.SessionID > 0 Then
                           MAPISession1.SignOff
                        End If
                     Else
                     End If
                  End If
               End If
               '''' fin del correo
            Else
               MsgBox "El movimiento no tiene un reporte asociado", vbOKOnly, "ATENCION"
            End If
         Else
            'var_si = MsgBox("?Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
            var_si = 1
            If var_si = 1 Then
               x = 0
               Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
               rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
               VAR_ZZ = rs.RecordCount
               'MsgBox CStr(var_zz)
               If var_clave_movimiento = "EC" Then
                  var_posible_entrada = True
                  If x = 0 Then
                     var_si_paso = Obtener_Identificador_Recepcion(cnnoracle, 0)
                  End If
                  
        
                  If var_clave_movimiento = "EC" Then
                     While Not rs.EOF
                           If x = 0 Then
                              If IsNumeric(txt_archivo) Then
                                 VAR_PASO = Insertar_Recepcion(CDbl(var_txt_archivo), CDbl(rs!p_rc_numero_linea), CDbl(rs!floa_ent_cantidaD), CDbl(rs!P_RC_LINEA_ID), CDbl(var_numero_folio), CDbl(var_unidad_OC), "0", Me.txt_factura)
                              End If
                           End If
                           rs.MoveNext
                     Wend
                  End If
               Else
                  var_posible_entrada = True
               End If
               If rs.RecordCount > 0 Then
                  rs.MoveFirst
               End If
   'GoTo   salir:r
               If var_posible_entrada = True Then
                  var_posible = 1
                  cnn.BeginTrans
                  If Not rs.EOF Then
                     var_inserta = False
                     x = 1
                     If x = 1 Then
                        'If var_posible_kanban = 1 Then
                        If var_clave_movimiento = "EP" And var_empresa = "02" Then
                  
                           var_cadena = "SELECT dbo.TB_TEMPORAL_ENTRADAS.VCHA_EMP_EMPRESA_ID, dbo.TB_TEMPORAL_ENTRADAS.VCHA_UOR_UNIDAD_ID, dbo.TB_TEMPORAL_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_TEMPORAL_ENTRADAS.INTE_ENT_NUMERO, dbo.TB_TEMPORAL_ENTRADAS.VCHA_ART_ARTICULO_ID , dbo.TB_Articulos.INTE_ART_KANBAN FROM dbo.TB_TEMPORAL_ENTRADAS INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_TEMPORAL_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_ARTICULOS.INTE_ART_KANBAN = 1) AND (dbo.TB_TEMPORAL_ENTRADAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND  (dbo.TB_TEMPORAL_ENTRADAS.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') AND (dbo.TB_TEMPORAL_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "') AND  (dbo.TB_TEMPORAL_ENTRADAS.INTE_ENT_NUMERO = " + CStr(var_numero_folio) + ")"
                           rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux10.EOF Then
                              Set TB_PROC_KANBANS_EN_MOV_ENTRADA = New TB_PROC_KANBANS_EN_MOV_ENTRADA
                              var_inserta = TB_PROC_KANBANS_EN_MOV_ENTRADA.Anadir(var_almacen_Destino, var_clave_movimiento, CDbl(Me.txt_folio), "", "")
                              If var_kanban_exito = "N" Then
                                 var_posible = 0
                              End If
                           Else
                              var_posible = 1
                           End If
                        Else
                           var_posible = 1
                        End If
                     'End If
                     'cnn.RollbackTrans
                     End If
                     
                     If var_posible = True Then
                        var_inserta = TB_ENTRADAS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, var_tipo_documento, var_almacen_origen, var_folio_enviado)
                        rsaux4.Open "UPDATE TB_aRCHIVO_COMPARACION SET VCHA_COM_TIPO_LECTURA = 'F' WHERE VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND VCHA_COM_REFERENCIA = '" + Me.txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
                     End If
                  End If
                  rs.Close
                  If var_tipo_documento = "V" Then
                     var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, "", Now, var_tipo_Cambio)
                  Else
                     var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, var_tipo_Cambio)
                  End If
                  var_estatus_movimiento = "I"
                  If var_tipo_documento = "V" Then
                     var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, "I", Now, var_tipo_Cambio)
                  Else
                     var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, var_tipo_Cambio)
                  End If
                  cnn.CommitTrans
                  rsaux2.Open "select * from tb_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  VAR_ZZ = CStr(rsaux2.RecordCount)
                  'MsgBox (var_zz)
                  If Not rsaux2.EOF Then
                     If var_posible = True Then
                        If Trim(var_reporte_imprimir) <> "" Then
                           Set reporte = appl.OpenReport(App.Path + "\" + Trim(var_reporte_imprimir) + ".rpt")
                           reporte.RecordSelectionFormula = "{VW_ENTRADAS_COMPARACION.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_ENTRADAS_COMPARACION.VCHA_EMO_ALMACEN_DESTINO} = '" + var_almacen_Destino + "' AND {VW_ENTRADAS_COMPARACION.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and {VW_ENTRADAS_COMPARACION.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' and {VW_ENTRADAS_COMPARACION.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "'"
                           frmvistasprevias.cr.ReportSource = reporte
                           For ntablas = 1 To reporte.Database.Tables.Count
                               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                           Next ntablas
                           frmvistasprevias.cr.ViewReport
                           frmvistasprevias.Caption = "Reporte de Movimientos"
                           frmvistasprevias.Show 1
                           Set reporte = Nothing
                           rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                       
                           rsaux4.Open "select * from TB_EXISTENCIAS_SERIES WHERE vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_eMO_numero = " + Str(var_numero_folio)
                           If Not rsaux4.EOF Then
                              Set reporte = appl.OpenReport(App.Path + "\rep_entradas_numero_serie.rpt")
                              frmvistasprevias.cr.ReportSource = reporte
                              reporte.RecordSelectionFormula = "{VW_EXISTENCIAS_SERIES.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and {VW_EXISTENCIAS_SERIES.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' AND {VW_EXISTENCIAS_SERIES.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_Destino + "' and {VW_EXISTENCIAS_SERIES.vcha_mov_movimiento_id} = '" + var_clave_movimiento + "' and {VW_EXISTENCIAS_SERIES.INTE_EMO_NUMERO} = " + CStr(var_numero_folio)
                              For ntablas = 1 To reporte.Database.Tables.Count
                                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                              Next ntablas
                              frmvistasprevias.cr.ViewReport
                              frmvistasprevias.Caption = "Reporte de Movimientos"
                              frmvistasprevias.Show
                              Set reporte = Nothing
                           End If
                           rsaux4.Close
                           var_m = 0
                           If var_empresa = "18" Then
                              If var_clave_movimiento = "DT" Then
                                 rsaux9.Open "select vcha_age_agente_id from tb_clientes where vcha_cli_clave_id = '" + var_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux9.EOF Then
                                    var_clave_agente_correo = IIf(IsNull(rsaux9!VCHA_AGE_AGENTE_ID), "", rsaux9!VCHA_AGE_AGENTE_ID)
                                 End If
                                 rsaux9.Close
                                 rsaux9.Open "select * from tb_agentes where vcha_age_agente_id =  '" + var_clave_agente_correo + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux9.EOF Then
                                    var_correo_electronico = IIf(IsNull(rsaux9!VCHA_AGE_EMAIL), "", rsaux9!VCHA_AGE_EMAIL)
                                    var_nombre_agente_correo = IIf(IsNull(rsaux9!VCHA_AGE_NOMBRE), "", rsaux9!VCHA_AGE_NOMBRE)
                                 End If
                                 rsaux9.Close
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
                                    MAPIMessages1.MsgSubject = "Devoluci?n de tienda " + Me.txt_archivo
                                    MAPIMessages1.MsgNoteText = "Se anexa informaci?n de la devoluci?n  " + Me.txt_archivo
                                    var_Archivo = App.Path & "\Devolucion_" + Trim(Me.txt_archivo) + ".txt"
                                    Open (App.Path & "\Devolucion_" + Trim(Me.txt_archivo) + ".txt") For Output As #1
                                    Print #1, "Se genero la devoluci?n " + Trim(Me.txt_archivo) + " con los siguientes datos"
                                    Print #1, ""
                                    Print #1, "Cliente: " + Trim(var_nombre_agente_correo)
                                    Print #1, ""
                                    rsaux9.Open "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_Alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                                    Print #1, "CODIGO       DESCRIPCION                                      CANTIDAD"
                                    Print #1, "======================================================================"
                                    var_total_correo = 0
                                    While Not rsaux9.EOF
                                          var_linea = ""
                                          var_cantidad_correo = Format(IIf(IsNull(rsaux9!floa_ent_cantidaD), 0, rsaux9!floa_ent_cantidaD), "###,###,##0.00")
                                          var_total_correo = var_total_correo + IIf(IsNull(rsaux9!floa_ent_cantidaD), 0, rsaux9!floa_ent_cantidaD)
                                          rsaux8.Open "select vcha_Art_nombre_espa?ol from tb_articulos where vcha_Art_articulo_id = '" + rsaux9!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux8.EOF Then
                                             var_nombre_articulo_correo = Mid(IIf(IsNull(rsaux8!vcha_Art_nombre_espa?ol), "", rsaux8!vcha_Art_nombre_espa?ol), 1, 40)
                                          End If
                                          rsaux8.Close
                                          For var_j = Len(var_cantidad_correo) To 14
                                              var_cantidad_correo = " " + var_cantidad_correo
                                          Next var_j
                                          For var_j = Len(var_nombre_articulo_correo) To 40
                                              var_nombre_articulo_correo = var_nombre_articulo_correo + " "
                                          Next var_j
                                          var_linea = rsaux9!VCHA_ART_ARTICULO_ID + " " + var_nombre_articulo_correo + " " + var_cantidad_correo
                                          Print #1, var_linea
                                          rsaux9.MoveNext
                                    Wend
                                    Print #1, "======================================================================"
                                    rsaux9.Close
                                    var_total_correo_str = Format(var_total_correo, "###,###,##0.00")
                                    For var_j = Len(var_total_correo_str) To 14
                                        var_total_correo_str = " " + var_total_correo_str
                                    Next var_j
                                    var_linea = "                                       POR UN TOTAL DE:" + var_total_correo_str
                                    Print #1, var_linea
                                    Print #1, ""
                                    Close #1
                                    MAPIMessages1.AttachmentPathName = var_Archivo
                                    MAPIMessages1.Send True
                                    If MAPISession1.SessionID > 0 Then
                                       MAPISession1.SignOff
                                    End If
                                 Else
                                 End If ' fin del correo
                              End If ' fin del movimiento dt
                           End If 'fin empresa 18
                        Else
                           MsgBox "El movimiento no tiene un reporte asociado", vbOKOnly, "ATENCION"
                        End If
                     Else
                        MsgBox "No se pudo cerrar el movimiento kanban", vbOKOnly, "ATENCION"
                     End If 'vaR_posible
                  Else
                     If var_empresa <> "18" Then
                        MsgBox "El movimiento no a afectado el inventario intentelo nuevamente o consulte a sistemas", vbOKOnly, "ATENCION"
                        rsaux4.Open "update tb_encabezado_movimientos set char_emo_estatus = '' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                        var_estatus_movimiento = ""
                        Me.txt_codigo.Enabled = True
                     Else
                        MsgBox "El movimiento a sido cerrado", vbOKOnly, "ATENCION"
                     End If
                  End If
                  rsaux2.Close
               Else
                  MsgBox "A surgido un problema al afectar ORACLE vuelva a intentar de nuevo", vbOKOnly, "ATENCION"
               End If
               txt_codigo.Enabled = False
               txt_foco.Enabled = False
            End If
         End If
      Else
         MsgBox "No se a seleccionado ning?n movimiento", vbOKOnly, "ATENCION"
      End If
''' fin de impresion
   End If
   Exit Sub
salir:
   If Err.Number = -2147217871 Then
      Resume
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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 1 And KeyCode = 116 Then
      frmexisten_rapidas.Show
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
   If Shift = 4 And KeyCode = 67 Then
      cmd_cancelar_Click
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If var_ventana = 0 Then
         If Me.txt_folio <> "" Then
            var_si = MsgBox("?Deseas cerrar el movimiento?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Unload Me
            End If
         Else
            Unload Me
         End If
      Else
         If Me.frm_articulo_caja.Visible = True Then
            Me.txt_codigo.Enabled = True
            var_costo_tela = 0
            Me.txt_codigo.SetFocus
            Me.frm_articulo_caja.Visible = False
         End If
         Me.frm_busqueda.Visible = False
         Me.frm_eliminar.Visible = False
         Me.frm_peso_caja.Visible = False
         Me.frmnumero_serie.Visible = False
         var_ventana = 0
      End If
   End If
End Sub



Private Sub Form_Load()
   If var_unidad_organizacional = "19" Then
      Me.cmd_crossdocking.Visible = True
   Else
      If var_empresa = "31" And (var_clave_movimiento = "EP" Or var_clave_movimiento = "EI") Then
         Me.cmd_crossdocking.Visible = True
      Else
         If var_empresa = "17" Then
            Me.cmd_crossdocking.Visible = True
            
         Else
            Me.cmd_crossdocking.Visible = False
         End If
      End If
   End If
   'If var_empresa = 16 Then
   '   Me.cmd_pasar_todos.Visible = True
   'Else
   '   Me.cmd_pasar_todos.Visible = False
   'End If
   var_codigo_seleccionado = ""
   Me.cmd_tipo_lectura.Visible = False
   var_costo_tela = 0
   var_numero_serie = 0
   frmnumero_serie.Visible = False
   lbl_cancelado = ""
   var_cajas = False
   frm_articulo_caja.Visible = False
   frm_peso_caja.Visible = False
   var_cadena_seguridad = ""
   Top = 0
   Left = 0
   var_ventana = 0
   var_cantidad_leida = 1#
   var_estatus_movimiento = ""
   var_almacen_Destino = ""
   var_almacen_origen = ""
   var_proveedor = ""
   var_factura = ""
   frm_eliminar.Visible = False
   var_modifica = False
   txt_Cantidad.Visible = False
   lbl_Cantidad.Visible = False
   txt_codigo.Enabled = False
   var_primera_vez = True
   var_numero_folio = 0
   frm_busqueda.Visible = False
   Set var_tabla = CreateObject("ADODB.connection")
   var_suma_cantidad_enviada = 0
   var_suma_cantidad_recibida = 0
   rs.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_tolerancia_peso_caja = IIf(IsNull(rs!floa_pri_tolerancia_peso), 0, rs!floa_pri_tolerancia_peso)
   Else
      var_tolerancia_peso_caja = 0
   End If
   rs.Close
   If var_unidad_organizacional = "28" And (var_clave_usuario_global = "U0000000212" Or var_clave_usuario_global = "U0000000211") Then
      
      Me.cmd_pasar_todos.Visible = True
   Else
      If var_empresa = "17" Then
         Me.cmd_pasar_todos.Visible = True
      End If
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
   
   If var_solo_lectura = False Then
      Set TB_BLOQUEOS = New TB_BLOQUEOS
      var_global_bloqueado = 0
      ok = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, txt_archivo, Now, "", "")
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
   End If
   Call activa_forma(var_activa_forma_entradas)
End Sub


Private Sub lv_entradas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_entradas, ColumnHeader)
End Sub

Private Sub lv_entradas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         MsgBox "Imposible modificar el movimiento", vbOKOnly, "ATENCION"
      Else
         If var_empresa = "16" Then
            var_ventana = 2
            frm_eliminar.Visible = True
            Me.txt_cantidad_eliminar = Me.lv_entradas.selectedItem.SubItems(4)
            txt_cantidad_eliminar.SetFocus
         Else
            var_ventana = 2
            frm_eliminar.Visible = True
            Me.txt_cantidad_eliminar = ""
            txt_cantidad_eliminar.SetFocus
         End If
      End If
   End If
End Sub


Private Sub lv_lista_LostFocus()
   var_clave_almacen_seleccionado = lv_lista.selectedItem
End Sub

Private Sub txt_archivo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If var_clave_movimiento = "ETA" Or var_clave_movimiento = "EI" Or var_clave_movimiento = "EP" Then
         rs.Open "select distinct VCHA_PLA_PLANTA_ID from tb_plantas where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
         var_i = 0
         While Not rs.EOF
               var_i = var_i + 1
               var_planta_transito_global = IIf(IsNull(rs!vcha_pla_planta_id), "", rs!vcha_pla_planta_id)
               rs.MoveNext
         Wend
         rs.Close
         If var_i > 1 Then
            frmplantas_transito.Show 1
            Me.txt_archivo = var_nota_traspasos_transito
            Me.lbl_transito = var_nota_traspasos_transito
         Else
            Select Case var_clave_movimiento
            Case "EP"
                frmnotas_traspasos_plantas.var_str_encabezado_forma = "Entradas De Envios De Produccion"
                rs.Open "Select numb_tra_consecutivo, VCHA_TRA_NOTA_ENVIO , " & _
                            "VCHA_TRA_ALMACEN_ORIGEN, VCHA_ART_ARTICULO_ORIGEN, " & _
                            "NUMB_TRA_CANTIDAD_ENVIADA, VCHA_TRA_REFERENCIA1, NVL(VCHA_TRA_CONTENEDOR_ID,' ') VCHA_TRA_CONTENEDOR_ID " & _
                        "from tb_transito " & _
                        "where VCHA_TRA_NOTA_ENVIO = '" & UCase(var_nota_traspasos_transito) & "' ", _
                    cnnoracle, _
                    adOpenDynamic, _
                    adLockOptimistic
            
                For fila = 1 To rs.RecordCount
                    rsaux.Open "Update tb_archivo_comparacion " & _
                                "set inte_com_consecutivo =" & rs("numb_tra_consecutivo").Value & ", " & _
                                    "vcha_com_referencia_almacen = '" & rs("VCHA_TRA_ALMACEN_ORIGEN").Value & "'" & _
                                "where vcha_art_articulo_id ='" & rs("VCHA_ART_ARTICULO_ORIGEN").Value & "' " & _
                                "and VCHA_COM_REFERENCIA ='" & var_nota_traspasos_transito & "' " & _
                                "and inte_com_lote ='" & rs("VCHA_TRA_REFERENCIA1").Value & "' and (vcha_com_caja = '" & rs("VCHA_TRA_CONTENEDOR_ID").Value & "' or vcha_com_caja = '' )", _
                            cnn, _
                            adOpenDynamic, _
                            adLockOptimistic
                    rs.MoveNext
                Next
                rs.Close
            Case Else
                frmnotas_traspasos_plantas.var_str_encabezado_forma = "Entradas Por traspaso Entre Plantas"
            End Select
            
            frmnotas_traspasos_plantas.Show 1
            Me.txt_archivo = var_nota_traspasos_transito
            Me.lbl_transito = var_nota_traspasos_transito
         End If
         
      End If
   End If
End Sub

Private Sub txt_archivo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   txt_archivo = UCase(txt_archivo)
   If KeyAscii = 13 Then
      Call ejecuta
   End If
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
Dim var_busqueda_folio As Integer
Dim var_busqueda_numero As Integer
Dim var_busqueda_referencia As String
Dim var_almacen_tem_agente As String
Dim var_movimiento_bloqueado As Integer
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
If Len(Trim(txt_busqueda_folio)) > 0 Then
   Me.cmd_tipo_lectura.Visible = False
   Select Case KeyAscii
   Case 48 To 57, 52, 8, 46
   Case 13
      txt_archivo = UCase(txt_archivo)
      Set TB_BLOQUEOS = New TB_BLOQUEOS
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select * from tb_movimientos where vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
      var_global_relectura = IIf(IsNull(rs!INTE_MOV_RELECTURA), 0, rs!INTE_MOV_RELECTURA)
      var_global_aceptar_demas = 0
      var_global_aceptar_demas = IIf(IsNull(rs!INTE_MOV_ACEPTAR_MAS), 0, rs!INTE_MOV_ACEPTAR_MAS)
      rs.Close
      If var_numero_folio = CDbl(txt_busqueda_folio) Then
         rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
      End If
      cnn.CommandTimeout = 360
      rs.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + txt_busqueda_folio + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.lbl_transito = IIf(IsNull(rs!vcha_emo_referencia_transito), "", rs!vcha_emo_referencia_transito)
         If var_numero_folio > 0 Then
            rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
         End If
         var_movimiento_bloqueado = IIf(IsNull(rs!INTE_EMO_BLOQUEADO), 0, rs!INTE_EMO_BLOQUEADO)
         If var_movimiento_bloqueado = 0 Then
            var_primera_vez = False
            var_estatus_movimiento = rs!char_Emo_estatus
            var_busqueda_referencia = rs!vcha_emo_Referencia
            var_numero_folio = Val(txt_busqueda_folio)
            var_fecha_movimiento = rs!dtim_emo_fecha
            var_almacen_destino_tem = rs!VCHA_ALM_ALMACEN_ID
            var_almacen_destino_tem_agente = IIf(IsNull(rs!VCHA_EMO_ALMACEN_DESTINO), "", rs!VCHA_EMO_ALMACEN_DESTINO)
            var_pedimento = IIf(IsNull(rs!VCHA_EMO_PEDIMENTO), "", rs!VCHA_EMO_PEDIMENTO)
            var_posible = 1
            If var_tipo_permiso = 1 Then
               If var_tipo_documento = "V" Then
                  rsaux.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_per_almacen_1 = '" + var_almacen_destino_tem_agente + "'", cnn, adOpenDynamic, adLockBatchOptimistic
                  If rsaux.EOF Then
                     var_posible = 0
                  End If
                  rsaux.Close
               Else
                  rsaux.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_per_almacen_1 = '" + var_almacen_destino_tem + "'", cnn, adOpenDynamic, adLockBatchOptimistic
                  If rsaux.EOF Then
                     var_posible = 0
                  End If
                  rsaux.Close
               End If
            Else
               var_posible = 1
            End If
            
            If var_posible = 1 Then
               rs.Close
               rs.Open "select * from VW_ENTRADAS_COMPARACION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + txt_busqueda_folio, cnn, adOpenDynamic, adLockOptimistic
               txt_folio = txt_busqueda_folio
               var_unidad_OC = CStr(IIf(IsNull(rs!P_RC_ORD_ID), 0, rs!P_RC_ORD_ID))
               If IsNull(rs!vcha_com_Referencia) Then
                  txt_archivo = ""
               Else
                  If var_clave_movimiento = "EC" Then
                     'MsgBox Mid(IIf(IsNull(rs!vcha_com_Referencia), "", Trim(rs!vcha_com_Referencia)), 1, 3)
                     If Mid(IIf(IsNull(rs!vcha_com_Referencia), "", Trim(rs!vcha_com_Referencia)), 1, 3) = "545" Then
                        var_txt_archivo = Mid(IIf(IsNull(rs!vcha_com_Referencia), "", Trim(rs!vcha_com_Referencia)), 4)
                     Else
                        var_txt_archivo = Mid(IIf(IsNull(rs!vcha_com_Referencia), "", Trim(rs!vcha_com_Referencia)), 3)
                     End If
                  End If
                  'MsgBox var_txt_archivo
                  txt_archivo = IIf(IsNull(rs!vcha_com_Referencia), "", Trim(rs!vcha_com_Referencia))
                  txt_archivo.Enabled = False
               End If
               If IsNull(rs!VCHA_ALM_ALMACEN_ID) Then
                  var_almacen_Destino = ""
               Else
                  var_almacen_Destino = Trim(rs!VCHA_ALM_ALMACEN_ID)
               End If
               If IsNull(rs!VCHA_COM_PROVEEDOR) Then
                  var_origen = ""
               Else
                  var_origen = Trim(rs!VCHA_COM_PROVEEDOR)
               End If
               
               If IsNull(rs!vcha_emo_factura) Then
                  var_factura = ""
               Else
                  var_factura = Trim(rs!vcha_emo_factura)
               End If
               txt_factura = var_factura
               Me.txt_pedimento = var_pedimento
               If IsNull(rs!CHAR_COM_TIPO_PROVEEDOR) Then
                  var_tipo_proveedor = ""
               Else
                  var_tipo_proveedor = rs!CHAR_COM_TIPO_PROVEEDOR
               End If
               If IsNull(rs!VCHA_COM_TRANSPORTO) Then
                  var_transporto = ""
                  txt_transporto = Trim(var_transporto)
               Else
                  var_transporto = rs!VCHA_COM_TRANSPORTO
                  txt_transporto = var_transporto
               End If
               
               var_entrada_calidad = False
               If var_causa_devolucion = True Then
                  rsaux3.Open "select * from tb_almacenes where inte_alm_costeo =  1 and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     var_entrada_calidad = True
                     If IsNull(rsaux3!VCHA_ALM_ALMACEN_ID) Then
                        var_entrada_calidad = False
                     Else
                        var_almacen_costeo = rsaux3!VCHA_ALM_ALMACEN_ID
                     End If
                  Else
                     var_entrada_calidad = False
                  End If
                  rsaux3.Close
               End If
               If var_almacen_Destino <> "" Then
                  If rsaux.State = 1 Then
                     rsaux.Close
                  End If
                  rsaux.Open "select vcha_alm_almacen_id,vcha_alm_nombre from tb_almacenes where vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     txt_destino = rsaux(1).Value
                     var_almacen_Destino = rsaux(0).Value
                     rsaux.Close
                  Else
                     rsaux.Close
                  End If
               End If
               If var_tipo_proveedor <> "" Then
                  If var_tipo_proveedor = "U" Then
                     rsaux.Open "select vcha_uor_unidad_id,vcha_uor_nombre from tb_unidadesorganizacionales where vcha_UOR_unidad_id = '" + var_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        var_proveedor = rsaux!VCHA_UOR_UNIDAD_ID
                        txt_origen = rsaux!VCHA_UOR_NOMBRE
                     End If
                     rsaux.Close
                  End If
                  If var_tipo_proveedor = "P" Then
                     rsaux.Open "select vcha_pro_proveedor_id,vcha_pro_nombre from tb_proveedores where vcha_pro_proveedor_id = '" + var_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        var_proveedor = rsaux!VCHA_PRO_PROVEEDOR_ID
                        txt_origen = rsaux!VCHA_PRO_NOMBRE
                     End If
                     rsaux.Close
                  End If
                  If var_tipo_proveedor = "T" Then
                     rsaux.Open "select vcha_cli_clave_id,vcha_cli_nombre from tb_clientes where vcha_cli_clave_id = '" + var_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        var_proveedor = rsaux!vcha_cli_clave_id
                        txt_origen = rsaux!VCHA_CLI_NOMBRE
                     End If
                     rsaux.Close
                  End If
                  If var_tipo_proveedor = "A" Then
                     rsaux.Open "select vcha_alm_almacen_id,vcha_alm_nombre from tb_almacenes where vcha_alm_almacen_id = '" + var_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        var_proveedor = RSUAX!VCHA_ALM_ALMACEN_ID
                        txt_origen = rsaux!VCHA_ALM_NOMBRE
                     End If
                     rsaux.Close
                  End If
                  If var_tipo_proveedor = "G" Then
                     rsaux.Open "select vcha_age_agente_id,vcha_age_nombre from tb_agentes where vcha_age_agente_id = '" + var_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        var_proveedor = RSUAX!VCHA_AGE_AGENTE_ID
                        txt_origen = rsaux!VCHA_AGE_NOMBRE
                     End If
                     rsaux.Close
                  End If
               End If
               If IsNull(rs!INTE_COM_NUMERO) Then
                  var_folio_enviado = 0
               Else
                  var_folio_enviado = rs!INTE_COM_NUMERO
                  If var_tipo_documento = "V" Then
                     rsaux2.Open "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id ='" + var_unidad_organizacional + "' and inte_emo_numero = " + Str(var_folio_enviado) + " and vcha_mov_movimiento_id = 'SV'", cnn, adOpenDynamic, adLockOptimistic
                     var_almacen_origen = rsaux2!VCHA_EMO_ALMACEN_DESTINO
                     var_almacen_Destino = rsaux2!vcha_emo_almacen_origen
                     rsaux2.Close
                  End If
               End If
               rs.Close
               rs.Open "select * from VW_ENTRADAS_COMPARACION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + txt_busqueda_folio, cnn, adOpenDynamic, adLockOptimistic
               var_suma_cantidad_enviada = 0
               var_suma_cantidad_recibida = 0
               lbl_enviados.Caption = "0"
               lbl_recibidos.Caption = "0"
               lv_entradas.ListItems.Clear
               While Not rs.EOF
                  rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     Set list_item = lv_entradas.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
                         list_item.SubItems(1) = IIf(IsNull(rsaux(1).Value), "", rsaux(1).Value)
                         list_item.SubItems(2) = Format(IIf(IsNull(rs!FLOA_COM_CANTIDAD_ENVIADA), 0, rs!FLOA_COM_CANTIDAD_ENVIADA), "###,###,##0.00")
                         list_item.SubItems(3) = Format(IIf(IsNull(rs!FLOA_com_cANTIDAD_RECIBIDA), 0, rs!FLOA_com_cANTIDAD_RECIBIDA), "###,###,##0.00")
                         list_item.SubItems(4) = Format(IIf(IsNull(rs!floa_ent_cantidaD), 0, rs!floa_ent_cantidaD), "###,###,##0.00")
                         list_item.SubItems(5) = Format(list_item.SubItems(2) - list_item.SubItems(3), "###,###,##0.00")
                         list_item.SubItems(6) = IIf(IsNull(rs!FLOA_COM_COSTO), "", rs!FLOA_COM_COSTO)
                         list_item.SubItems(7) = IIf(IsNull(rsaux(2).Value), "", rsaux(2).Value)
                         list_item.SubItems(8) = IIf(IsNull(rs!INTE_COM_LOTE), 0, rs!INTE_COM_LOTE)
                         list_item.SubItems(9) = IIf(IsNull(rs!INTE_COM_CONSECUTIVO), "", rs!INTE_COM_CONSECUTIVO)
                         list_item.SubItems(10) = IIf(IsNull(rs!INTE_COM_A?O), "", rs!INTE_COM_A?O)
                         list_item.SubItems(11) = IIf(IsNull(rs!VCHA_COM_CAJA), "", rs!VCHA_COM_CAJA)
                         list_item.SubItems(12) = IIf(IsNull(rs!FLOA_COM_PESO), 0, rs!FLOA_COM_PESO)
                         list_item.SubItems(13) = IIf(IsNull(rs!P_RC_LINEA_ID), 0, rs!P_RC_LINEA_ID)
                         list_item.SubItems(14) = IIf(IsNull(rs!p_rc_numero_linea), 0, rs!p_rc_numero_linea)
                         list_item.SubItems(16) = IIf(IsNull(rs!FLOA_COM_PRECIO), 0, rs!FLOA_COM_PRECIO)
                         
                         var_suma_cantidad_enviada = var_suma_cantidad_enviada + rs!FLOA_COM_CANTIDAD_ENVIADA
                         var_suma_cantidad_recibida = var_suma_cantidad_recibida + rs!FLOA_com_cANTIDAD_RECIBIDA
                  End If
                  rsaux.Close
                  rs.MoveNext:
               Wend
               rs.Close
               'cajas
               rs.Open "select * from tb_archivo_comparacion where vcha_emp_empresa_id ='" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_com_referencia = '" + txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Len(Trim(IIf(IsNull(rs!VCHA_COM_CAJA), "", rs!VCHA_COM_CAJA))) > 0 Then
                  lbl_tipo = "C?digo de la caja:"
                  var_cajas = True
               Else
                 lbl_tipo = "C?digo del art?culo:"
                 var_cajas = False
               End If
               rs.Close
               var_n = lv_entradas.ListItems.Count
               If var_n > 11 Then
                  lv_entradas.ColumnHeaders(2).Width = 4700.01
               Else
                  lv_entradas.ColumnHeaders(2).Width = 4930.01
               End If
               var_renglon = -1
               rs.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 1 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
               Call ilumina_grid
               lbl_enviados = Format(Str(var_suma_cantidad_enviada), "###,###,##0.00")
               lbl_recibidos = Format(Str(var_suma_cantidad_recibida), "###,###,##0.00")
               txt_factura.Enabled = False
               Me.txt_pedimento.Enabled = False
               If var_solo_lectura = False Then
                  var_global_bloqueado = 0
                  ok = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, txt_archivo, Now, "", "")
               End If
               rs.Open "select * from vw_bloqueos where vcha_blo_bloqueado_por = '" + txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_solo_lectura = True
                  MsgBox "No puede modificar este movimiento ya que el la relaci?n esta siendo utilizada por el usuario: '" + Trim(rs!VCHA_USU_NOMBRE) + " " + Trim(rs!vcha_usu_apellidos) + "' en la m?quina: '" + Trim(rs!vcha_blo_maquina) + "'", vbOKOnly, "ATENCION"
               Else
                  var_solo_lectura = False
               End If
               rs.Close
               If var_estatus_movimiento = "I" Then
                  txt_codigo.Enabled = False
               Else
                  If var_solo_lectura = True Then
                     txt_codigo.Enabled = False
                  Else
                     var_global_bloqueado = 1
                     ok = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, txt_archivo, Now, var_clave_usuario_global, fun_NombrePc)
                     txt_codigo.Enabled = True
                     var_costo_tela = 0
                     txt_codigo.SetFocus
                  End If
               End If
               frm_busqueda.Visible = False
               var_ventana = 0
            Else
               rs.Close
               MsgBox "No esta autorizado para modificar este movimiento", vbOKOnly, "ATENCION"
               frm_busqueda.Visible = False
               var_ventana = 0
            End If
            If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
               If var_estatus_movimiento = "C" Then
                  Me.cmd_cancelar.Enabled = False
                  Me.cmd_imprimir.Enabled = False
                  lbl_cancelado = "MOVIMIENTO CANCELADO"
               End If
               Me.txt_codigo.Enabled = False
            Else
               Me.cmd_cancelar.Enabled = True
               Me.cmd_imprimir.Enabled = True
               Me.txt_codigo.Enabled = True
               lbl_cancelado = ""
            End If
         Else
            rs.Close
            MsgBox "El movimiento esta siendo utilizado por otro usuario", vbOKOnly, "ATENCION"
         End If
      Else
         rs.Close
         MsgBox "El n?mero de movimiento " + Trim(txt_busqueda_folio) + " no existe", vbOKOnly, "ATENCION"
         frm_busqueda.Visible = False
         var_ventana = 0
      End If
   Case 27
      frm_busqueda.Visible = False
      var_ventana = 0
   End Select
End If
If KeyAscii = 27 Then
   frm_busqueda.Visible = False
   var_ventana = 0
End If
End Sub

Private Sub txt_busqueda_folio_LostFocus()
   var_ventana = 0
   frm_busqueda.Visible = False
End Sub

Private Sub txt_cantidad_eliminar_GotFocus()
   'txt_cantidad_eliminar = ""
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   'If Not var_posible_kanban = 1 Then
   '   Select Case KeyAscii
   '   Case 48 To 57, 52, 13, 8, 46, 27
   '   Case Else
   '        KeyAscii = 0
   '   End Select
   'End If
   If var_empresa = "16" Then
      If KeyAscii <> 13 Then
         KeyAscii = 0
      End If
   End If
   If KeyAscii = 13 Then
      Dim var_posible_eliminar_kanban As Boolean
      Dim var_si_kanban As Boolean
      If var_posible_kanban = 1 Then
         
      End If
      var_posible_eliminar_kanban = True
      var_si_kanban = False
      var_kanban_es_un_kanban = "N"
      VAR_SI_ELIMINA = 1
      If Mid(Me.txt_cantidad_eliminar, 1, 1) = "K" Then
         If IsNumeric(Me.txt_cantidad_eliminar) Then
            rs.Open "SELECT * FROM tb_fuera_kanban_en_entrada WHERE BINT_NUMERO_MOVIMIENTO = " + Me.txt_folio + " AND VCHA_TIPO_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND vcha_almacen_id = '" + var_almacen_Destino + "' AND VCHA_ARTICULO_ID = '" + Me.lv_entradas.selectedItem + "' AND FLOA_CANTIDAD >= " + Me.txt_cantidad_eliminar, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               VAR_SI_ELIMINA = 1
               var_si_kanban = True
            Else
               VAR_SI_ELIMINA = 0
               var_si_kanban = False
            End If
            rs.Close
         Else
            Set TB_ES_UN_KANBAN = New TB_ES_UN_KANBAN
            var_kanban = Me.txt_cantidad_eliminar
            var_inserta = TB_ES_UN_KANBAN.Anadir(Me.txt_cantidad_eliminar, "", "", "", "", "")
            var_kanban_es_un_kanban = var_kanban_es_un_kanban
            var_kanban_almacen_id = var_kanban_almacen_id
            var_kanban_articulo_id = var_kanban_articulo_id
            var_kanban_exito = var_kanban_exito
            var_kanban_mensaje = var_kanban_mensaje
            If var_kanban_es_un_kanban = "S" Then
               Me.txt_codigo = var_kanban_articulo_id
               If Me.txt_codigo = Me.lv_entradas.selectedItem Then
                  rsaux2.Open "SELECT * FROM TB_KANBANS_EN_MOVIMIENTOS WHERE VCHA_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_TIPO_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND BINT_NUMERO_MOVIMIENTO = " + Me.txt_folio + " AND VCHA_KANBAN_ID = '" + var_kanban + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     Me.txt_cantidad_eliminar = 1
                     var_posible_eliminar_kanban = True
                     var_si_kanban = True
                  Else
                     MsgBox "El kanban leido no se encuentra en esta nota", vbOKOnly, "ATENCION"
                  End If
                  rsaux2.Close
               Else
                  MsgBox "El c?digo del kanban leido no corresponde al c?digo del kanban seleccionado", vbOKOnly, "ATENCION"
                  var_posible_eliminar_kanban = False
               End If
            Else
               rsaux2.Open "SELECT * FROM TB_KANBANS_NOTA_ENVIO WHERE VCHA_COM_rEFERENCIA = '" + Me.txt_archivo + "' AND vcha_art_articulo_id = '" + Me.lv_entradas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  MsgBox "Debe de eliminar kanbans", vbOKOnly, "ATENCION"
                  var_posible_eliminar_kanban = False
               Else
                  var_posible_eliminar_kanban = True
               End If
               rsaux2.Close
            End If
         End If
      Else
         var_posible_eliminar_kanban = True
      End If
      
      If var_posible_eliminar_kanban = True Then
         If IsNumeric(txt_cantidad_eliminar) Then
            rsaux5.Open "select * from TB_EXISTENCIAS_SERIES where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_art_articulo_id = '" + Me.lv_entradas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux5.EOF Then
               If CDbl(Me.txt_cantidad_eliminar) = 1 Then
                  Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
                  Set TB_ARCH_COMPARACION_M = New TB_ARCH_COMPARACION_M
                  var_cantidad_eliminar = Val(txt_cantidad_eliminar)
                  var_cantidad_eliminar_arch = lv_entradas.selectedItem.SubItems(3) - Val(txt_cantidad_eliminar)
                  var_cantidad_eliminar_mov = lv_entradas.selectedItem.SubItems(4) - Val(txt_cantidad_eliminar)
                  If var_cantidad_eliminar_arch < 0 Or var_cantidad_eliminar_mov < 0 Then
                     MsgBox "No es posible eliminar esta cantidad", vbOKOnly, "ATENCION"
                  Else
                     If Trim(txt_cantidad_eliminar) <> "" Then
                        frmnumero_series.lv_lista.ListItems.Clear
                        frmnumero_series.txt_almacen_destino = var_almacen_Destino
                        frmnumero_series.txt_Cantidad = Me.txt_cantidad_eliminar
                        frmnumero_series.txt_movimiento = var_clave_movimiento
                        frmnumero_series.txt_numero = var_numero_folio
                        var_ventana = 1000
                        var_si_elimino = 0
                        While Not rsaux5.EOF
                           Set list_item = frmnumero_series.lv_lista.ListItems.Add(, , rsaux5!VCHA_ART_ARTICULO_ID)
                           list_item.SubItems(1) = IIf(IsNull(rsaux5!vcha_Art_numero_Serie), "", rsaux5!vcha_Art_numero_Serie)
                           rsaux5.MoveNext
                        Wend
                        frmnumero_series.Show 1
                        var_ventana = 0
                        var_inserta = False
                        If var_si_elimino = 1 Then
                           var_consecutivo = lv_entradas.selectedItem.SubItems(9)
                           rsaux3.Open "UPDATE TB_TEMPORAL_ENTRADAS SET VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "', FLOA_ENT_CANTIDAD = FLOA_ENT_CANTIDAD - " + txt_cantidad_eliminar + " Where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_ENT_NUMERO= '" + CStr(var_numero_folio) + "' AND VCHA_ART_ARTICULO_ID  = '" + lv_entradas.selectedItem + "' and inte_ent_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                           var_inserta = False
                           var_inserta = TB_ARCH_COMPARACION_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), var_tipo_proveedor, var_origen, lv_entradas.selectedItem, 0 - Val(txt_cantidad_eliminar), txt_archivo, var_consecutivo)
                           lbl_recibidos = Round(CDbl(lbl_recibidos) - var_cantidad_eliminar)
                           frm_eliminar.Visible = False
                           var_costo_tela = 0
                           txt_codigo.SetFocus
                           lv_entradas.selectedItem.SubItems(3) = Format((lv_entradas.selectedItem.SubItems(3) * 1) - var_cantidad_eliminar, "###,###,##0.00")
                           lv_entradas.selectedItem.SubItems(4) = Format((lv_entradas.selectedItem.SubItems(4) * 1) - var_cantidad_eliminar, "###,###,##0.00")
                           lv_entradas.selectedItem.SubItems(5) = Format((lv_entradas.selectedItem.SubItems(5) * 1) + var_cantidad_eliminar, "###,###,##0.00")
                           var_ventana = 0
                           var_renglon = lv_entradas.selectedItem.Index
                           Call ilumina_grid
                        End If
                     End If
                  End If
               Else
                  MsgBox "La cantidad debe de ser igual a 1", vbOKOnly, "ATENCION"
               End If
            Else
               Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
               Set TB_ARCH_COMPARACION_M = New TB_ARCH_COMPARACION_M
               var_cantidad_eliminar = Val(txt_cantidad_eliminar)
               var_cantidad_eliminar_arch = lv_entradas.selectedItem.SubItems(3) - Val(txt_cantidad_eliminar)
               var_cantidad_eliminar_mov = lv_entradas.selectedItem.SubItems(4) - Val(txt_cantidad_eliminar)
               If var_cantidad_eliminar_arch < 0 Or var_cantidad_eliminar_mov < 0 Then
                  MsgBox "No es posible eliminar esta cantidad", vbOKOnly, "ATENCION"
               Else
                  If Trim(txt_cantidad_eliminar) <> "" Then
                     If var_si_kanban = True Then
                        If var_kanban_es_un_kanban = "N" Then
                           If VAR_SI_ELIMINA = 1 Then
                              rs.Open "update tb_fuera_kanban_en_entrada set floa_cantidad =  floa_cantidad - " + Me.txt_cantidad_eliminar + " where BINT_NUMERO_MOVIMIENTO = " + Me.txt_folio + " AND VCHA_TIPO_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND vcha_almacen_id = '" + var_almacen_Destino + "' AND VCHA_ARTICULO_ID = '" + Me.lv_entradas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                           Else
                              var_kanban_mensaje = "La cantidad supera a la leida"
                           End If
                        Else
                           Set TB_CANCELAR_RES_KANBAN_ENTRADA = New TB_CANCELAR_RES_KANBAN_ENTRADA
                           var_inserta = TB_CANCELAR_RES_KANBAN_ENTRADA.Anadir(var_almacen_Destino, var_clave_movimiento, var_numero_folio, var_kanban, "", "")
                           If var_kanban_exito = "S" Then
                              VAR_SI_ELIMINA = 1
                           Else
                              VAR_SI_ELIMINA = 0
                           End If
                        End If
                     Else
                        If var_posible_kanban = 1 Then
                           VAR_SI_ELIMINA = 0
                        Else
                           VAR_SI_ELIMINA = 1
                        End If
                     End If
                     If VAR_SI_ELIMINA = 1 Then
                        var_inserta = False
                        var_consecutivo = lv_entradas.selectedItem.SubItems(9)

                        rsaux3.Open "UPDATE TB_TEMPORAL_ENTRADAS SET VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "', FLOA_ENT_CANTIDAD = FLOA_ENT_CANTIDAD - " + txt_cantidad_eliminar + " Where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_ENT_NUMERO= '" + CStr(var_numero_folio) + "' AND VCHA_ART_ARTICULO_ID  = '" + lv_entradas.selectedItem + "' and inte_ent_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                        '1
                        If var_clave_movimiento = "ETA" Or var_clave_movimiento = "EP" Or var_clave_movimiento = "EI" Then
                           '3
                           rsaux3.Open "update tb_transito set floa_tra_cantidad_recibida = isnull(floa_Tra_cantidad_recibida,0) - " + CStr(txt_cantidad_eliminar) + " where vcha_tra_nota_envio = '" + Me.lbl_transito + "' and vcha_Art_articulo_recivo = '" + lv_entradas.selectedItem + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
                        End If
                        var_inserta = False
                        var_inserta = TB_ARCH_COMPARACION_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), var_tipo_proveedor, var_origen, lv_entradas.selectedItem, 0 - Val(txt_cantidad_eliminar), txt_archivo, var_consecutivo)
                        lbl_recibidos = CDbl(lbl_recibidos) - var_cantidad_eliminar
                        frm_eliminar.Visible = False
                        var_costo_tela = 0
                        If Me.txt_codigo.Enabled = True Then
                           txt_codigo.SetFocus
                        End If
                        lv_entradas.selectedItem.SubItems(3) = Format((lv_entradas.selectedItem.SubItems(3) * 1) - var_cantidad_eliminar, "###,###,##0.00")
                        lv_entradas.selectedItem.SubItems(4) = Format((lv_entradas.selectedItem.SubItems(4) * 1) - var_cantidad_eliminar, "###,###,##0.00")
                        lv_entradas.selectedItem.SubItems(5) = Format((lv_entradas.selectedItem.SubItems(5) * 1) + var_cantidad_eliminar, "###,###,##0.00")
                        var_ventana = 0
                        var_renglon = lv_entradas.selectedItem.Index
                        Call ilumina_grid
                     Else
                        MsgBox var_kanban_mensaje, vbOKOnly, "ATENCION"
                     End If
                  End If
               End If
            End If
            rsaux5.Close
         Else
            MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   If KeyAscii = 27 Then
      frm_eliminar.Visible = False
      var_costo_tela = 0
      If Me.txt_codigo.Enabled = True Then
         txt_codigo.SetFocus
      End If
      var_ventana = 0
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   If var_ventana <> 1000 Then
      frm_eliminar.Visible = False
      var_costo_tela = 0
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_cantidad_GotFocus()
    Me.txt_codigo = var_codigo_tela
    txt_Cantidad = ""
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txt_Cantidad) <> "" Then
         var_cantidad_leida = txt_Cantidad
         txt_foco.Enabled = True
         txt_foco.SetFocus
         lbl_Cantidad.Visible = False
         txt_Cantidad.Visible = False
      End If
   End If
End Sub

Private Sub txt_codigo_caja_GotFocus()
   txt_codigo_caja = ""
End Sub

Private Sub txt_codigo_caja_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo_caja + "'", cnn, adOpenDynamic, adLockOptimistic
      If rs.EOF Then
         rsaux.Open "SELECT * FROM TB_EQUIVALENCIAS WHERE VCHA_EQU_CODIGO_EQUIVALENTE = '" + txt_codigo_caja + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            txt_codigo_caja = rsaux!VCHA_ART_ARTICULO_ID
         End If
         rsaux.Close
      End If
      rs.Close
      If Me.txt_codigo_caja = var_codigo_caja Then
         var_cantidad_leida = 1
         Me.txt_foco.Enabled = True
         Me.txt_foco.SetFocus
      Else
          frmmensaje.lbl_mensaje = "El art?culo no viene en la caja"
          frmmensaje.Show 1
         'MsgBox "El art?culo no viene en la caja", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_codigo_GotFocus()
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
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

   var_cantidad_multibondeados = 0
   var_renglon = 0
   var_numero_serie = 0
   var_posible_kanban = 0
   If var_codigo_seleccionado = "" Then
      txt_codigo = ""
   Else
      Me.txt_codigo = var_codigo_seleccionado
      var_codigo_seleccionado = ""
   End If
End Sub

Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_codigo_seleccionado = ""
      frmbusqueda_articulo.Show 1
      Me.txt_codigo = var_codigo_seleccionado
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   Dim var_recontable As Integer
   txt_codigo = Trim(txt_codigo)
   Dim var_caja As String
   Dim var_cantidad_caja As Integer
   Dim var_paquete As String
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If var_empresa <> "16" Then
      If var_empresa <> "06" Then
         If KeyAscii = 39 Or KeyAscii = 61 Then
            KeyAscii = 0
         End If
      End If
   End If
   If KeyAscii = 13 Then
      If rs.State = 1 Then
         rs.Close
      End If
      If rsaux.State = 1 Then
         rsaux.Close
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
      var_cantidad_multibondeados = 0
      If var_empresa = "16" Or var_empresa = "06" Then
         If var_empresa = "16" Then
            'var_cadena_X = ""
            'For var_jj = 1 To Len(Me.txt_codigo)
            '    If Mid(Me.txt_codigo, var_jj, 1) = "'" Then
            '       var_cadena_X = var_cadena_X + "-"
            '    Else
            '       var_cadena_X = var_cadena_X + Mid(Me.txt_codigo, var_jj, 1)
            '    End If
            'Next var_jj
            'Me.txt_codigo = var_cadena_X
            'var_cadena_X = ""
            'For var_jj = 1 To Len(Me.txt_codigo)
            '    If Mid(Me.txt_codigo, var_jj, 1) = "," Then
            '       var_cadena_X = var_cadena_X + "."
            '    Else
            '       var_cadena_X = var_cadena_X + Mid(Me.txt_codigo, var_jj, 1)
            '    End If
            'Next var_jj
            
            ''MsgBox var_cadena_x
            'Me.txt_codigo = var_cadena_X
            ''MsgBox Me.txt_codigo
         Else
            If Len(Me.txt_codigo) > 17 Then
               Me.txt_codigo = Mid(Me.txt_codigo, 3, Len(Me.txt_codigo))
               For var_jj = 1 To Len(Me.txt_codigo)
                   If Mid(Me.txt_codigo, var_jj, 1) = "'" Then
                      var_cadena_X = var_cadena_X + "-"
                   Else
                      var_cadena_X = var_cadena_X + Mid(Me.txt_codigo, var_jj, 1)
                   End If
               Next var_jj
               Me.txt_codigo = var_cadena_X
               'MsgBox Me.txt_codigo
               var_codigo = ""
               var_cantidad_str = ""
               var_j = Len(Me.txt_codigo)
               var_codigo_2 = 1
               var_lote_str = ""
               var_rollo_str = ""
               For var_j = 1 To Len(Me.txt_codigo)
                     If var_codigo_2 = 1 Then
                        If Mid(Me.txt_codigo, var_j, 1) <> "-" Then
                           var_lote_str = var_lote_str + Mid(Me.txt_codigo, var_j, 1)
                        Else
                           var_codigo_2 = 2
                        End If
                     Else
                        If var_codigo_2 = 2 Then
                           If Mid(Me.txt_codigo, var_j, 1) <> "-" Then
                              var_rollo_str = var_rollo_str + Mid(Me.txt_codigo, var_j, 1)
                           Else
                              var_codigo_2 = 3
                           End If
                        End If
                     End If
               Next var_j
               If var_lote_str = "0000000" Then
                  var_j = Len(Me.txt_codigo)
                  var_codigo_2 = 1
                  While var_j > 1
                     If var_codigo_2 = 1 Then
                        If Mid(Me.txt_codigo, var_j, 1) <> "-" Then
                           var_codigo = Mid(Me.txt_codigo, var_j, 1) + var_codigo
                           var_j = var_j - 1
                        Else
                           var_codigo_2 = 2
                           var_j = var_j - 1
                        End If
                     Else
                        If var_codigo_2 = 2 Then
                           If Mid(Me.txt_codigo, var_j, 1) <> "-" Then
                              var_cantidad_str = Mid(Me.txt_codigo, var_j, 1) + var_cantidad_str
                              var_j = var_j - 1
                           Else
                              var_codigo_2 = 3
                              var_j = 1
                           End If
                        End If
                     End If
                  Wend
               Else
                  var_lote_str = CStr(CDbl(var_lote_str))
                  rs.Open "select * from tb_rollos where vcha_lot_lote_id = '0_" + var_lote_str + "' and bint_num_rollo =" + var_rollo_str, cnn_estampados, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_codigo = rs!vcha_pro_producto_id
                     var_cantidad_str = rs!floa_cantidad_mts
                  Else
                     var_codigo = ""
                     var_cantidad_str = ""
                  End If
                  rs.Close
               End If
               If IsNumeric(var_cantidad_str) Then
                  var_cantidad_multibondeados = CDbl(var_cantidad_str)
               Else
                  var_cantidad_multibondeados = 0
               End If
               Me.txt_codigo = var_codigo
            End If
         
         
         
         
         End If
      End If
      If var_empresa <> "16" Then
         var_codigo_tela = ""
         var_costo_tela = 0
         If var_cajas = True Then
            var_peso_caja = 0
            var_paquete = Left(txt_codigo, 2)
            var_codigo_caja = ""
            var_cantidad_caja_peso = 0
            If var_paquete = "CA" Then
               If rs.State = 1 Then
                  rs.Close
               End If
               rs.Open "select * from tb_archivo_comparacion where  vcha_com_referencia = '" + txt_archivo + "' and vcha_com_caja = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  If rs!FLOA_com_cANTIDAD_RECIBIDA = 0 Then
                     var_codigo_caja = rs!VCHA_ART_ARTICULO_ID
                     var_cantidad_caja_peso = IIf(IsNull(rs!FLOA_COM_CANTIDAD_ENVIADA), 0, rs!FLOA_COM_CANTIDAD_ENVIADA)
                     rsaux2.Open "select * FROM TB_ARTICULOS where vcha_art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        var_posible_kanban = IIf(IsNull(rsaux2!INTE_ART_KANBAN), 0, rsaux2!INTE_ART_KANBAN)
                        Me.txt_articulo_caja = rsaux2!vcha_Art_nombre_espa?ol
                        var_numero_serie = IIf(IsNull(rsaux2!inte_art_numero_serie), 0, rsaux2!inte_art_numero_serie)
                     Else
                        Me.txt_articulo_caja = ""
                     End If
                     rsaux2.Close
                     var_peso_caja = IIf(IsNull(rs!FLOA_COM_PESO), 0, rs!FLOA_COM_PESO)
                     If var_peso_caja > 0 Then
                        txt_peso_caja = ""
                        frm_peso_caja.Visible = True
                        var_ventana = 1
                        txt_peso_caja.SetFocus
                     Else
                     End If
                  Else
                     txt_codigo = ""
                     frmmensaje.lbl_mensaje = "La caja ya fue leida"
                     frmmensaje.Show 1
                     'MsgBox "La caja ya fue leida", vbOKOnly, "ATENCION"
                  End If
               Else
                  txt_codigo = ""
                  frmmensaje.lbl_mensaje = "C?digo de caja no existe"
                  frmmensaje.Show 1
                  'MsgBox "C?digo de caja no existe", vbOKOnly, "ATENCION"
               End If
               If rs.State = 1 Then
                  rs.Close
               End If
            Else
               txt_codigo = ""
               frmmensaje.lbl_mensaje = "Debera de leer cajas"
               frmmensaje.Show 1
               'MsgBox "Debera de leer cajas", vbOKOnly, "ATENCION"
            End If
         Else
            If UCase(Mid(Me.txt_codigo, 1, 1)) = "K" And var_clave_movimiento = "EP" Then
               Set TB_ES_UN_KANBAN = New TB_ES_UN_KANBAN
               var_kanban = Me.txt_codigo
               var_inserta = TB_ES_UN_KANBAN.Anadir(Me.txt_codigo, "", "", "", "", "")
               var_kanban_es_un_kanban = var_kanban_es_un_kanban
               var_kanban_almacen_id = var_kanban_almacen_id
               var_kanban_articulo_id = var_kanban_articulo_id
               var_kanban_exito = var_kanban_exito
               var_kanban_mensaje = var_kanban_mensaje
                    
               If var_kanban_es_un_kanban = "S" Then
                  Me.txt_codigo = var_kanban_articulo_id
               End If
                                                   
               var_verificador = True
               If Len(Trim(txt_codigo)) = 12 Then
                  If var_empresa <> 31 Then
                     Call calcula_verificador(Trim(txt_codigo))
                  End If
               End If
               If var_verificador = True Then
                  If Trim(txt_codigo) <> "" Then
                     var_caja = Left(txt_codigo, 6)
                     If Left(Me.txt_codigo, 1) = "0" And var_empresa <> "31" Then
                     'If var_caja = "000005" Or var_caja = "000010" Or var_caja = "000015" Or var_caja = "000020" Or var_caja = "000001" Or var_caja = "000002" Or var_caja = "000003" Or var_caja = "000004" Or var_caja = "000006" Or var_caja = "000007" Or var_caja = "000008" Or var_caja = "000009" Or var_caja = "000011" Or var_caja = "0000012" Or var_caja = "0000013" Or var_caja = "0000014" Or var_caja = "000015" Or var_caja = "000016" Or var_caja = "000017" Or var_caja = "000018" Or var_caja = "000019" Or var_caja = "000021" Or var_caja = "000022" Or var_caja = "000023" Or var_caja = "000024" Or var_caja = "000025" Or var_caja = "000026" Or var_caja = "000027" Or var_caja = "000028" Or var_caja = "000029" Or var_caja = "000030" Then
                        var_cantidad_caja = CInt(var_caja)
                        txt_codigo = Mid(txt_codigo, 7, 5)
                     End If
                     If rsaux4.State = 1 Then
                        rsaux4.Close
                     End If
                     rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux4.EOF Then
                        var_posible_kanban = IIf(IsNull(rsaux4!INTE_ART_KANBAN), 0, rsaux4!INTE_ART_KANBAN)
                        var_codigo_tela = Me.txt_codigo
                        var_numero_serie = IIf(IsNull(rsaux4!inte_art_numero_serie), 0, rsaux4!inte_art_numero_serie)
                        If var_canitad_caja = 0 Then
                           If IsNull(rsaux4(43).Value) Then
                              If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                                 var_recontable = 1
                              Else
                                 var_recontable = 0
                              End If
                           Else
                              If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                                 var_recontable = 1
                              Else
                                 var_recontable = rsaux4(43).Value
                              End If
                           End If
                        Else
                           If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                              var_recontable = 1
                           Else
                              var_recontable = 0
                           End If
                        End If
                        var_nombre_articulo = IIf(IsNull(rsaux4!vcha_Art_nombre_espa?ol), "", rsaux4!vcha_Art_nombre_espa?ol)
                        rsaux4.Close
                        If var_numero_serie = 1 Then
                           If var_recontable = 1 Then
                              var_cantidad_leida = 1#
                              lbl_Cantidad.Visible = True
                              txt_Cantidad.Visible = True
                           End If
                           'Me.txt_numero_serie = ""
                           'Me.frmnumero_serie.Visible = True
                           'txt_numero_serie.Enabled = True
                           'var_ventana = 1
                           'Me.txt_numero_serie.SetFocus
                           Me.txt_foco.SetFocus
                        Else
                           If var_recontable = 1 Then
                              var_cantidad_leida = 1#
                              lbl_Cantidad.Visible = True
                              txt_Cantidad.Visible = True
                              'If (var_empresa = "18" Or var_empresa = "06") And var_proveedor = "2458" Then
                              '   var_codigo_tela = Me.txt_codigo
                              '   rsaux10.Open "select * from tb_equivalencias where vcha_art_Articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                              '   If Not rsaux10.EOF Then
                              '      rsaux11.Open "SELECT * FROM TB_FACTURAS_ESTAMPADOS_REFUGIO WHERE VCHA_FAC_FACTURA = '" + Me.txt_factura + "' and vcha_fac_codigo_externo = '" + IIf(IsNull(rsaux10!vcha_equ_codigo_equivalente), "", rsaux10!vcha_equ_codigo_equivalente) + "'", cnn, adOpenDynamic, adLockOptimistic
                              '      If Not rsaux11.EOF Then
                              '         var_costo_tela = IIf(IsNull(rsaux11!floa_fac_costo), 0, rsaux11!floa_fac_costo)
                              '      Else
                              '         frmfactura_orden_compra_estampados.lbl_articulo = var_nombre_articulo
                              '         frmfactura_orden_compra_estampados.txt_factura = var_factura
                              '         frmfactura_orden_compra_estampados.Show 1
                              '         Me.txt_codigo = var_codigo_tela
                              '      End If
                              '      rsaux11.Close
                              '   Else
                              '      frmfactura_orden_compra_estampados.lbl_articulo = var_nombre_articulo
                              '      frmfactura_orden_compra_estampados.txt_factura = var_factura
                              '      frmfactura_orden_compra_estampados.Show 1
                              '      Me.txt_codigo = var_codigo_tela
                              '   End If
                              '   rsaux10.Close
                              'End If
                              txt_Cantidad.SetFocus
                           Else
                              If var_cantidad_caja > 0 Then
                                 var_cantidad_leida = var_cantidad_caja
                              Else
                                 var_cantidad_leida = 1#
                              End If
                              'If (var_empresa = "18" Or var_empresa = "06") And var_proveedor = "2458" Then
                              '   var_codigo_tela = Me.txt_codigo
                              '   frmfactura_orden_compra_estampados.lbl_articulo = var_nombre_articulo
                              '   frmfactura_orden_compra_estampados.txt_factura = var_factura
                              '   frmfactura_orden_compra_estampados.Show 1
                              '   Me.txt_codigo = var_codigo_tela
                              'End If
                              txt_foco.Enabled = True
                              txt_foco.SetFocus
                           End If
                        End If
                     Else
                        rsaux4.Close
                        If rs.State = 1 Then
                           rs.Close
                        End If
                        rs.Open "select * from tb_equivalencias where VCHA_EQU_CODIGO_EQUIVALENTE = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           txt_codigo = rs(0).Value
                           rs.Close
                           rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rs.EOF Then
                              var_posible_kanban = IIf(IsNull(rs!INTE_ART_KANBAN), 0, rs!INTE_ART_KANBAN)
                              var_codigo_tela = Me.txt_codigo
                              var_nombre_articulo = IIf(IsNull(rs!vcha_Art_nombre_espa?ol), "", rs!vcha_Art_nombre_espa?ol)
                              var_numero_serie = IIf(IsNull(rs!inte_art_numero_serie), 0, rs!inte_art_numero_serie)
                              If var_cantidad_caja = 0 Then
                                 If IsNull(rs(43).Value) Then
                                    If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                                       var_recontable = 1
                                    Else
                                       var_recontable = 0
                                    End If
                                 Else
                                    If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                                       var_recontable = 1
                                    Else
                                       var_recontable = rs(43).Value
                                    End If
                                 End If
                              Else
                                 If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                                    var_recontable = 1
                                 Else
                                    var_recontable = 0
                                 End If
                              End If
                              rs.Close
                              If var_numero_serie = 1 Then
                                 If var_recontable = 1 Then
                                    var_cantidad_leida = 1#
                                    lbl_Cantidad.Visible = True
                                    txt_Cantidad.Visible = True
                                 End If
                                 'Me.txt_numero_serie = ""
                                 'Me.frmnumero_serie.Visible = True
                                 'txt_numero_serie.Enabled = True
                                 'var_ventana = 1
                                 'Me.txt_numero_serie.SetFocus
                              Else
                                 If var_recontable = 1 Then
                                    'If var_empresa = "18" And var_proveedor = "2458" Then
                                    '   var_codigo_tela = Me.txt_codigo
                                    '   frmfactura_orden_compra_estampados.lbl_articulo = var_nombre_articulo
                                    '   frmfactura_orden_compra_estampados.txt_factura = var_factura
                                    '   frmfactura_orden_compra_estampados.Show 1
                                    '   Me.txt_codigo = var_codigo_tela
                                    'End If
                                    var_cantidad_leida = 1#
                                    lbl_Cantidad.Visible = True
                                    txt_Cantidad.Visible = True
                                    txt_Cantidad.SetFocus
                                 Else
                                    If var_cantidad_caja = 0 Then
                                       var_cantidad_leida = 1#
                                    Else
                                       var_cantidad_leida = var_cantidad_caja
                                    End If
                                    'If var_empresa = "18" And var_proveedor = "2458" Then
                                    '   var_codigo_tela = Me.txt_codigo
                                    '   frmfactura_orden_compra_estampados.lbl_articulo = var_nombre_articulo
                                    '   frmfactura_orden_compra_estampados.txt_factura = var_factura
                                    '   frmfactura_orden_compra_estampados.Show 1
                                    '   Me.txt_codigo = var_codigo_tela
                                    'End If
                                    txt_foco.Enabled = True
                                    txt_foco.SetFocus
                                 End If
                              End If
                           Else
                              txt_codigo = ""
                              frmmensaje.lbl_mensaje = "El art?culo no existe"
                              frmmensaje.Show 1
                              'MsgBox "El art?culo no existe", vbOKOnly, "ATENCION"
                           End If
                        Else
                           txt_codigo = ""
                           frmmensaje.lbl_mensaje = "El art?culo no existe"
                           frmmensaje.Show 1
                           'MsgBox "El art?culo no existe", vbOKOnly, "ATENCION"
                           rs.Close
                        End If
                     End If
                  Else
                      txt_codigo = ""
                      frmmensaje.lbl_mensaje = "C?digo Incorrecto"
                      frmmensaje.Show 1
                     'MsgBox "C?digo Incorrecto", vbOKOnly, "ATENCION"
                   End If
               Else
                  txt_codigo = ""
                  frmmensaje.lbl_mensaje = "Error en C?digo"
                  frmmensaje.Show 1
                  'MsgBox "Error en C?digo", vbOKOnly, "ATENCION"
               End If
               'fin kanban
            Else
               var_verificador = True
               If Len(Trim(txt_codigo)) = 12 Then
                  If var_empresa <> 31 Then
                     Call calcula_verificador(Trim(txt_codigo))
                  End If
               End If
               If var_verificador = True Then
                  If Trim(txt_codigo) <> "" Then
                     var_caja = Left(txt_codigo, 6)
                     If Left(Me.txt_codigo, 1) = "0" And var_empresa <> "31" Then
                     'If var_caja = "000005" Or var_caja = "000010" Or var_caja = "000015" Or var_caja = "000020" Or var_caja = "000001" Or var_caja = "000002" Or var_caja = "000003" Or var_caja = "000004" Or var_caja = "000006" Or var_caja = "000007" Or var_caja = "000008" Or var_caja = "000009" Or var_caja = "000011" Or var_caja = "0000012" Or var_caja = "0000013" Or var_caja = "0000014" Or var_caja = "000015" Or var_caja = "000016" Or var_caja = "000017" Or var_caja = "000018" Or var_caja = "000019" Or var_caja = "000021" Or var_caja = "000022" Or var_caja = "000023" Or var_caja = "000024" Or var_caja = "000025" Or var_caja = "000026" Or var_caja = "000027" Or var_caja = "000028" Or var_caja = "000029" Or var_caja = "000030" Then
                        var_cantidad_caja = CInt(var_caja)
                        txt_codigo = Mid(txt_codigo, 7, 5)
                     End If
                     If rsaux4.State = 1 Then
                        rsaux4.Close
                     End If
                     rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux4.EOF Then
                        var_posible_kanban = IIf(IsNull(rsaux4!INTE_ART_KANBAN), 0, rsaux4!INTE_ART_KANBAN)
                        var_codigo_tela = Me.txt_codigo
                        var_numero_serie = IIf(IsNull(rsaux4!inte_art_numero_serie), 0, rsaux4!inte_art_numero_serie)
                        If var_canitad_caja = 0 Then
                           If IsNull(rsaux4(43).Value) Then
                              If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                                 var_recontable = 1
                              Else
                                 var_recontable = 0
                              End If
                           Else
                              If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                                 var_recontable = 1
                              Else
                                 var_recontable = rsaux4(43).Value
                              End If
                           End If
                        Else
                           If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                              var_recontable = 1
                           Else
                              var_recontable = 0
                           End If
                        End If
                        var_nombre_articulo = IIf(IsNull(rsaux4!vcha_Art_nombre_espa?ol), "", rsaux4!vcha_Art_nombre_espa?ol)
                        rsaux4.Close
                        If var_numero_serie = 1 Then
                           If var_recontable = 1 Then
                              var_cantidad_leida = 1#
                              lbl_Cantidad.Visible = True
                              txt_Cantidad.Visible = True
                           End If
                           'Me.txt_numero_serie = ""
                           'Me.frmnumero_serie.Visible = True
                           'txt_numero_serie.Enabled = True
                           'var_ventana = 1
                           'Me.txt_numero_serie.SetFocus
                           Me.txt_foco.SetFocus
                        Else
                           If var_cantidad_multibondeados > 0 Then
                              If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                                 var_recontable = 1
                              Else
                                 var_recontable = 0
                              End If
                           End If
                           If var_recontable = 1 Then
                              var_cantidad_leida = 1#
                              lbl_Cantidad.Visible = True
                              txt_Cantidad.Visible = True
                              
                              'If (var_empresa = "18" Or var_empresa = "06") And var_proveedor = "2458" Then
                              '   var_codigo_tela = Me.txt_codigo
                              '   frmfactura_orden_compra_estampados.lbl_articulo = var_nombre_articulo
                              '   frmfactura_orden_compra_estampados.txt_factura = var_factura
                              '   frmfactura_orden_compra_estampados.Show 1
                              '   Me.txt_codigo = var_codigo_tela
                              'End If
                              txt_Cantidad.SetFocus
                           Else
                              If var_cantidad_caja > 0 Then
                                 var_cantidad_leida = var_cantidad_caja
                              Else
                                 If var_cantidad_multibondeados = 0 Then
                                    var_cantidad_leida = 1#
                                 Else
                                    var_cantidad_leida = var_cantidad_multibondeados
                                 End If
                              End If
                              'If (var_empresa = "18" Or var_empresa = "06") And var_proveedor = "2458" Then
                              '   ' AQUI VA EL CODIGO_MAL
                              '   var_codigo_tela = Me.txt_codigo
                              '   rsaux10.Open "select * from tb_equivalencias where vcha_art_Articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                              '   If Not rsaux10.EOF Then
                              '      rsaux11.Open "SELECT * FROM TB_FACTURAS_ESTAMPADOS_REFUGIO WHERE VCHA_FAC_FACTURA = '" + Me.txt_factura + "' and vcha_fac_codigo_externo = '" + IIf(IsNull(rsaux10!vcha_equ_codigo_equivalente), "", rsaux10!vcha_equ_codigo_equivalente) + "'", cnn, adOpenDynamic, adLockOptimistic
                              '      If Not rsaux11.EOF Then
                              '         var_costo_tela = IIf(IsNull(rsaux11!floa_fac_costo), 0, rsaux11!floa_fac_costo)
                              '      Else
                              '         frmfactura_orden_compra_estampados.lbl_articulo = var_nombre_articulo
                              '         frmfactura_orden_compra_estampados.txt_factura = var_factura
                              '         frmfactura_orden_compra_estampados.Show 1
                              '         Me.txt_codigo.SetFocus
                              '         Me.txt_codigo = var_codigo_tela
                              '
                              '      End If
                              '      rsaux11.Close
                              '   Else
                              '      frmfactura_orden_compra_estampados.lbl_articulo = var_nombre_articulo
                              '      frmfactura_orden_compra_estampados.txt_factura = var_factura
                              '      frmfactura_orden_compra_estampados.Show 1
                              '      Me.txt_codigo = var_codigo_tela
                              '   End If
                              '   rsaux10.Close
                              '
                              '
                              '
                              'End If
                              txt_foco.Enabled = True
                              txt_foco.SetFocus
                           End If
                        End If
                     Else
                        rsaux4.Close
                        If rs.State = 1 Then
                           rs.Close
                        End If
                        rs.Open "select * from tb_equivalencias where VCHA_EQU_CODIGO_EQUIVALENTE = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           txt_codigo = rs(0).Value
                           rs.Close
                           rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rs.EOF Then
                              var_codigo_tela = Me.txt_codigo
                              var_posible_kanban = IIf(IsNull(rs!INTE_ART_KANBAN), 0, rs!INTE_ART_KANBAN)
                              var_nombre_articulo = IIf(IsNull(rs!vcha_Art_nombre_espa?ol), "", rs!vcha_Art_nombre_espa?ol)
                              var_numero_serie = IIf(IsNull(rs!inte_art_numero_serie), 0, rs!inte_art_numero_serie)
                              If var_cantidad_caja = 0 Then
                                 If IsNull(rs(43).Value) Then
                                    var_recontable = 0
                                 Else
                                    If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                                       var_recontable = 1
                                    Else
                                       var_recontable = rs(43).Value
                                    End If
                                 End If
                              Else
                                 If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                                    var_recontable = 1
                                 Else
                                    var_recontable = 0
                                 End If
                              End If
                              rs.Close
                              If var_numero_serie = 1 Then
                                 If var_recontable = 1 Then
                                    var_cantidad_leida = 1#
                                    lbl_Cantidad.Visible = True
                                    txt_Cantidad.Visible = True
                                 End If
                                 'Me.txt_numero_serie = ""
                                 'Me.frmnumero_serie.Visible = True
                                 'txt_numero_serie.Enabled = True
                                 'var_ventana = 1
                                 'Me.txt_numero_serie.SetFocus
                              Else
                                 If var_recontable = 1 Then
                                    'If var_empresa = "18" And var_proveedor = "2458" Then
                                    '   var_codigo_tela = Me.txt_codigo
                                    '   frmfactura_orden_compra_estampados.lbl_articulo = var_nombre_articulo
                                    '   frmfactura_orden_compra_estampados.txt_factura = var_factura
                                    '   frmfactura_orden_compra_estampados.Show 1
                                    '   Me.txt_codigo = var_codigo_tela
                                    'End If
                                    var_cantidad_leida = 1#
                                    lbl_Cantidad.Visible = True
                                    txt_Cantidad.Visible = True
                                    txt_Cantidad.SetFocus
                                 Else
                                    If var_cantidad_caja = 0 Then
                                       var_cantidad_leida = 1#
                                    Else
                                       var_cantidad_leida = var_cantidad_caja
                                    End If
                                    'If var_empresa = "18" And var_proveedor = "2458" Then
                                    '   var_codigo_tela = Me.txt_codigo
                                    '   frmfactura_orden_compra_estampados.lbl_articulo = var_nombre_articulo
                                    '   frmfactura_orden_compra_estampados.txt_factura = var_factura
                                    '   frmfactura_orden_compra_estampados.Show 1
                                    '   Me.txt_codigo = var_codigo_tela
                                    'End If
                                    txt_foco.Enabled = True
                                    txt_foco.SetFocus
                                 End If
                              End If
                           Else
                              txt_codigo = ""
                              frmmensaje.lbl_mensaje = "El art?culo no existe"
                              frmmensaje.Show 1
                              'MsgBox "El art?culo no existe", vbOKOnly, "ATENCION"
                           End If
                        Else
                           txt_codigo = ""
                           frmmensaje.lbl_mensaje = "El art?culo no existe"
                           frmmensaje.Show 1
                           'MsgBox "El art?culo no existe", vbOKOnly, "ATENCION"
                           rs.Close
                        End If
                     End If
                  Else
                      txt_codigo = ""
                      frmmensaje.lbl_mensaje = "C?digo Incorrecto"
                      frmmensaje.Show 1
                     'MsgBox "C?digo Incorrecto", vbOKOnly, "ATENCION"
                   End If
               Else
                  txt_codigo = ""
                  frmmensaje.lbl_mensaje = "Error en C?digo"
                  frmmensaje.Show 1
                  'MsgBox "Error en C?digo", vbOKOnly, "ATENCION"
               End If
               'fin kanban
            End If
         End If
      Else
         If Me.txt_codigo <> "" Then
            If Len(Trim(Me.txt_codigo)) > 9 Then
''''cccc
               If var_empresa <> "16" Then
                  var_cx = Trim(Mid(Me.txt_codigo, 7, 10))
                  var_cc = Mid(Me.txt_codigo, 1, 6)
                  If IsNumeric(Mid(Me.txt_codigo, 1, 6)) Then
                     rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id = '" + Trim(Mid(Me.txt_codigo, 7, 10)) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux4.EOF Then
                        var_posible_kanban = IIf(IsNull(rsaux4!INTE_ART_KANBAN), 0, rsaux4!INTE_ART_KANBAN)
                        If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                           var_recontable = 1
                        Else
                           var_recontable = 0
                        End If
                        var_nombre_articulo = IIf(IsNull(rsaux4!vcha_Art_nombre_espa?ol), "", rsaux4!vcha_Art_nombre_espa?ol)
                        rsaux4.Close
                        var_cantidad_leida = CDbl(Mid(Me.txt_codigo, 1, 6))
                        Me.txt_codigo = Trim(Mid(Me.txt_codigo, 7, 10))
                        txt_foco.Enabled = True
                        txt_foco.SetFocus
                     Else
                        txt_codigo = ""
                        frmmensaje.lbl_mensaje = "El art?culo no existe"
                        frmmensaje.Show 1
                        rsaux4.Close
                     End If
                  Else
                     txt_codigo = ""
                     frmmensaje.lbl_mensaje = "Cantidad incorrecta"
                     frmmensaje.Show 1
                  End If
''''cccc
               Else
               ''' aqui debe de ir lo de multibondeados
                  If var_empresa = "16" Then
                     If var_empresa = "16" Then
                        x = Mid(Me.txt_codigo, Len(Me.txt_codigo), 1)
                        'MsgBox x
                        If x = "-" Then
                           If Len(Me.txt_codigo) > 8 Then
                              var_c = Mid(Me.txt_codigo, 3, Len(Me.txt_codigo) - 10)
                              'MsgBox var_c
                              If IsNumeric(var_c) Then
                                 var_cantidad_multibondeados = CDbl(var_c)
                                 Me.txt_codigo = Right(Me.txt_codigo, 8)
                                 'MsgBox Me.txt_codigo
                              End If
                           End If
                        End If
                        If UCase(x) = "B" Or UCase(x) = "R" Or UCase(x) = "C" Then
                           If Len(Me.txt_codigo) > 9 Then
                              var_c = Mid(Me.txt_codigo, 1, Len(Me.txt_codigo) - 9)
                              If IsNumeric(var_c) Then
                                 var_cantidad_multibondeados = CDbl(var_c)
                                 Me.txt_codigo = Right(Me.txt_codigo, 9)
                                 'MsgBox Me.txt_codigo
                              End If
                           End If
                        End If
                     Else
                     End If
                     rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux4.EOF Then
                        var_posible_kanban = IIf(IsNull(rsaux4!INTE_ART_KANBAN), 0, rsaux4!INTE_ART_KANBAN)
                        var_codigo_tela = Me.txt_codigo
                        var_numero_serie = IIf(IsNull(rsaux4!inte_art_numero_serie), 0, rsaux4!inte_art_numero_serie)
                        If var_canitad_caja = 0 Then
                           If IsNull(rsaux4(43).Value) Then
                              If var_cantidad_multibondeados = 0 Then
                                 If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                                    var_recontable = 1
                                 Else
                                    var_recontable = IIf(IsNull(rs(43).Value), 0, rs(43).Value)
                                 End If
                              Else
                                 If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                                    var_recontable = 1
                                 Else
                                    var_recontable = 0
                                 End If
                              End If
                           Else
                              If var_cantidad_multibondeados = 0 Then
                                 If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                                    var_recontable = 1
                                 Else
                                    var_recontable = IIf(IsNull(rsaux4(43).Value), 0, rsaux4(43).Value)
                                 End If
                              Else
                                 If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                                    var_recontable = 1
                                 Else
                                    var_recontable = 0
                                 End If
                              End If
                           End If
                        Else
                           If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                              var_recontable = 1
                           Else
                              var_recontable = 0
                           End If
                        End If
                        var_nombre_articulo = IIf(IsNull(rsaux4!vcha_Art_nombre_espa?ol), "", rsaux4!vcha_Art_nombre_espa?ol)
                        rsaux4.Close
                        If var_numero_serie = 1 Then
                           If var_recontable = 1 Then
                              If var_cantidad_multibondeados = 0 Then
                                 var_cantidad_leida = 1#
                              Else
                                 var_cantidad_leida = var_cantidad_multibondeados
                              End If
                              lbl_Cantidad.Visible = True
                              txt_Cantidad.Visible = True
                           End If
                           Me.txt_foco.SetFocus
                        Else
                           If var_recontable = 1 Then
                              var_cantidad_leida = 1#
                              lbl_Cantidad.Visible = True
                              txt_Cantidad.Visible = True
                              txt_Cantidad.SetFocus
                           Else
                              If var_cantidad_caja > 0 Then
                                 var_cantidad_leida = var_cantidad_caja
                              Else
                                 'MsgBox CStr(var_cantidad_multibondeados)
                                 If var_cantidad_multibondeados = 0 Then
                                    var_cantidad_leida = 1#
                                 Else
                                    var_cantidad_leida = var_cantidad_multibondeados
                                 End If
                              End If
                              txt_foco.Enabled = True
                              txt_foco.SetFocus
                           End If
                        End If
                     Else
                        txt_codigo = ""
                        frmmensaje.lbl_mensaje = "El art?culo no existe"
                        frmmensaje.Show 1
                     End If
                     If rsaux4.State = 1 Then
                        rsaux4.Close
                     End If
                  End If
''' fin de multibondeados
               End If
            Else
               ''' aqui debe de ir lo de multibondeados
               If var_empresa = "16" Then
                  'x = Mid(Me.txt_codigo, Len(Me.txt_codigo), 1)
                  'If x = "-" Then
                  '   If Len(Me.txt_codigo) > 8 Then
                  '      var_c = Mid(Me.txt_codigo, 1, Len(Me.txt_codigo) - 8)
                  '      If IsNumeric(var_c) Then
                  '         var_cantidad_multibondeados = CDbl(var_c)
                  '         Me.txt_codigo = Right(Me.txt_codigo, 8)
                  '         'MsgBox Me.txt_codigo
                  '      End If
                  '   End If
                  'End If
                  'If UCase(x) = "B" Or UCase(x) = "R" Or UCase(x) = "C" Then
                  '   If Len(Me.txt_codigo) > 9 Then
                  '      var_c = Mid(Me.txt_codigo, 1, Len(Me.txt_codigo) - 9)
                  '      If IsNumeric(var_c) Then
                  '         var_cantidad_multibondeados = CDbl(var_c)
                  '         Me.txt_codigo = Right(Me.txt_codigo, 9)
                  '         'MsgBox Me.txt_codigo
                  '      End If
                  '   End If
                  'End If
                  
                  rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux4.EOF Then
                     var_posible_kanban = IIf(IsNull(rsaux4!INTE_ART_KANBAN), 0, rsaux4!INTE_ART_KANBAN)
                     var_codigo_tela = Me.txt_codigo
                     var_numero_serie = IIf(IsNull(rsaux4!inte_art_numero_serie), 0, rsaux4!inte_art_numero_serie)
                     If var_canitad_caja = 0 Then
                        If IsNull(rsaux4(43).Value) Then
                           If var_cantidad_multibondeados = 0 Then
                              var_recontable = IIf(IsNull(rs(43).Value), 0, rs(43).Value)
                           Else
                              var_recontable = 0
                           End If
                        Else
                           If var_cantidad_multibondeados = 0 Then
                              var_recontable = IIf(IsNull(rsaux4(43).Value), 0, rsaux4(43).Value)
                           Else
                              var_recontable = 0
                           End If
                        End If
                     Else
                        var_recontable = 0
                     End If
                     var_nombre_articulo = IIf(IsNull(rsaux4!vcha_Art_nombre_espa?ol), "", rsaux4!vcha_Art_nombre_espa?ol)
                     rsaux4.Close
                     If var_numero_serie = 1 Then
                        If var_recontable = 1 Then
                           If var_cantidad_multibondeados = 0 Then
                              var_cantidad_leida = 1#
                           Else
                              var_cantidad_leida = var_cantidad_multibondeados
                           End If
                           lbl_Cantidad.Visible = True
                           txt_Cantidad.Visible = True
                        End If
                        Me.txt_foco.SetFocus
                     Else
                        If var_recontable = 1 Then
                           var_cantidad_leida = 1#
                           lbl_Cantidad.Visible = True
                           txt_Cantidad.Visible = True
                           txt_Cantidad.SetFocus
                        Else
                           If var_cantidad_caja > 0 Then
                              var_cantidad_leida = var_cantidad_caja
                           Else
                              If var_cantidad_multibondeados = 0 Then
                                 var_cantidad_leida = 1#
                              Else
                                 var_cantidad_leida = var_cantidad_multibondeados
                              End If
                           End If
                           txt_foco.Enabled = True
                           txt_foco.SetFocus
                        End If
                     End If
                  Else
                     txt_codigo = ""
                     frmmensaje.lbl_mensaje = "El art?culo no existe"
                     frmmensaje.Show 1
                  End If
                  If rsaux4.State = 1 Then
                     rsaux4.Close
                  End If
               
               
               
               End If
''' fin de multibondeados
               
               
               rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                  var_posible_kanban = IIf(IsNull(rsaux4!INTE_ART_KANBAN), 0, rsaux4!INTE_ART_KANBAN)
                  var_codigo_tela = Me.txt_codigo
                  var_numero_serie = IIf(IsNull(rsaux4!inte_art_numero_serie), 0, rsaux4!inte_art_numero_serie)
                  If var_canitad_caja = 0 Then
                     If IsNull(rsaux4(43).Value) Then
                        If var_cantidad_multibondeados = 0 Then
                           If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                              var_recontable = 1
                           Else
                              var_recontable = IIf(IsNull(rs(43).Value), 0, rs(43).Value)
                           End If
                        Else
                           If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                              var_recontable = 1
                           Else
                              var_recontable = 0
                           End If
                        End If
                     Else
                        If var_cantidad_multibondeados = 0 Then
                           If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                              var_recontable = 1
                           Else
                              var_recontable = IIf(IsNull(rsaux4(43).Value), 0, rsaux4(43).Value)
                           End If
                        Else
                           If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                              var_recontable = 1
                           Else
                              var_recontable = 0
                           End If
                        End If
                     End If
                  Else
                     If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                        var_recontable = 1
                     Else
                        var_recontable = 0
                     End If
                  End If
                  var_nombre_articulo = IIf(IsNull(rsaux4!vcha_Art_nombre_espa?ol), "", rsaux4!vcha_Art_nombre_espa?ol)
                  rsaux4.Close
                  If var_numero_serie = 1 Then
                     If var_recontable = 1 Then
                        If var_cantidad_multibondeados = 0 Then
                           var_cantidad_leida = 1#
                        Else
                           var_cantidad_leida = var_cantidad_multibondeados
                        End If
                        lbl_Cantidad.Visible = True
                        txt_Cantidad.Visible = True
                     End If
                     Me.txt_foco.SetFocus
                  Else
                     If var_recontable = 1 Then
                        var_cantidad_leida = 1#
                        lbl_Cantidad.Visible = True
                        txt_Cantidad.Visible = True
                        txt_Cantidad.SetFocus
                     Else
                        If var_cantidad_caja > 0 Then
                           var_cantidad_leida = var_cantidad_caja
                        Else
                           If var_cantidad_multibondeados = 0 Then
                              var_cantidad_leida = 1#
                           Else
                              var_cantidad_leida = var_cantidad_multibondeados
                           End If
                        End If
                        txt_foco.Enabled = True
                        txt_foco.SetFocus
                     End If
                  End If
               Else
                  rsaux4.Close
                  rs.Open "select * from tb_equivalencias where VCHA_EQU_CODIGO_EQUIVALENTE = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     txt_codigo = rs(0).Value
                     rs.Close
                     rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_posible_kanban = IIf(IsNull(rs!INTE_ART_KANBAN), 0, rs!INTE_ART_KANBAN)
                        var_codigo_tela = Me.txt_codigo
                        var_nombre_articulo = IIf(IsNull(rs!vcha_Art_nombre_espa?ol), "", rs!vcha_Art_nombre_espa?ol)
                        var_numero_serie = IIf(IsNull(rs!inte_art_numero_serie), 0, rs!inte_art_numero_serie)
                        If var_cantidad_caja = 0 Then
                           If IsNull(rs(43).Value) Then
                              If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                                 var_recontable = 1
                              Else
                                 var_recontable = 0
                              End If
                           Else
                              If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                                 var_recontable = 1
                              Else
                                 var_recontable = rs(43).Value
                              End If
                           End If
                        Else
                           If var_empresa = "31" And var_clave_movimiento = "DCOM" Then
                              var_recontable = 1
                           Else
                              var_recontable = 0
                           End If
                        End If
                        rs.Close
                        If var_numero_serie = 1 Then
                           If var_recontable = 1 Then
                              var_cantidad_leida = 1#
                              lbl_Cantidad.Visible = True
                              txt_Cantidad.Visible = True
                           End If
                        Else
                           If var_recontable = 1 Then
                              var_cantidad_leida = 1#
                              lbl_Cantidad.Visible = True
                              txt_Cantidad.Visible = True
                              txt_Cantidad.SetFocus
                           Else
                              If var_cantidad_caja = 0 Then
                                 var_cantidad_leida = 1#
                              Else
                                 var_cantidad_leida = var_cantidad_caja
                              End If
                              txt_foco.Enabled = True
                              txt_foco.SetFocus
                           End If
                        End If
                     Else
                        txt_codigo = ""
                        frmmensaje.lbl_mensaje = "El art?culo no existe"
                        frmmensaje.Show 1
                     End If
                  Else
                     txt_codigo = ""
                     frmmensaje.lbl_mensaje = "El art?culo no existe"
                     frmmensaje.Show 1
                     rs.Close
                  End If
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub txt_codigo_LostFocus()
    var_codigo_seleccionado = ""
End Sub

Private Sub txt_factura_GotFocus()
   Me.txt_codigo.Enabled = False
End Sub

Private Sub txt_factura_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_factura) <> "" Then
         var_factura = txt_factura
         If var_clave_movimiento = "EC" Then
            If (var_empresa = "18" Or var_empresa = "06") And var_proveedor = "2458" Then
               rsaux.Open "SELECT * FROM TB_FACTURAS_ESTAMPADOS_REFUGIO WHERE VCHA_FAC_FACTURA = '" + Me.txt_factura + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux.EOF Then
                  'cnn_importacion.Open "Provider=OraOLEDB.Oracle.1;User ID=INTERFACE;Data Source=VENTAS;Extended Properties=;Persist Security Info=True;Password=INTERFACE"
                  cnn_importacion.Open "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=SIDTEXTILERA;Data Source=sqlquezada2"
                  cnn_importacion.CursorLocation = adUseClient
                  
                  'var_cadena = "SELECT DISTINCT cfaven.faccod, lfaven.FACSER, lfaven.facdsc descripcion, lfaven.facmts cantidad, (lfaven.facpremts + albrec.albrpre) precio,  albrec.albrpre tela, lfaven.faclin From cfaven@cipic.vianney.com.mx, lfaven@cipic.vianney.com.mx, barcad@cipic.vianney.com.mx, barpie@cipic.vianney.com.mx, albrec@cipic.vianney.com.mx, calprd@cipic.vianney.com.mx, clienv@cipic.vianney.com.mx Where cfaven.emprcod = lfaven.emprcod AND cfaven.faccod = lfaven.faccod AND cfaven.emprcod = clienv.emprcod"
                  'var_cadena = var_cadena + " AND cfaven.clicod = clienv.clicod AND lfaven.emprcod = barcad.emprcod AND lfaven.facbarcod = barcad.barcod aND lfaven.facbarreo = barcad.barcodreo AND lfaven.facbarpar = barcad.barcodpar AND lfaven.emprcod = calprd.emprcod AND lfaven.facalbcod = calprd.albprocod AND barcad.emprcod = barpie.emprcod aND barcad.barcod = barpie.barcod"
                  'var_cadena = var_cadena + " AND barcad.barcodreo = barpie.barcodreo AND barcad.barcodpar = barpie.barcodpar AND barpie.emprcod = albrec.emprcod AND barpie.albreccod = albrec.albreccod AND calprd.albdomenv = clienv.clienvlin AND cfaven.faccod = " + Me.txt_factura
                  
                  var_cadena = "SELECT INTE_CAR_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_COSTO+floa_sal_precio as floa_Sal_precio, VCHA_SAL_DESCRIPCION_FACTURA, FLOA_SAL_CANTIDAD From dbo.TB_SALIDAS WHERE     (VCHA_EMP_EMPRESA_ID = '15') AND (VCHA_CAR_DOCUMENTO = 'FA') AND (INTE_CAR_NUMERO = " + Me.txt_factura + ")"
                  
                  
                  rs.Open var_cadena, cnn_importacion, adOpenDynamic, adLockOptimistic
                  var_consecutivo = 0
                  While Not rs.EOF
                        var_consecutivo = var_consecutivo + 1
                        'rsaux2.Open "insert into tb_facturas_estampados_refugio (vcha_fac_factura, vcha_fac_codigo_externo, vcha_art_articulo_id, vcha_fac_descripcion, floa_fac_cantidad, floa_fac_COSTO, inte_fac_consecutivo) values ('" + Me.txt_factura + "', '" + CStr(rs!facser) + "','', '" + rs!descripcion + "'," + CStr(rs!Cantidad) + "," + CStr(rs!Precio) + ", " + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                        rsaux2.Open "insert into tb_facturas_estampados_refugio (vcha_fac_factura, vcha_fac_codigo_externo, vcha_art_articulo_id, vcha_fac_descripcion, floa_fac_cantidad, floa_fac_COSTO, inte_fac_consecutivo) values ('" + Me.txt_factura + "', '" + CStr(rs!VCHA_ART_ARTICULO_ID) + "','', '" + rs!vcha_sal_descripcion_factura + "'," + CStr(rs!floa_Sal_Cantidad) + "," + CStr(rs!floa_Sal_precio) + ", " + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                        rs.MoveNext
                  Wend
                  cnn_importacion.Close
               End If
               rsaux.Close
            End If
         End If
         If Me.txt_archivo <> "" Then
            If var_clave_movimiento = "EC" Then
               Me.txt_pedimento.Enabled = True
               var_costo_tela = 0
               Me.txt_pedimento.SetFocus
               txt_factura.Enabled = False
            Else
               txt_codigo.Enabled = True
               var_costo_tela = 0
               txt_codigo.SetFocus
               txt_factura.Enabled = False
            End If
         Else
            MsgBox "No se a indicado una orden de compra", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Debe de indicarse el n?mero de la factura", vbOKOnly, "ATENCION"
         txt_factura.SetFocus
      End If
   End If
End Sub


Private Sub txt_foco_GotFocus()
   Dim pError As ADODB.Error
   Dim var_codigo_barras_caja As String
   Dim var_actualiza As Boolean
   Dim var_inserta As Boolean
   Dim bandera_suma As Boolean
   Dim var_cantidad As Variant
   Dim var_costo As Variant
   Dim var_precio As Variant
   Dim var_consecutivo_serie  As Double
   Dim var_posible As Boolean
   Dim var_P_RC_LINEA_ID As Double
   Dim var_P_RC_NUMERO_LINEA As Double
   Set TB_ARCH_COMPARACION_M = New TB_ARCH_COMPARACION_M
   Set TB_ARCH_COMPARACION_I = New TB_ARCH_COMPARACION_I
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   Dim var_posible_lectura_kanban As Boolean
   'On Error GoTo salir:
   cnn.CommandTimeout = 360
   
   If Me.txt_codigo <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      If rsaux.State = 1 Then
         rsaux.Close
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
   End If
   
   If var_clave_movimiento = "EC" Then
      If Me.txt_codigo <> "" Then
         If (var_empresa = "18" Or var_empresa = "06") And var_proveedor = "2458" Then
            ' AQUI VA EL CODIGO_MAL
            var_codigo_tela = Me.txt_codigo
            rsaux10.Open "select * from tb_equivalencias where vcha_art_Articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux10.EOF Then
               rsaux11.Open "SELECT * FROM TB_FACTURAS_ESTAMPADOS_REFUGIO WHERE VCHA_FAC_FACTURA = '" + Me.txt_factura + "' and vcha_fac_codigo_externo = '" + IIf(IsNull(rsaux10!vcha_equ_codigo_equivalente), "", rsaux10!vcha_equ_codigo_equivalente) + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux11.EOF Then
                  var_costo_tela = IIf(IsNull(rsaux11!floa_fac_costo), 0, rsaux11!floa_fac_costo)
               Else
                  frmfactura_orden_compra_estampados.lbl_articulo = var_nombre_articulo
                  frmfactura_orden_compra_estampados.txt_factura = var_factura
                  frmfactura_orden_compra_estampados.Show 1
                  Me.txt_codigo.SetFocus
                  Me.txt_codigo = var_codigo_tela
               End If
               rsaux11.Close
            Else
               frmfactura_orden_compra_estampados.lbl_articulo = var_nombre_articulo
               frmfactura_orden_compra_estampados.txt_factura = var_factura
               frmfactura_orden_compra_estampados.Show 1
               Me.txt_codigo = var_codigo_tela
            End If
            rsaux10.Close
         End If
      End If
   End If
   
   
   var_codigo_seleccionado = ""
   If var_posible_kanban = 1 Then
      var_global_aceptar_demas = 0
   End If
   If var_empresa <> "18" Then
      If var_empresa <> "06" Then
         var_costo_tela = 0
      End If
   End If
   If var_clave_movimiento <> "EC" Then
      var_costo_tela = 0
   End If
   If Trim(txt_codigo.Text) <> "" Then
      lv_entradas.Font.Bold = False
      bandera_suma = False
      If var_primera_vez = True Then
         rsaux11.Open "UPDATE TB_ARCHIVO_COMPARACION SET VCHA_COM_TIPO_LECTURA = 'M' WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND VCHA_COM_REFERENCIA = '" + Me.txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
         If var_tipo_documento = "V" Then
            var_inserta = False
            var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, var_folio_enviado, "", var_proveedor, var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, var_factura, "", txt_archivo, "", "B", "", "", 0, 0, 0, var_clave_moneda, var_tipo_Cambio)
            var_numero_folio = var_numero_folio_regreso
         Else
            var_inserta = False
            If var_clave_movimiento = "EI" Then
               var_serie = ""
               var_numero_factura = ""
               For var_j = 1 To Len(Me.txt_archivo)
                   If Not IsNumeric(Mid(Me.txt_archivo, var_j, 1)) Then
                      var_serie = var_serie + Mid(Me.txt_archivo, var_j, 1)
                   Else
                      var_numero_factura = var_numero_factura + Mid(Me.txt_archivo, var_j, 20)
                      Exit For
                   End If
               Next var_j
               var_factura = var_numero_factura
            End If
            var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, var_folio_enviado, "", var_proveedor, var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, var_factura, "", txt_archivo, "", "B", "", "", 0, 0, 0, var_clave_moneda, var_tipo_Cambio)
            var_numero_folio = var_numero_folio_regreso
         End If
         txt_folio = var_numero_folio
         rsaux10.Open "update tb_encabezado_movimientos set VCHA_EMO_PEDIMENTO = '" + Me.txt_pedimento + "', vcha_emo_referencia_transito = '" + Me.lbl_transito + "' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
         var_primera_vez = False
         var_fecha_movimiento = Date
      End If
      var_posible = True
      If var_cajas = True Then
         var_posible = True
      Else
         rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_posible = True
         Else
            var_posible = False
         End If
         rsaux.Close
      End If
      If var_posible = True Then
         If var_cajas = True Then
            Cadena = "select * from tb_archivo_comparacion where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_com_referencia = '"
            Cadena = Cadena + txt_archivo + "' and vcha_com_caja = '" + txt_codigo + "'"
         Else
            Cadena = "select * from tb_archivo_comparacion where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_com_referencia = '"
            Cadena = Cadena + txt_archivo + "' and vcha_art_articulo_id = '" + txt_codigo + "'"
         End If
         If rs.State = 1 Then
            rs.Close
         End If
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            valor = txt_codigo
            var_n = lv_entradas.ListItems.Count
            var_encontro = 0
            var_i = 1
            If var_cajas = True Then
               While (var_i <= var_n)
                     lv_entradas.ListItems.item(var_i).Selected = True
                     valor = Trim(lv_entradas.selectedItem.SubItems(11))
                     If txt_codigo = valor Then
                        var_encontro = 1
                        var_i = var_n + 1
                     End If
                     var_i = var_i + 1
               Wend
            Else
               'MsgBox var_empresa
               If var_empresa = "16" Then
                  While (var_i <= var_n)
                        lv_entradas.ListItems.item(var_i).Selected = True
                        valor = Trim(lv_entradas.selectedItem)
                        If txt_codigo = valor And CDbl(lv_entradas.selectedItem.SubItems(2)) = CDbl(var_cantidad_leida) And CDbl(lv_entradas.selectedItem.SubItems(4)) = 0 And CDbl(lv_entradas.selectedItem.SubItems(5)) > 0 Then
                           var_encontro = 1
                           var_i = var_n + 1
                        Else
                           var_encontro = 0
                        End If
                        var_i = var_i + 1
                  Wend
               Else
                  While (var_i <= var_n)
                        lv_entradas.ListItems.item(var_i).Selected = True
                        valor = Trim(lv_entradas.selectedItem)
                        If txt_codigo = valor Then
                           var_cantidad_posible = lv_entradas.selectedItem.SubItems(2)
                           If var_cantidad_posible < lv_entradas.selectedItem.SubItems(3) + var_cantidad_leida Then
                              var_encontro = 0
                           Else
                              var_encontro = 1
                              var_i = var_n + 1
                           End If
                        End If
                        var_i = var_i + 1
                  Wend
               End If
            End If
            If var_encontro = 1 Then
               If var_cajas = True Then
                  var_codigo_barras_caja = txt_codigo
                  txt_codigo = var_codigo_caja
               End If
               bandera_suma = True
               var_posible = True
               If var_tipo_documento = "V" Then
                  If (lv_entradas.selectedItem.SubItems(2) * 1) < ((lv_entradas.selectedItem.SubItems(3) * 1) + var_cantidad_leida) Then
                     var_posible = False
                  End If
               Else
                  var_posible = True
               End If
               If var_posible = True Then
                  'If var_posible_kanban = 1 Then
                  '   var_global_aceptar_demas = 0
                  'Else
                  '   var_global_aceptar_demas = 1
                  'End If
                  If var_global_aceptar_demas = 0 Then
                     If (lv_entradas.selectedItem.SubItems(2) * 1) < ((lv_entradas.selectedItem.SubItems(3) * 1) + var_cantidad_leida) Then
                        var_posible = False
                     End If
                  End If
                  If var_posible = True Then
                     var_posible_lectura_kanban = True
                     If var_posible_kanban = 1 Then
                        Set TB_RESERVAR_KANBAN_ENTRADA = New TB_RESERVAR_KANBAN_ENTRADA
                        If var_kanban_es_un_kanban = "S" Then
                           var_inserta = TB_RESERVAR_KANBAN_ENTRADA.Anadir(var_kanban, var_clave_movimiento, var_numero_folio, var_almacen_Destino, Me.txt_codigo, Me.txt_archivo, "", "")
                           If var_kanban_exito = "S" Then
                              var_posible_lectura_kanban = True
                           Else
                              var_posible_lectura_kanban = False
                           End If
                        Else
                           'Set TB_RESERVAR_FUERA_KANBAN_ENT = New TB_RESERVAR_FUERA_KANBAN_ENT
                           'var_inserta = TB_RESERVAR_FUERA_KANBAN_ENT.Anadir(Me.txt_archivo, var_clave_movimiento, var_numero_folio, var_almacen_Destino, Me.txt_codigo, "", "")
                           var_posible_lectura_kanban = True
                        End If
                     Else
                        var_kanban_mensaje = ""
                        var_posible_lectura_kanban = True
                     End If
                     If var_posible_lectura_kanban = True Then
                        lv_entradas.selectedItem.Selected = True
                        lv_entradas.selectedItem.EnsureVisible
                        lv_entradas.selectedItem.SubItems(3) = Format(lv_entradas.selectedItem.SubItems(3) + var_cantidad_leida, "###,###,##0.00")
                        lv_entradas.selectedItem.SubItems(4) = Format(lv_entradas.selectedItem.SubItems(4) + var_cantidad_leida, "###,###,##0.00")
                        lv_entradas.selectedItem.SubItems(5) = Format(lv_entradas.selectedItem.SubItems(2) - lv_entradas.selectedItem.SubItems(3), "###,###,##0.00")
                        var_consecutivo = lv_entradas.selectedItem.SubItems(9)
                        If var_clave_movimiento = "EI" Or var_clave_movimiento = "ETA" Then
                           If var_empresa = "06" Or var_empresa = "18" Then
                              var_costo = CDbl(lv_entradas.selectedItem.SubItems(6)) + CDbl(lv_entradas.selectedItem.SubItems(16))
                           Else
                              var_costo = CDbl(lv_entradas.selectedItem.SubItems(6))
                           End If
                        Else
                           var_costo = lv_entradas.selectedItem.SubItems(6)
                        End If
                        var_precio = lv_entradas.selectedItem.SubItems(16)
                        var_cantidad = lv_entradas.selectedItem.SubItems(4)
                        If lv_entradas.selectedItem.SubItems(10) <> "" Then
                           var_a?o = lv_entradas.selectedItem.SubItems(10)
                        Else
                           var_a?o = 2005
                        End If
                     
                        If IsNumeric(lv_entradas.selectedItem.SubItems(14)) Then
                           VAR_RC_NL = lv_entradas.selectedItem.SubItems(14) * 1
                        Else
                           VAR_RC_NL = 0
                        End If
                        If IsNumeric(lv_entradas.selectedItem.SubItems(13)) Then
                           VAR_RC_LINEA_ID = lv_entradas.selectedItem.SubItems(13) * 1
                        Else
                           VAR_RC_LINEA_ID = 0
                        End If
                        var_P_RC_NUMERO_LINEA = VAR_RC_NL
                        var_P_RC_LINEA_ID = VAR_RC_LINEA_ID
                        lbl_recibidos = Format(CDbl(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                        var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                        cnn.BeginTrans
                        var_actualiza = TB_ARCH_COMPARACION_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), var_tipo_proveedor, var_origen, txt_codigo, var_cantidad_leida, txt_archivo, var_consecutivo)
                     Else
                        txt_codigo = ""
                        If var_kanban_mensaje = "" Then
                           frmmensaje.lbl_mensaje = "Debe de leer kanbans"
                        Else
                           frmmensaje.lbl_mensaje = var_kanban_mensaje
                        End If
                        frmmensaje.Show 1
                        GoTo salir:
                     End If
                  Else
                     txt_codigo = ""
                     frmmensaje.lbl_mensaje = "La cantidad exede a la cantidad en la relaci?n"
                     frmmensaje.Show 1
                     'MsgBox "La cantidad exede a la cantidad en la relaci?n", vbOKOnly, "ATENCION"
                  End If
               Else
                  txt_codigo = ""
                  frmmensaje.lbl_mensaje = "La cantidad exede a la incluida en la salida a vistas"
                  frmmensaje.Show 1
                  'MsgBox "La cantidad exede a la incluida en la salida a vistas", vbOKOnly, "ATENCION"
               End If
            Else
               valor = txt_codigo
               Set itmfound = lv_entradas.findItem(valor, lvwText, , lvwPartial)
               itmfound.EnsureVisible
               itmfound.Selected = True
               bandera_suma = True
               var_posible = True
               If var_tipo_documento = "V" Then
                  If (lv_entradas.selectedItem.SubItems(2) * 1) < ((lv_entradas.selectedItem.SubItems(3) * 1) + var_cantidad_leida) Then
                     var_posible = False
                  End If
               Else
                  var_posible = True
               End If
               If var_posible = True Then
                  If var_global_aceptar_demas = 0 Then
                     If Round(CDbl((lv_entradas.selectedItem.SubItems(2) * 1)), 4) < Round(CDbl(((lv_entradas.selectedItem.SubItems(3) * 1) + var_cantidad_leida)), 4) Then
                        var_posible = False
                     End If
                  End If
                  If var_posible_kanban = 1 Then
                     If var_kanban_es_un_kanban = "S" Then
                        Set TB_RESERVAR_KANBAN_ENTRADA = New TB_RESERVAR_KANBAN_ENTRADA
                        If var_kanban_es_un_kanban = "S" Then
                           var_inserta = TB_RESERVAR_KANBAN_ENTRADA.Anadir(var_kanban, var_clave_movimiento, var_numero_folio, var_almacen_Destino, Me.txt_codigo, Me.txt_archivo, "", "")
                           If var_kanban_exito = "S" Then
                              var_posible_lectura_kanban = True
                           Else
                              
                              var_posible_lectura_kanban = False
                              var_posible = False
                              txt_codigo = ""
                              frmmensaje.lbl_mensaje = var_kanban_mensaje
                              frmmensaje.Show 1
                           
                           End If
                        Else
                           'Set TB_RESERVAR_FUERA_KANBAN_ENT = New TB_RESERVAR_FUERA_KANBAN_ENT
                           'var_inserta = TB_RESERVAR_FUERA_KANBAN_ENT.Anadir(Me.txt_archivo, var_clave_movimiento, var_numero_folio, var_almacen_Destino, Me.txt_codigo, "", "")
                           var_posible_lectura_kanban = True
                        End If
                     Else
                        txt_codigo = ""
                        frmmensaje.lbl_mensaje = "Se deben de leer etiquetas Kanban"
                        frmmensaje.Show 1
                     End If
                  Else
                     var_posible_lectura_kanban = True
                  End If
                  If var_posible_lectura_kanban = True Then
                     If var_posible = True Then
                        lv_entradas.selectedItem.Selected = True
                        lv_entradas.selectedItem.EnsureVisible
                        lv_entradas.selectedItem.SubItems(3) = Format(lv_entradas.selectedItem.SubItems(3) + var_cantidad_leida, "###,###,##0.00")
                        lv_entradas.selectedItem.SubItems(4) = Format(lv_entradas.selectedItem.SubItems(4) + var_cantidad_leida, "###,###,##0.00")
                        lv_entradas.selectedItem.SubItems(5) = Format(lv_entradas.selectedItem.SubItems(2) - lv_entradas.selectedItem.SubItems(3), "###,###,##0.00")
                        var_consecutivo = lv_entradas.selectedItem.SubItems(9)
                        If var_clave_movimiento = "EI" Then
                           If var_empresa = "06" Or var_empresa = "18" Then
                              var_costo = CDbl(lv_entradas.selectedItem.SubItems(6)) + CDbl(lv_entradas.selectedItem.SubItems(16))
                           Else
                              var_costo = CDbl(lv_entradas.selectedItem.SubItems(6))
                           End If
                        End If
                        var_precio = lv_entradas.selectedItem.SubItems(7)
                        var_cantidad = lv_entradas.selectedItem.SubItems(4)
                        lbl_recibidos = Format(CDbl(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                        var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                        cnn.BeginTrans
                        var_actualiza = TB_ARCH_COMPARACION_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), var_tipo_proveedor, var_origen, txt_codigo, var_cantidad_leida, txt_archivo, var_consecutivo)
                        var_renglon = var_n
                     Else
                        txt_codigo = ""
                        frmmensaje.lbl_mensaje = "La cantidad exede a la cantidad en la relaci?n"
                        frmmensaje.Show 1
                        'MsgBox "La cantidad exede a la cantidad en la relaci?n", vbOKOnly, "ATENCION"
                     End If
                  End If
               Else
                  txt_codigo = ""
                  frmmensaje.lbl_mensaje = "La cantidad exede a la incluida en la salida a vistas"
                  frmmensaje.Show 1
                  'MsgBox "La cantidad exede a la incluida en la salida a vistas", vbOKOnly, "ATENCION"
               End If
            End If
         Else
            var_posible = True
            If var_tipo_documento = "V" Then
               txt_codigo = ""
               frmmensaje.lbl_mensaje = "El art?culo no existe dentro de la salida a vistas"
               frmmensaje.Show 1
               'MsgBox "El art?culo no existe dentro de la salida a vistas", vbOKOnly, "ATENCION"
               var_posible = False
            Else
               If var_global_aceptar_demas = 1 Then
                  rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     rsaux3.Open "select max(inte_com_consecutivo) as maximo from tb_Archivo_comparacion where vcha_emp_empresa_id = '" + var_empresa + "' and  vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "'  and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "'  and INTE_COM_NUMERO = " + CStr(CDbl(var_folio_enviado)) + " and VCHA_COM_REFERENCIA = '" + txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux3.EOF Then
                        var_consecutivo = IIf(IsNull(rsaux3!maximo), 0, rsaux3!maximo) + 1
                     Else
                        var_consecutivo = 1
                     End If
                     rsaux3.Close
                     Set list_item = lv_entradas.ListItems.Add(, , rsaux(0).Value)
                     list_item.SubItems(1) = IIf(IsNull(rsaux(1).Value), "", rsaux(1).Value)
                     list_item.SubItems(2) = Format(0, "###,###,##0.00")
                     list_item.SubItems(3) = Format(var_cantidad_leida, "###,###,##0.00")
                     list_item.SubItems(4) = Format(var_cantidad_leida, "###,###,##0.00")
                     list_item.SubItems(5) = Format(list_item.SubItems(2) - list_item.SubItems(3), "###,###,##0.00")
                     list_item.SubItems(6) = IIf(IsNull(rsaux(3).Value), "", rsaux(3).Value)
                     list_item.SubItems(7) = IIf(IsNull(rsaux(2).Value), "", rsaux(2).Value)
                     list_item.SubItems(8) = 0
                     list_item.SubItems(9) = var_consecutivo
                     list_item.SubItems(10) = 2005
                     var_n = lv_entradas.ListItems.Count
                     lv_entradas.ListItems.item(var_n).Selected = True
                     var_precio = rsaux(2).Value
                     var_costo = rsaux!mone_Art_costo_estandar
                     
                     If var_entrada_calidad = True Then
                        rsaux2.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_costeo + "' and vcha_art_articulo_id = '" + Trim(txt_codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           var_costo = rsaux2!FLOA_eXI_COSTO
                        End If
                        rsaux2.Close
                     End If
                     bandera_suma = True
                     lbl_recibidos = Format(CDbl(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                     var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                     cnn.BeginTrans
                     If var_cajas = True Then
                        ok = TB_ARCH_COMPARACION_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), Date, var_tipo_proveedor, var_origen, txt_codigo, var_costo, 0, var_cantidad_leida, var_transporto, txt_archivo, 0, var_consecutivo, 2005, txt_codigo, var_peso_caja)
                        var_posible_lectura_kanban = True
                     Else
                        ok = TB_ARCH_COMPARACION_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), Date, var_tipo_proveedor, var_origen, txt_codigo, var_costo, 0, var_cantidad_leida, var_transporto, txt_archivo, 0, var_consecutivo, 2005, "", 0)
                        var_posible_lectura_kanban = True
                     End If
                     var_renglon = lv_entradas.ListItems.Count
                  Else
                     txt_codigo = ""
                     frmmensaje.lbl_mensaje = "El art?culo no existe"
                     frmmensaje.Show 1
                     'MsgBox "El art?culo no existe", vbOKOnly, "ATENCION"
                     bandera_suma = False
                  End If
                  rsaux.Close
               Else
                  txt_codigo = ""
                  frmmensaje.lbl_mensaje = "El art?culo no se encuentra dentro de la relaci?n"
                  frmmensaje.Show 1
                  'MsgBox "El art?culo no se encuentra dentro de la relaci?n", vbOKOnly, "ATENCION"
               End If
            End If
            If var_global_aceptar_demas = 0 Then
               var_posible = False
            Else
               var_posible = True
            End If
         End If
         If rs.State = 1 Then
            rs.Close
         End If
      Else
         txt_codigo = ""
         frmmensaje.lbl_mensaje = "El art?culo no existe"
         frmmensaje.Show 1
         'MsgBox "El art?culo no existe", vbOKOnly, "ATENCION"
         rsaux.Close
      End If
      If bandera_suma = True Then
         If var_tipo_documento = "V" Then
            If var_posible = True Then
               
               Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "' and inte_ent_consecutivo = " + CStr(var_consecutivo)
               rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_inserta = False
                  rsaux3.Open "UPDATE TB_TEMPORAL_ENTRADAS SET VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "', FLOA_ENT_CANTIDAD = FLOA_ENT_CANTIDAD + " + CStr(var_cantidad_leida) + " Where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_ENT_NUMERO= '" + CStr(var_numero_folio) + "' AND VCHA_ART_ARTICULO_ID  = '" + txt_codigo + "' and inte_ent_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  rs.Close
               Else
                  var_inserta = False
                  If var_empresa = "18" And var_proveedor = "2458" Then
                     rsaux3.Open "INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN,INTE_ENT_CONSECUTIVO, INTE_ENT_A?O, P_RC_NUMERO_LINEA, P_RC_LINEA_ID) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo_tela) + ", " + CStr(var_precio) + ", 0,'" + var_almacen_origen + "', " + CStr(var_consecutivo) + ", " + CStr(var_a?o) + ", " + CStr(var_P_RC_LINEA_ID) + "," + CStr(var_P_RC_NUMERO_LINEA) + ")", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux3.Open "INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN,INTE_ENT_CONSECUTIVO, INTE_ENT_A?O, P_RC_NUMERO_LINEA, P_RC_LINEA_ID) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo + var_costo_tela) + ", " + CStr(var_precio) + ", 0,'" + var_almacen_origen + "', " + CStr(var_consecutivo) + ", " + CStr(var_a?o) + ", " + CStr(var_P_RC_LINEA_ID) + "," + CStr(var_P_RC_NUMERO_LINEA) + ")", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rs.Close
               End If
               If var_numero_serie = 1 Then
                  Cadena = "select MAX(INTE_EXI_CONSECUTIVO) from TB_EXISTENCIAS_SERIES WHERE vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_eMO_numero = " + Str(var_numero_folio)
                  var_consecutivo_serie = 0
                  rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_consecutivo_serie = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                  Else
                     var_consecutivo_serie = 1
                  End If
                  rs.Close
                  rsaux.Open "insert into tb_exiStencias_series INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, VCHA_ART_ARTICULO_ID, VCHA_ART_NUMERO_SERIE, INTE_EXI_CONSECUTIVO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', '" + Me.txt_numero_serie + "'," + CStr(var_consecutivo_serie) + ")", cnn, adOpenDynamic, adLockOptimistic
               End If
            End If
         Else
            If var_posible = True Then
               var_posible_lectura_kanban = True
               If var_posible_lectura_kanban = True Then
                  If var_posible_kanban = 1 Then
                     If var_kanban_es_un_kanban = "N" Or var_kanban_es_un_kanban = "" Then
                        Set TB_RESERVAR_FUERA_KANBAN_ENT = New TB_RESERVAR_FUERA_KANBAN_ENT
                        var_inserta = TB_RESERVAR_FUERA_KANBAN_ENT.Anadir(Me.txt_archivo, var_clave_movimiento, var_numero_folio, var_almacen_Destino, Me.txt_codigo, "", "")
                        If var_kanban_exito = "N" Then
                           var_posible = False
                        End If
                     End If
                  End If
                  If var_posible = True Then
                     'cnn.BeginTrans
                     Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "' and inte_ent_consecutivo = " + CStr(var_consecutivo)
                     rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_inserta = False
                        rsaux3.Open "UPDATE TB_TEMPORAL_ENTRADAS SET VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "', FLOA_ENT_CANTIDAD = FLOA_ENT_CANTIDAD + " + CStr(var_cantidad_leida) + " Where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_ENT_NUMERO= '" + CStr(var_numero_folio) + "' AND VCHA_ART_ARTICULO_ID  = '" + txt_codigo + "' and inte_ent_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                        rs.Close
                        If var_clave_movimiento = "ETA" Or var_clave_movimiento = "EI" Or var_clave_movimiento = "EP" Then
                           '4
                           rsaux3.Open "update tb_transito set floa_tra_cantidad_recibida = isnull(floa_Tra_cantidad_recibida,0) + " + CStr(var_cantidad_leida) + " where vcha_tra_nota_envio = '" + Me.lbl_transito + "' and vcha_Art_articulo_recivo = '" + Me.txt_codigo + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
                        End If
                     Else
                        var_inserta = False
                        var_costo = IIf(IsNull(var_costo), 0, var_costo)
                        If var_costo = 0 Or var_clave_movimiento = "DT" Then
                           rsaux4.Open "select floa_exi_costo_2005, floa_Exi_costo from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux4.EOF Then
                              var_costo = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                           Else
                              var_costo = 0
                           End If
                           rsaux4.Close
                           If var_costo = 0 And var_almacen_Destino = "14" Then
                              rsaux4.Open "select floa_exi_costo_2005, floa_Exi_costo from tb_existencias where vcha_alm_almacen_id = '8' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux4.EOF Then
                                 var_costo = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                              Else
                                 var_costo = 0
                              End If
                              rsaux4.Close
                           End If
                           If var_costo = 0 And var_almacen_Destino = "11" Then
                              rsaux4.Open "select floa_exi_costo_2005, floa_Exi_costo from tb_existencias where vcha_alm_almacen_id = '8' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux4.EOF Then
                                 var_costo = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                              Else
                                 var_costo = 0
                              End If
                              rsaux4.Close
                           End If
                           If var_costo = 0 Then
                              rsaux4.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux4.EOF Then
                                 var_costo = IIf(IsNull(rsaux4!mone_Art_costo_estandar), 0, rsaux4!mone_Art_costo_estandar)
                              End If
                              rsaux4.Close
                           End If
                        End If
                        If var_empresa = "18" And var_proveedor = "2458" Then
                           rsaux3.Open "INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN, INTE_ENT_CONSECUTIVO, INTE_ENT_A?O, P_RC_LINEA_ID, P_RC_NUMERO_LINEA) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo_tela) + ", " + CStr(var_precio) + ", 0,'" + var_almacen_origen + "', " + CStr(var_consecutivo) + ", " + CStr(var_a?o) + "," + CStr(var_P_RC_LINEA_ID) + "," + CStr(var_P_RC_NUMERO_LINEA) + ")", cnn, adOpenDynamic, adLockOptimistic
                        Else
                           If var_empresa = "06" And var_proveedor = "2458" Then
                              rsaux3.Open "INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN, INTE_ENT_CONSECUTIVO, INTE_ENT_A?O, P_RC_LINEA_ID, P_RC_NUMERO_LINEA) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo_tela) + ", " + CStr(var_precio) + ", 0,'" + var_almacen_origen + "', " + CStr(var_consecutivo) + ", " + CStr(var_a?o) + "," + CStr(var_P_RC_LINEA_ID) + "," + CStr(var_P_RC_NUMERO_LINEA) + ")", cnn, adOpenDynamic, adLockOptimistic
                           Else
                              If var_precio = "" Then
                                 var_precio = 0
                              End If

                              rsaux3.Open "INSERT INTO TB_TEMPORAL_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN, INTE_ENT_CONSECUTIVO, INTE_ENT_A?O, P_RC_LINEA_ID, P_RC_NUMERO_LINEA, floa_sal_cantidad_metros) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo + var_costo_tela) + ", " + CStr(var_precio) + ", 0,'" + var_almacen_origen + "', " + CStr(var_consecutivo) + ", " + CStr(var_a?o) + "," + CStr(var_P_RC_LINEA_ID) + "," + CStr(var_P_RC_NUMERO_LINEA) + "," + CStr(CDbl(Me.lv_entradas.selectedItem.SubItems(2))) + ")", cnn, adOpenDynamic, adLockOptimistic
                              If var_clave_movimiento = "ETA" Or var_clave_movimiento = "EP" Or var_clave_movimiento = "EI" Then
                                 '1
                                 rsaux3.Open "update tb_transito set floa_tra_cantidad_recibida = isnull(floa_Tra_cantidad_recibida,0) + " + CStr(var_cantidad_leida) + " where vcha_tra_nota_envio = '" + Me.lbl_transito + "' and vcha_Art_articulo_recivo = '" + Me.txt_codigo + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
                              End If
                           End If
                        End If
                        rs.Close
                     End If
                     cnn.CommitTrans
                  Else
                  'cuando no es kanban
                     lv_entradas.selectedItem.Selected = True
                     lv_entradas.selectedItem.EnsureVisible
                     lv_entradas.selectedItem.SubItems(3) = Format(lv_entradas.selectedItem.SubItems(3) - var_cantidad_leida, "###,###,##0.00")
                     lv_entradas.selectedItem.SubItems(4) = Format(lv_entradas.selectedItem.SubItems(4) - var_cantidad_leida, "###,###,##0.00")
                     lv_entradas.selectedItem.SubItems(5) = Format(lv_entradas.selectedItem.SubItems(5) + var_cantidad_leida, "###,###,##0.00")
                     var_consecutivo = lv_entradas.selectedItem.SubItems(9)
                     var_costo = lv_entradas.selectedItem.SubItems(6)
                     var_precio = lv_entradas.selectedItem.SubItems(7)
                     var_cantidad = lv_entradas.selectedItem.SubItems(4)
                     lbl_recibidos = Format(CDbl(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                     var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                     var_actualiza = TB_ARCH_COMPARACION_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), var_tipo_proveedor, var_origen, txt_codigo, 0 - var_cantidad_leida, txt_archivo, var_consecutivo)
                     var_renglon = var_n
                     'fin de cuando no es un kanban
                     frmmensaje.lbl_mensaje = var_kanban_mensaje
                     frmmensaje.Show 1
                  End If
                  
               End If 'kanban
               If var_numero_serie = 1 Then
                  'Cadena = "select MAX(INTE_EXI_CONSECUTIVO) from TB_EXISTENCIAS_SERIES WHERE vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_eMO_numero = " + Str(var_numero_folio)
                  Cadena = "select MAX(INTE_EXI_CONSECUTIVO) from TB_EXISTENCIAS_SERIES WHERE VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'"
                  var_consecutivo_serie = 0
                  rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_consecutivo_serie = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                  Else
                     var_consecutivo_serie = 1
                  End If
                  rs.Close
                  var_consecutivo_serie_str = CStr(var_consecutivo_serie)
                  If Len(var_consecutivo_serie_str) = 1 Then
                     var_consecutivo_serie_str = "000" + Trim(var_consecutivo_serie_str)
                  Else
                     If Len(var_consecutivo_serie_str) = 2 Then
                        var_consecutivo_serie_str = "00" + Trim(var_consecutivo_serie_str)
                     Else
                        If Len(var_consecutivo_serie_str) = 3 Then
                           var_consecutivo_serie_str = "0" + Trim(var_consecutivo_serie_str)
                        Else
                           If Len(var_consecutivo_serie_str) = 4 Then
                              var_consecutivo_serie_str = Trim(var_consecutivo_serie_str)
                           Else
                           End If
                        End If
                     End If
                  End If
                  Me.txt_numero_serie = Me.txt_codigo + var_consecutivo_serie_str
                  rsaux.Open "insert into tb_exiStencias_series (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, VCHA_ART_ARTICULO_ID, VCHA_ART_NUMERO_SERIE, INTE_EXI_CONSECUTIVO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ",'" + txt_codigo + "', '" + Me.txt_numero_serie + "'," + CStr(var_consecutivo_serie) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rsaux9.Open "select substring(vcha_art_nombre_espa?ol,1,23) AS descripcion, substring(vcha_art_nombre_espa?ol,24,46) as descripcion_2 from tb_articulos where vcha_art_Articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux9.EOF Then
                     var_descripcion_etiqueta = IIf(IsNull(rsaux9!descripcion), "", rsaux9!descripcion)
                     var_descripcion_etiqueta2 = IIf(IsNull(rsaux9!descripcion_2), "", rsaux9!descripcion_2)
                  End If
                  rsaux9.Close
                  Open (App.Path & "\etiqueta.bat") For Output As #2
                  Print #2, "copy " + App.Path + "\etiqueta.txt lpt1"
                  Open (App.Path & "\etiqueta.txt") For Output As #1
                  Close #2
                  Print #1, "US"
                  Print #1, "q392"
                  Print #1, "Q256,24+0"
                  Print #1, "S2"
                  Print #1, "D8"
                  Print #1, "ZT"
                  Print #1, "TTh: m"
                  Print #1, "TDy2.mn.dd"
                  Print #1, "A39,20,0,3,1,1,N,""" + var_descripcion_etiqueta + """"
                  Print #1, "A39,40,0,3,1,1,N,""" + var_descripcion_etiqueta2 + """"
                  Print #1, "B39,90,0,3,2,4,101,B,""" + Trim(Me.txt_numero_serie) + """"
                  Print #1, "P1"
                  Close #1
                  x = Shell(App.Path & "\etiqueta.bat", vbHide)
               End If
            End If
         End If
         bandera_suma = False
         var_renglon = lv_entradas.selectedItem.Index
         Call ilumina_grid
      End If
      If var_n > 11 Then
         lv_entradas.ColumnHeaders(2).Width = 4700.01
      Else
         lv_entradas.ColumnHeaders(2).Width = 4930.01
      End If
      If var_cajas = True Then
         If var_peso_correcto = True Then
            var_costo_tela = 0
            Me.txt_codigo.SetFocus
         Else
            txt_codigo = var_codigo_barras_caja
            Me.txt_codigo_caja.SetFocus
         End If
      Else
         var_costo_tela = 0
         Me.txt_codigo = ""
         txt_codigo.SetFocus
      End If
   End If
   Exit Sub
    
salir:
   'MsgBox CStr(Err.Number)
   If Err.Number = -2147168237 Then
      'MsgBox Err.Description
      cnn.CommitTrans
      Resume
   Else
      If Err.Number = -2147217871 Then
         Resume
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
      MsgBox "Se a generado un error y el sistema se cerrara", vbOKOnly, "ATENCION"
      End
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
   End If
End Sub


Sub ilumina_grid()
   var_n = lv_entradas.ListItems.Count
   For var_i = 1 To var_n
       If var_i = var_renglon Then
          lv_entradas.ListItems.item(var_i).Bold = True
          lv_entradas.ListItems.item(var_i).ListSubItems(1).Bold = True
          lv_entradas.ListItems.item(var_i).ListSubItems(2).Bold = True
          lv_entradas.ListItems.item(var_i).ListSubItems(3).Bold = True
          lv_entradas.ListItems.item(var_i).ListSubItems(4).Bold = True
          lv_entradas.ListItems.item(var_i).ListSubItems(5).Bold = True
          lv_entradas.ListItems.item(var_i).ListSubItems(6).Bold = True
          lv_entradas.ListItems.item(var_i).ListSubItems(7).Bold = True
          lv_entradas.ListItems.item(var_i).ForeColor = &H8000&
          lv_entradas.ListItems.item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_entradas.ListItems.item(var_i).ListSubItems(2).ForeColor = &H8000&
          lv_entradas.ListItems.item(var_i).ListSubItems(3).ForeColor = &H8000&
          lv_entradas.ListItems.item(var_i).ListSubItems(4).ForeColor = &H8000&
          lv_entradas.ListItems.item(var_i).ListSubItems(5).ForeColor = &H8000&
          lv_entradas.ListItems.item(var_i).ListSubItems(6).ForeColor = &H8000&
          lv_entradas.ListItems.item(var_i).ListSubItems(7).ForeColor = &H8000&
       Else
          If (lv_entradas.ListItems.item(var_i).ListSubItems(5) * 1) < 0 Then
             lv_entradas.ListItems.item(var_i).Bold = True
             lv_entradas.ListItems.item(var_i).ListSubItems(1).Bold = True
             lv_entradas.ListItems.item(var_i).ListSubItems(2).Bold = True
             lv_entradas.ListItems.item(var_i).ListSubItems(3).Bold = True
             lv_entradas.ListItems.item(var_i).ListSubItems(4).Bold = True
             lv_entradas.ListItems.item(var_i).ListSubItems(5).Bold = True
             lv_entradas.ListItems.item(var_i).ListSubItems(6).Bold = True
             lv_entradas.ListItems.item(var_i).ListSubItems(7).Bold = True
             lv_entradas.ListItems.item(var_i).ForeColor = &HFF&
             lv_entradas.ListItems.item(var_i).ListSubItems(1).ForeColor = &HFF&
             lv_entradas.ListItems.item(var_i).ListSubItems(2).ForeColor = &HFF&
             lv_entradas.ListItems.item(var_i).ListSubItems(3).ForeColor = &HFF&
             lv_entradas.ListItems.item(var_i).ListSubItems(4).ForeColor = &HFF&
             lv_entradas.ListItems.item(var_i).ListSubItems(5).ForeColor = &HFF&
             lv_entradas.ListItems.item(var_i).ListSubItems(6).ForeColor = &HFF&
             lv_entradas.ListItems.item(var_i).ListSubItems(7).ForeColor = &HFF&
          Else
             lv_entradas.ListItems.item(var_i).Bold = False
             lv_entradas.ListItems.item(var_i).ListSubItems(1).Bold = False
             lv_entradas.ListItems.item(var_i).ListSubItems(2).Bold = False
             lv_entradas.ListItems.item(var_i).ListSubItems(3).Bold = False
             lv_entradas.ListItems.item(var_i).ListSubItems(4).Bold = False
             lv_entradas.ListItems.item(var_i).ListSubItems(5).Bold = False
             lv_entradas.ListItems.item(var_i).ListSubItems(6).Bold = False
             lv_entradas.ListItems.item(var_i).ListSubItems(7).Bold = False
             lv_entradas.ListItems.item(var_i).ForeColor = &H80000012
             lv_entradas.ListItems.item(var_i).ListSubItems(1).ForeColor = &H80000012
             lv_entradas.ListItems.item(var_i).ListSubItems(2).ForeColor = &H80000012
             lv_entradas.ListItems.item(var_i).ListSubItems(3).ForeColor = &H80000012
             lv_entradas.ListItems.item(var_i).ListSubItems(4).ForeColor = &H80000012
             lv_entradas.ListItems.item(var_i).ListSubItems(5).ForeColor = &H80000012
             lv_entradas.ListItems.item(var_i).ListSubItems(6).ForeColor = &H80000012
             lv_entradas.ListItems.item(var_i).ListSubItems(7).ForeColor = &H80000012
          End If
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_entradas.ListItems.item(var_renglon).Selected = True
      lv_entradas.selectedItem.EnsureVisible
   End If
   lv_entradas.Refresh
End Sub

Sub ejecuta()
   Dim var_nombre_proveedor As String
   Dim var_caja As String
   Dim var_peso As Double
   Dim var_fecha_envio As Date
   Dim var_posible_existen As Boolean
   Dim var_oc_oracle As String
   Dim var_requisicion As Double
   Dim var_cantidad_kilos_metros As Double
   On Error GoTo ersalir:
   var_cajas = False
   var_a?o = 2005
   Dim var_posible_oracle As Boolean
   var_posible_oracle = True
   var_clave_movimiento = txt_clave_movimiento
   var_txt_archivo = txt_archivo
   If var_clave_movimiento = "EC" Then
      var_oc_oracle = Mid(Trim(txt_archivo), 1, 3)
      If var_oc_oracle = "ECP" Then
         var_posible_oracle = False
      Else
         var_posible_oracle = True
      End If
   End If
   If var_clave_movimiento = "EP" Then
   
        rs.Open "Select numb_tra_consecutivo, VCHA_TRA_NOTA_ENVIO , " & _
                        "VCHA_TRA_ALMACEN_ORIGEN, VCHA_ART_ARTICULO_ORIGEN, " & _
                        "NUMB_TRA_CANTIDAD_ENVIADA, VCHA_TRA_REFERENCIA1, vcha_tra_contenedor_id " & _
                    "from tb_transito " & _
                    "where VCHA_TRA_NOTA_ENVIO = '" & txt_archivo.Text & "' ", _
                cnnoracle, _
                adOpenDynamic, _
                adLockOptimistic
        If rs.RecordCount > 0 Then
        For fila = 1 To rs.RecordCount
            rsaux.Open "Update tb_archivo_comparacion " & _
                                "set inte_com_consecutivo =" & rs("numb_tra_consecutivo").Value & ", " & _
                                    "vcha_com_referencia_almacen = '" & rs("VCHA_TRA_ALMACEN_ORIGEN").Value & "'" & _
                                "where vcha_art_articulo_id ='" & rs("VCHA_ART_ARTICULO_ORIGEN").Value & "' " & _
                                "and VCHA_COM_REFERENCIA ='" & UCase(var_nota_traspasos_transito) & "' " & _
                                "and inte_com_lote ='" & rs("VCHA_TRA_REFERENCIA1").Value & "' and (vcha_com_caja = '" & rs("VCHA_TRA_CONTENEDOR_ID").Value & "' or vcha_com_caja = '' )", _
                            cnn, _
                            adOpenDynamic, _
                            adLockOptimistic
            rs.MoveNext
        Next
        Else
            MsgBox "El movimiento no se encontr? en trasito favor de marcar a sistemas", vbCritical, "SID"
        End If
        rs.Close

   End If
   
   If var_posible_oracle = True Then
   rs.Open "select * from vw_bloqueos where vcha_blo_bloqueado_por = '" + txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_solo_lectura = True
      MsgBox "No puede modificar este movimiento ya que la relaci?n esta siendo utilizada por el usuario: '" + Trim(rs!VCHA_USU_NOMBRE) + " " + Trim(rs!vcha_usu_apellidos) + "' en la m?quina: '" + Trim(rs!vcha_blo_maquina) + "'", vbOKOnly, "ATENCION"
   Else
      var_solo_lectura = False
   End If
   rs.Close
   If var_clave_movimiento = "EC" Then
      frmunidad_orden_compra.Show 1
      txt_archivo = Trim(CStr(var_unidad_OC)) + Trim(txt_archivo)
   End If
   Set TB_BLOQUEOS = New TB_BLOQUEOS
   var_clave_moneda = ""
   rs.Open "select * from tb_monedas where inte_mon_moneda_local = 1", cnn, adOpenDynamic, adLockOptimistic
   var_clave_moneda = rs!vcha_mon_moneda_id
   rs.Close
   var_tipo_Cambio = 1
   rs.Open "select * from tb_monedas where inte_mon_moneda_local = 1", cnn, adOpenDynamic, adLockOptimistic
   var_tipo_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
   rs.Close
   rs.Open "select * from tb_movimientos where vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
   var_global_relectura = IIf(IsNull(rs!INTE_MOV_RELECTURA), 0, rs!INTE_MOV_RELECTURA)
   var_global_aceptar_demas = 0
   var_global_aceptar_demas = IIf(IsNull(rs!INTE_MOV_ACEPTAR_MAS), 0, rs!INTE_MOV_ACEPTAR_MAS)
   rs.Close
   var_posible_relectura = False
   If var_global_relectura = 1 Then
      var_posible_relectura = True
   Else
     Cadena = "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_Emo_referencia = '" + txt_archivo + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'"
      Text1 = Cadena
      rs.Open "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_Emo_referencia = '" + txt_archivo + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
   
      If Not rs.EOF Then
         If rs!char_Emo_estatus = "C" Then
            var_posible_relectura = True
         Else
            var_posible_relectura = False
         End If
      Else
         var_posible_relectura = True
      End If
      rs.Close
   End If
   If var_posible_relectura = False Then
      MsgBox "El archivo ya fue leido en otro movimiento", vbOKOnly, "ATENCION"
   Else
      var_clave_movimiento = txt_clave_movimiento
      var_tipo_documento = txt_tipo_documento
      txt_archivo = UCase(txt_archivo)
      
      If var_clave_movimiento = "EI" Then
         rsaux.Open "SELECT VCHA_ALM_ALMACEN_ID FROM TB_ALMACENES WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND INTE_INT_ENTRADA_INTERCOMPA?IA = 1", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_almacen_INT = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
         Else
            If var_empresa = "03" Then
               var_almacen_INT = "8"
            Else
               GoTo ernoalmacen_int:
            End If
         End If
         rsaux.Close
         var_serie = ""
         var_numero_factura = ""
         For var_j = 1 To Len(Me.txt_archivo)
             If Not IsNumeric(Mid(Me.txt_archivo, var_j, 1)) Then
                var_serie = var_serie + Mid(Me.txt_archivo, var_j, 1)
             Else
                var_numero_factura = var_numero_factura + Mid(Me.txt_archivo, var_j, 20)
                Exit For
             End If
         Next var_j
         
         rsaux.Open "select * from tb_archivo_comparacion where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'EI' and vcha_Com_referencia = '" + Me.txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
         If rsaux.EOF Then
            rsaux.Close
            If parametros(1) <> "SIDALMACENBKP" Then
               rsaux.Open "select * from tb_encabezado_cartera where vcha_ser_Serie_id = '" + var_serie + "' and vcha_car_documento = 'FA' and inte_Car_numero = " + var_numero_factura, cnn_distribucion, adOpenDynamic, adLockOptimistic
               'rsaux.Open "select * from tb_encabezado_cartera where vcha_ser_Serie_id = '" + var_serie + "' and vcha_car_documento = 'FA' and inte_Car_numero = " + var_numero_factura, cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
            Else
               rsaux.Open "select * from tb_encabezado_cartera where vcha_ser_Serie_id = '" + var_serie + "' and vcha_car_documento = 'FA' and inte_Car_numero = " + var_numero_factura, cnn, adOpenDynamic, adLockOptimistic
            End If
            
            If Not rsaux.EOF Then
               var_clave_almacen_cantia_autoservicio = ""
               var_clave_almacen_cantia_autoservicio = IIf(IsNull(rsaux!vcha_ESB_ESTABLECIMIENTO_id), "", rsaux!vcha_ESB_ESTABLECIMIENTO_id)
               var_unidad_factura = IIf(IsNull(rsaux!VCHA_UOR_UNIDAD_ID), "", rsaux!VCHA_UOR_UNIDAD_ID)
               If rsaux2.State = 1 Then
                  rsaux2.Close
               End If
               If rsaux2.State = 1 Then
                  rsaux2.Close
               End If
               rsaux2.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_factura + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  If var_unidad_factura = "17" Then
                     var_conexion_facturas_ei = "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=sidtextilera;Data Source=sqlquezada2"
                  Else
                     'MsgBox var_unidad_factura
                     'MsgBox IIf(IsNull(rsaux2!vcha_uor_conexion), "", rsaux2!vcha_uor_conexion)
                     var_conexion_facturas_ei = IIf(IsNull(rsaux2!vcha_uor_conexion), "", rsaux2!vcha_uor_conexion)
                  End If
                  'MsgBox var_conexion_facturas_ei
                  If var_conexion_facturas_ei <> "" Then
                     'MsgBox cnn_facturas_ei
                     If cnn_facturas_ei.State = 1 Then
                        cnn_facturas_ei.Close
                     End If
                     'MsgBox var_conexion_facturas_ei
                     cnn_facturas_ei.Open var_conexion_facturas_ei
                     
                     
                     var_cadena = "SELECT TB_ENCABEZADO_CARTERA.vcha_esb_establecimiento_id, TB_ENCABEZADO_CARTERA.vcha_cli_clave_id, ISNULL(dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_CLASIFICACION, 'PRIMERA') AS VCHA_EMO_CLASIFICACION, dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID , "
                     var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO FROM dbo.TB_ENCABEZADO_CARTERA INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO WHERE (dbo.TB_ENCABEZADO_CARTERA.vcha_Car_documento = 'FA') AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = '" + var_serie + "') AND (dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO = " + var_numero_factura + ")"
                     'MsgBox cnn_facturas_ei.ConnectionString
                     'cnn_facturas_ei.Close
                     'cnn_facturas_ei.Open "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=sidtextilera;Data Source=sqlquezada2"
                     rsaux3.Open var_cadena, cnn_facturas_ei, adOpenDynamic, adLockOptimistic
                     If Not rsaux3.EOF Then
                        var_cliente_factura = rsaux3!vcha_cli_clave_id
                        var_establecimiento_Factura_ei = IIf(IsNull(rsaux3!vcha_ESB_ESTABLECIMIENTO_id), "", rsaux3!vcha_ESB_ESTABLECIMIENTO_id)
                        var_clasificacion = IIf(IsNull(rsaux3!vcha_Emo_clasificacion), "PRIMERA", rsaux3!vcha_Emo_clasificacion)
                     End If
                     rsaux3.Close
                     If var_empresa = "02" And var_cliente_factura = "C000007597" Then
                        GoTo ernoempresa_ei:
                     Else
                        var_costo_promedio = 0
                        var_precio_promedio = 0
                        If var_clasificacion = "" Then
                           var_clasificacion = "PRIMERA"
                        End If
                     
                        If var_clasificacion <> "PRIMERA" Then
                           rsaux3.Open "select sum((floa_sal_Cantidad* (floa_sal_precio * (1 - (FLOA_SAL_DESCUENTO_1 / 100))) * (1 - (FLOA_SAL_DESCUENTO_2 / 100)))) as precio, sum((floa_sal_Cantidad* floa_sal_costo)) as costo, sum(floa_sal_Cantidad) as cantidad from tb_Salidas where vcha_uor_unidad_id = '" + var_unidad_factura + "' and vcha_Ser_serie_id = '" + var_serie + "' and inte_Car_numero = " + var_numero_factura, cnn_facturas_ei, adOpenDynamic, adLockOptimistic
                           var_costo_promedio = IIf(IsNull(rsaux3!Costo), 0, rsaux3!Costo) / IIf(IsNull(rsaux3!Cantidad), 1, rsaux3!Cantidad)
                           var_precio_promedio = IIf(IsNull(rsaux3!Precio), 0, rsaux3!Precio) / IIf(IsNull(rsaux3!Cantidad), 1, rsaux3!Cantidad)
                           rsaux3.Close
                           
                           rsaux5.Open "SELECT * FROM TB_ENTRADAS_INTERCOMPA?IA_CODIGOS_SEGUNDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                           VAR_CODIGO_SEGUNDA = ""
                           If Not rsaux5.EOF Then
                              'MsgBox var_clasificacion
                              VAR_CODIGO_SEGUNDA = IIf(IsNull(rsaux5!VCHA_ART_ARTICULO_ID), "", rsaux5!VCHA_ART_ARTICULO_ID)
                              VAR_KILO_POR_METRO = IIf(IsNull(rsaux5!FLOA_TEM_CANTIDAD_KILO_METRO), 0, rsaux5!FLOA_TEM_CANTIDAD_KILO_METRO)
                           End If
                           rsaux5.Close
                        End If
                        var_cantidad_kilos_metros = 0
                        rsaux3.Open "select a.*, b.vcha_Art_nombre_Espa?ol from tb_Salidas a, tb_Articulos b where a.vcha_uor_unidad_id = '" + var_unidad_factura + "' and a.vcha_Ser_serie_id = '" + var_serie + "' and a.inte_Car_numero = " + var_numero_factura + " and a.vcha_Art_articulo_id = b.vcha_Art_articulo_id", cnn_facturas_ei, adOpenDynamic, adLockOptimistic
                        var_codigos_faltantes = ""
                        If var_empresa = "18" Then
                           If VAR_CODIGO_SEGUNDA = "" Then
                              While Not rsaux3.EOF
                                    rsaux7.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + rsaux3!VCHA_ART_ARTICULO_ID + "' and len(vcha_Art_articulo_id) = 12", cnn, adOpenDynamic, adLockOptimistic
                                    If rsaux7.EOF Then
                                       If var_codigos_faltantes = "" Then
                                          var_codigos_faltantes = rsaux3!VCHA_ART_ARTICULO_ID
                                       Else
                                          var_codigo_faltantes = var_codigo_faltantes + ", " + rsaux3!VCHA_ART_ARTICULO_ID
                                       End If
                                    End If
                                    'MsgBox rsaux3!VCHA_ART_ARTICULO_ID
                                    rsaux7.Close
                                    rsaux3.MoveNext
                              Wend
                           End If
                        End If
                        'MsgBox cnn_facturas_ei
                        cnn.CommandTimeout = 360
                        rsaux3.MoveFirst
                        If var_codigos_faltantes = "" Then
                           If Not rsaux3.EOF Then
                              While Not rsaux3.EOF
                                    If var_clasificacion = "PRIMERA" Then
                                       var_codigo = IIf(IsNull(rsaux3!VCHA_ART_ARTICULO_ID), "", rsaux3!VCHA_ART_ARTICULO_ID)
                                    Else
                                       rsaux5.Open "SELECT * FROM TB_ENTRADAS_INTERCOMPA?IA_CODIGOS_SEGUNDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux5.EOF Then
                                          var_codigo = IIf(IsNull(rsaux5!VCHA_ART_ARTICULO_ID), "", rsaux5!VCHA_ART_ARTICULO_ID)
                                       Else
                                          var_codigo = IIf(IsNull(rsaux3!VCHA_ART_ARTICULO_ID), "", rsaux3!VCHA_ART_ARTICULO_ID)
                                          var_costo_promedio = IIf(IsNull(rsaux3!floa_Sal_costo), 0, rsaux3!floa_Sal_costo)
                                          var_precio_promedio = (rsaux3!floa_Sal_precio * (1 - (rsaux3!FLOA_SAL_DESCUENTO_1 / 100))) * (1 - (rsaux3!FLOA_SAL_DESCUENTO_2 / 100))
                                       End If
                                       rsaux5.Close
                                    End If
                                    If rsaux4.State = 1 Then
                                       rsaux4.Close
                                    End If
                                    If var_empresa = "18" Then
                                       If Trim(var_codigo) = "" Then
                                          rsaux7.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + var_codigo + "' and len(vcha_Art_articulo_id) = 12", cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux7.EOF Then
                                             var_codigo = IIf(IsNull(rsaux7!VCHA_ART_ARTICULO_ID), "", rsaux7!VCHA_ART_ARTICULO_ID)
                                          Else
                                             var_codigo = ""
                                          End If
                                          rsaux7.Close
                                       End If
                                    End If
                                    If rsaux4.State = 1 Then
                                       rsaux4.Close
                                    End If
                                    rsaux4.Open "select * from tb_articulos where vcha_Art_Articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                    'MsgBox var_codigo
                                    If rsaux4.EOF Then
                                       'aqui se dan de alta los codigos que no existen
                                       If rsaux5.State = 1 Then
                                          rsaux5.Close
                                       End If
                                       ' fin de alta de codigos que no existe
                                       'MsgBox cnn.ConnectionString
                                       rsaux5.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux5.EOF Then
                                          var_codigo = IIf(IsNull(rsaux5!VCHA_ART_ARTICULO_ID), "", rsaux5!VCHA_ART_ARTICULO_ID)
                                          If rsaux6.State = 1 Then
                                             rsaux6.Close
                                          End If
                                          rsaux6.Open "select * from tb_articulos where vcha_Art_Articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                          If rsaux6.EOF Then
                                             If var_empresa = "31" And var_unidad_factura = "16" And var_clave_movimiento = "EI" Then
                                                If Mid(var_codigo, 1, 1) = "9" Then
                                                   var_cadena = "insert into tb_articulos (vcha_Art_articulo_id, vcha_Art_nombre_espa?ol, mone_Art_costo_Estandar, mone_Art_precio_base, dtim_Art_fecha_alta, vcha_lic_licencia_id, vcha_art_numero_lic, inte_art_detenido, vcha_equ_equivalencia_id, vcha_art_codigo_externo)"
                                                   var_cadena = var_cadena + "        values ('" + rsaux3!VCHA_ART_ARTICULO_ID + "','" + rsaux3!vcha_Art_nombre_espa?ol + "',0, " + CStr((rsaux3!floa_Sal_precio * (1 - (rsaux3!FLOA_SAL_DESCUENTO_1 / 100))) * (1 - (rsaux3!FLOA_SAL_DESCUENTO_2 / 100))) + ",getdate(),'SIN LICENCIA','SIN LICENCIA',0,'" + rsaux3!VCHA_ART_ARTICULO_ID + "','" + rsaux3!VCHA_ART_ARTICULO_ID + "')"
                                                   rsaux9.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                                   rsaux9.Open "insert into tb_equivalencias (vcha_Art_articulo_id, vcha_equ_Codigo_equivalente) values ('" + rsaux3!VCHA_ART_ARTICULO_ID + "','" + rsaux3!VCHA_ART_ARTICULO_ID + "')", cnn, adOpenDynamic, adLockOptimistic
                                                Else
                                                   GoTo ernocodigos:
                                                End If
                                             Else
                                                GoTo ernocodigos:
                                             End If
                                          End If
                                          rsaux6.Close
                                       Else
                                          If var_empresa = "31" And var_unidad_factura = "16" And var_clave_movimiento = "EI" Then
                                             If Mid(rsaux3!VCHA_ART_ARTICULO_ID, 1, 1) = "9" Or Mid(rsaux3!VCHA_ART_ARTICULO_ID, 1, 1) = "6" Then
                                                var_cadena = "insert into tb_articulos (vcha_Art_articulo_id, vcha_Art_nombre_espa?ol, mone_Art_costo_Estandar, mone_Art_precio_base, dtim_Art_fecha_alta, vcha_lic_licencia_id, vcha_art_numero_lic, inte_art_detenido, vcha_equ_equivalencia_id, vcha_art_codigo_externo)"
                                                var_cadena = var_cadena + "        values ('" + rsaux3!VCHA_ART_ARTICULO_ID + "','" + rsaux3!vcha_Art_nombre_espa?ol + "',0, " + CStr((rsaux3!floa_Sal_precio * (1 - (rsaux3!FLOA_SAL_DESCUENTO_1 / 100))) * (1 - (rsaux3!FLOA_SAL_DESCUENTO_2 / 100))) + ",getdate(),'SIN LICENCIA','SIN LICENCIA',0,'" + rsaux3!VCHA_ART_ARTICULO_ID + "','" + rsaux3!VCHA_ART_ARTICULO_ID + "')"
                                                rsaux6.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                                rsaux6.Open "insert into tb_equivalencias (vcha_Art_articulo_id, vcha_equ_Codigo_equivalente) values ('" + rsaux3!VCHA_ART_ARTICULO_ID + "','" + rsaux3!VCHA_ART_ARTICULO_ID + "')", cnn, adOpenDynamic, adLockOptimistic
                                             Else
                                                GoTo ernocodigos:
                                             End If
                                          Else
                                             GoTo ernocodigos:
                                          End If
                                       End If
                                       rsaux5.Close
                                    End If
                                    rsaux4.Close
                                    'MsgBox var_codigo
                                    rsaux3.MoveNext
                              Wend
                              rsaux3.MoveFirst
                              var_consecutivo = 0
                              While Not rsaux3.EOF
                                    var_consecutivo = var_consecutivo + 1
                                    If var_empresa = "18" Then
                                       'MsgBox var_codigo
                                       rsaux4.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + rsaux3!VCHA_ART_ARTICULO_ID + "' and len(vcha_Art_articulo_id) = 12", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux4.EOF Then
                                          var_codigo = IIf(IsNull(rsaux4!VCHA_ART_ARTICULO_ID), "", rsaux4!VCHA_ART_ARTICULO_ID)
                                       Else
                                          var_codigo = ""
                                       End If
                                       rsaux4.Close
                                    Else
                                       var_codigo = IIf(IsNull(rsaux3!VCHA_ART_ARTICULO_ID), "", rsaux3!VCHA_ART_ARTICULO_ID)
                                    End If
                                    rsaux4.Open "select * from tb_articulos where vcha_Art_Articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If rsaux4.EOF Then
                                       rsaux5.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux5.EOF Then
                                          If VAR_CODIGO_SEGUNDA <> "" Then
                                             var_codigo = VAR_CODIGO_SEGUNDA
                                             var_cantidad = VAR_KILO_POR_METRO * rsaux3!floa_Sal_Cantidad
                                          Else
                                             var_codigo = IIf(IsNull(rsaux5!VCHA_ART_ARTICULO_ID), "", rsaux5!VCHA_ART_ARTICULO_ID)
                                             var_cantidad = rsaux3!floa_Sal_Cantidad
                                             var_costo_promedio = IIf(IsNull(rsaux3!floa_Sal_costo), 0, rsaux3!floa_Sal_costo)
                                             var_precio_promedio = (rsaux3!floa_Sal_precio * (1 - (rsaux3!FLOA_SAL_DESCUENTO_1 / 100))) * (1 - (rsaux3!FLOA_SAL_DESCUENTO_2 / 100))
                                          End If
                                          rsaux6.Open "select * from tb_articulos where vcha_Art_Articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux6.EOF Then
                                             If rsaux8.State = 1 Then
                                                rsaux8.Close
                                             End If
                                             'MsgBox var_establecimiento_Factura_ei
                                             If var_establecimiento_Factura_ei = "E000010695" Then
                                                If var_unidad_factura = "27" Then
                                                   var_almacen_INT = "PTVH"
                                                Else
                                                   var_almacen_INT = "CC_1"
                                                End If
                                             End If
                                             rsaux8.Open "SELECT * FROM TB_ARCHIVO_COMPARACION WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_INT + "' AND VCHA_MOV_MOVIMIENTO_ID = 'EI' AND  vcha_com_referencia = '" + Me.txt_archivo + "' AND VCHA_aRT_ARTICULO_ID = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                             If rsaux8.EOF Then
                                                var_cadena = "insert into tb_Archivo_comparacion (vcha_emp_Empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_com_numero, vcha_Art_articulo_id, dtim_com_fecha, char_com_tipo_proveedor, vcha_com_proveedor, floa_com_costo, floa_Com_Cantidad_enviada, floa_com_cantidad_recibida, vcha_com_referencia, inte_com_consecutivo, inte_com_a?o, floa_com_precio) "
                                                If var_unidad_factura = "17" Then
                                                   var_cadena = var_cadena + "values ('" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_INT + "','EI'," + var_numero_factura + ",'" + var_codigo + "',getdate(),'U','" + var_unidad_factura + "'," + CStr(var_precio_promedio) + "," + CStr(var_cantidad) + ",0,'" + Me.txt_archivo + "', " + CStr(var_consecutivo) + ",2005," + CStr(var_costo_promedio) + ")"
                                                Else
                                                   var_cadena = var_cadena + "values ('" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_INT + "','EI'," + var_numero_factura + ",'" + var_codigo + "',getdate(),'U','" + var_unidad_factura + "'," + CStr(var_precio_promedio) + "," + CStr(var_cantidad) + ",0,'" + Me.txt_archivo + "', " + CStr(var_consecutivo) + ",2005,0)"
                                                End If
                                                rsaux7.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                             Else
                                                rsaux7.Open "UPDATE TB_ARCHIVO_COMPARACION set floa_com_cantidad_enviada = isnull(floa_com_cantidad_enviada,0) + " + CStr(var_cantidad) + " where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_INT + "' AND VCHA_MOV_MOVIMIENTO_ID = 'EI' AND  INTE_COM_NUMERO = " + var_numero_factura + " AND VCHA_aRT_ARTICULO_ID = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                             End If
                                             rsaux8.Close
                                          End If
                                          rsaux6.Close
                                       End If
                                       rsaux5.Close
                                    Else
                                       'MsgBox VAR_CODIGO_SEGUNDA
                                       If VAR_CODIGO_SEGUNDA <> "" Then
                                          var_codigo = VAR_CODIGO_SEGUNDA
                                          var_cantidad = VAR_KILO_POR_METRO * rsaux3!floa_Sal_Cantidad
                                          If var_empresa = "06" Then
                                             var_almacen_INT = "Q0Z"
                                          Else
                                             var_almacen_INT = "RETEX"
                                          End If
                                       Else
                                          If var_empresa = "18" Then
                                             rsaux7.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + IIf(IsNull(rsaux3!VCHA_ART_ARTICULO_ID), "", rsaux3!VCHA_ART_ARTICULO_ID) + "' and len(vcha_Art_articulo_id) = 12", cnn, adOpenDynamic, adLockOptimistic
                                             If Not rsaux7.EOF Then
                                                var_codigo = IIf(IsNull(rsaux7!VCHA_ART_ARTICULO_ID), "", rsaux7!VCHA_ART_ARTICULO_ID)
                                             Else
                                                var_codigo = ""
                                             End If
                                             rsaux7.Close
                                          Else
                                             var_codigo = IIf(IsNull(rsaux3!VCHA_ART_ARTICULO_ID), "", rsaux3!VCHA_ART_ARTICULO_ID)
                                          End If
                                          var_cantidad = rsaux3!floa_Sal_Cantidad
                                          var_costo_promedio = IIf(IsNull(rsaux3!floa_Sal_costo), 0, rsaux3!floa_Sal_costo)
                                          var_precio_promedio = (rsaux3!floa_Sal_precio * (1 - (rsaux3!FLOA_SAL_DESCUENTO_1 / 100))) * (1 - (rsaux3!FLOA_SAL_DESCUENTO_2 / 100))
                                       End If
                                       rsaux8.Open "SELECT * FROM TB_ARCHIVO_COMPARACION WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_INT + "' AND VCHA_MOV_MOVIMIENTO_ID = 'EI' AND  INTE_COM_NUMERO = " + var_numero_factura + " AND VCHA_aRT_ARTICULO_ID = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                       If rsaux8.EOF Then
                                          var_cadena = "insert into tb_Archivo_comparacion (vcha_emp_Empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_com_numero, vcha_Art_articulo_id, dtim_com_fecha, char_com_tipo_proveedor, vcha_com_proveedor, floa_com_costo, floa_Com_Cantidad_enviada, floa_com_cantidad_recibida, vcha_com_referencia, inte_com_consecutivo, inte_com_a?o, floa_com_precio) "
                                          'MsgBox var_almacen_INT
                                          If var_unidad_factura = "17" Then
                                             var_cadena = var_cadena + "values ('" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_INT + "','EI'," + var_numero_factura + ",'" + var_codigo + "',getdate(),'U','" + var_unidad_factura + "'," + CStr(var_precio_promedio) + "," + CStr(var_cantidad) + ",0,'" + Me.txt_archivo + "', " + CStr(var_consecutivo) + ",2005," + CStr(var_costo_promedio) + ")"
                                          Else
                                             var_cadena = var_cadena + "values ('" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_INT + "','EI'," + var_numero_factura + ",'" + var_codigo + "',getdate(),'U','" + var_unidad_factura + "'," + CStr(var_precio_promedio) + "," + CStr(var_cantidad) + ",0,'" + Me.txt_archivo + "', " + CStr(var_consecutivo) + ",2005,0)"
                                          End If
                                          rsaux7.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                       Else
                                          rsaux7.Open "UPDATE TB_ARCHIVO_COMPARACION set floa_com_cantidad_enviada = isnull(floa_com_cantidad_enviada,0) + " + CStr(var_cantidad) + " where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_INT + "' AND VCHA_MOV_MOVIMIENTO_ID = 'EI' AND  INTE_COM_NUMERO = " + var_numero_factura + " AND VCHA_aRT_ARTICULO_ID = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                       End If
                                       rsaux8.Close
                                    End If
                                    rsaux4.Close
                                    rsaux3.MoveNext
                              Wend
                           Else
                              GoTo ernofactura:
                           End If
                        Else
                           GoTo ernocodigos:
                        End If
                        rsaux3.Close
                     End If
                     
                  End If
               Else
                  GoTo ernounidadorganziacional:
               End If
               rsaux2.Close
            Else
               GoTo ersalir:
            End If
            If rsaux.State = 1 Then
               rsaux.Close
            End If
         Else
            rsaux.Close
         End If
      End If
      
      
      
      Set TB_ARCH_COMPARACION_I = New TB_ARCH_COMPARACION_I
      If var_clave_movimiento = "ETA" Then
        rs.Open "select * from tb_archivo_comparacion where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and VCHA_COM_REFERENCIA_TRANSITO = '" + lbl_transito.Caption + "'", cnn, adOpenDynamic, adLockOptimistic
        var_cadena = "select * from tb_archivo_comparacion where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and VCHA_COM_REFERENCIA_TRANSITO = '" + lbl_transito.Caption + "'"
      Else
        rs.Open "select * from tb_archivo_comparacion where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_com_referencia = '" + txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
        var_cadena = "select * from tb_archivo_comparacion where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_com_referencia = '" + txt_archivo + "'"
      End If
      
      'MsgBox var_cadena
      If Not rs.EOF Then
         var_tipo_lecturA_2 = IIf(IsNull(rs!VCHA_cOM_TIPO_LECTURA), "", rs!VCHA_cOM_TIPO_LECTURA)
         If IsNull(rs!VCHA_ALM_ALMACEN_ID) Then
            var_almacen_Destino = ""
         Else
            var_almacen_Destino = Trim(rs!VCHA_ALM_ALMACEN_ID)
         End If
         If IsNull(rs!VCHA_COM_PROVEEDOR) Then
            var_origen = ""
         Else
            'MsgBox var_origen
            var_origen = Trim(rs!VCHA_COM_PROVEEDOR)
         End If
         If IsNull(rs!CHAR_COM_TIPO_PROVEEDOR) Then
            var_tipo_proveedor = ""
         Else
            var_tipo_proveedor = rs!CHAR_COM_TIPO_PROVEEDOR
         End If
         If IsNull(rs!VCHA_COM_TRANSPORTO) Then
            var_transporto = ""
            txt_transporto = Trim(var_transporto)
         Else
            var_transporto = rs!VCHA_COM_TRANSPORTO
            txt_transporto = var_transporto
         End If
         If var_almacen_Destino <> "" Then
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open "select vcha_alm_almacen_id,vcha_alm_nombre from tb_almacenes where vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               'MsgBox var_almacen_Destino
               txt_destino = rsaux(1).Value
               var_almacen_Destino = rsaux(0).Value
               rsaux.Close
            Else
               rsaux.Close
               GoTo noalmacen:
            End If
         End If
         If Trim(var_tipo_proveedor_movimiento) <> Trim(var_tipo_proveedor) Then
            GoTo notipoproveedor:
         End If
         If var_tipo_proveedor <> "" Then
            If var_tipo_proveedor = "U" Then
               'MsgBox "select vcha_uor_unidad_id,vcha_uor_nombre from tb_unidadesorganizacionales where vcha_UOR_unidad_id = '" + Trim(var_origen) + "'"
               rsaux.Open "select vcha_uor_unidad_id,vcha_uor_nombre from tb_unidadesorganizacionales where vcha_UOR_unidad_id = '" + Trim(var_origen) + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  txt_origen = rsaux!VCHA_UOR_NOMBRE
                  var_proveedor = rsaux!VCHA_UOR_UNIDAD_ID
               Else
                  GoTo noproveedor:
               End If
               rsaux.Close
            End If
            If var_tipo_proveedor = "P" Then
               rsaux.Open "select vcha_pro_proveedor_id,vcha_pro_nombre from tb_proveedores where vcha_pro_proveedor_id = '" + var_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  txt_origen = rsaux!VCHA_PRO_NOMBRE
                  var_proveedor = rsaux!VCHA_PRO_PROVEEDOR_ID
               Else
                  GoTo noproveedor:
               End If
               rsaux.Close
            End If
            If var_tipo_proveedor = "T" Then
               rsaux.Open "select vcha_cli_clave_id,vcha_cli_nombre from tb_clientes where vcha_cli_clave_id = '" + var_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  txt_origen = rsaux!VCHA_CLI_NOMBRE
                  var_proveedor = rsaux!vcha_cli_clave_id
               Else
                  GoTo noproveedor:
               End If
               rsaux.Close
            End If
            If var_tipo_proveedor = "A" Then
               rsaux.Open "select vcha_alm_almacen_id,vcha_alm_nombre from tb_almacenes where vcha_alm_almacen_id = '" + var_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  txt_origen = rsaux!VCHA_ALM_NOMBRE
                  var_proveedor = RSUAX!VCHA_ALM_ALMACEN_ID
               Else
                  GoTo noproveedor:
               End If
               rsaux.Close
            End If
            If var_tipo_proveedor = "G" Then
               rsaux.Open "select vcha_age_agente_id,vcha_age_nombre from tb_agentes where vcha_alm_almacen_id = '" + var_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  txt_origen = rsaux!VCHA_AGE_NOMBRE
                  var_proveedor = rsaux!VCHA_AGE_AGENTE_ID
               Else
                  GoTo noproveedor:
               End If
               rsaux.Close
            End If
         Else
            GoTo noproveedor:
         End If
         If IsNull(rs!INTE_COM_NUMERO) Then
            var_folio_enviado = 0
         Else
            var_folio_enviado = rs!INTE_COM_NUMERO
            If var_tipo_documento = "V" Then
               rsaux2.Open "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id ='" + var_unidad_organizacional + "' and inte_emo_numero = " + Str(var_folio_enviado) + " and vcha_mov_movimiento_id = 'SV'", cnn, adOpenDynamic, adLockOptimistic
               var_almacen_origen = rsaux2!VCHA_EMO_ALMACEN_DESTINO
               var_almacen_Destino = rsaux2!vcha_emo_almacen_origen
               rsaux2.Close
            End If
         End If
         rs.Close
         
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If Len(Trim(IIf(IsNull(rs!VCHA_COM_CAJA), "", rs!VCHA_COM_CAJA))) > 0 Then
            lbl_tipo = "C?digo de la caja:"
            var_cajas = True
         Else
            lbl_tipo = "C?digo del art?culo:"
            var_cajas = False
         End If
         var_suma_cantidad_enviada = 0
         var_suma_cantidad_recibida = 0
         
         lbl_enviados.Caption = Format("0", "###,###,##0.00")
         lbl_recibidos.Caption = Format("0", "###,###,##0.00")
         lv_entradas.ListItems.Clear
         While Not rs.EOF
            rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Set list_item = lv_entradas.ListItems.Add(, , Trim(rs!VCHA_ART_ARTICULO_ID))
                   list_item.SubItems(1) = Trim(IIf(IsNull(rsaux(1).Value), "", rsaux(1).Value))
                   list_item.SubItems(2) = Format(IIf(IsNull(rs!FLOA_COM_CANTIDAD_ENVIADA), 0, rs!FLOA_COM_CANTIDAD_ENVIADA), "###,###,##0.00")
                   list_item.SubItems(3) = Format(IIf(IsNull(rs!FLOA_com_cANTIDAD_RECIBIDA), 0, rs!FLOA_com_cANTIDAD_RECIBIDA), "###,###,##0.00")
                   list_item.SubItems(4) = Format(0, "###,###,##0.00")
                   list_item.SubItems(5) = Format(list_item.SubItems(2) - list_item.SubItems(3), "###,###,##0.00")
                   list_item.SubItems(6) = IIf(IsNull(rs!FLOA_COM_COSTO), "", rs!FLOA_COM_COSTO)
                   list_item.SubItems(7) = IIf(IsNull(rsaux(2).Value), "", rsaux(2).Value)
                   list_item.SubItems(8) = IIf(IsNull(rs!INTE_COM_LOTE), "", rs!INTE_COM_LOTE)
                   list_item.SubItems(9) = IIf(IsNull(rs!INTE_COM_CONSECUTIVO), "", rs!INTE_COM_CONSECUTIVO)
                   list_item.SubItems(10) = IIf(IsNull(rs!INTE_COM_A?O), "", rs!INTE_COM_A?O)
                   list_item.SubItems(11) = IIf(IsNull(rs!VCHA_COM_CAJA), "", rs!VCHA_COM_CAJA)
                   list_item.SubItems(12) = IIf(IsNull(rs!FLOA_COM_PESO), 0, rs!FLOA_COM_PESO)
                   list_item.SubItems(13) = IIf(IsNull(rs!P_RC_LINEA_ID), 0, rs!P_RC_LINEA_ID)
                   list_item.SubItems(14) = IIf(IsNull(rs!p_rc_numero_linea), 0, rs!p_rc_numero_linea)
                   list_item.SubItems(15) = IIf(IsNull(rs!VCHA_cOM_TIPO_LECTURA), "", rs!VCHA_cOM_TIPO_LECTURA)
                   list_item.SubItems(16) = IIf(IsNull(rs!FLOA_COM_PRECIO), 0, rs!FLOA_COM_PRECIO)
                   var_suma_cantidad_enviada = Format(var_suma_cantidad_enviada + rs!FLOA_COM_CANTIDAD_ENVIADA, "###,###,##0.00")
                   var_suma_cantidad_recibida = Format(var_suma_cantidad_recibida + rs!FLOA_com_cANTIDAD_RECIBIDA, "###,###,##0.00")
                   var_n = lv_entradas.ListItems.Count
                   If var_n > 10 Then
                      lv_entradas.ColumnHeaders(2).Width = 4700.01
                   Else
                      lv_entradas.ColumnHeaders(2).Width = 4900.01
                   End If
            End If
            rsaux.Close
            rs.MoveNext:
         Wend
         rs.Close
         If var_tipo_documento = "V" Then
            rsaux2.Open "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_ID ='" + var_unidad_organizacional + "' and inte_emo_numero = " + Str(var_folio_enviado) + " and vcha_mov_movimiento_id = 'SV'", cnn, adOpenDynamic, adLockOptimistic
            var_almacen_origen = rsaux2!VCHA_EMO_ALMACEN_DESTINO
            var_almacen_Destino = rsaux2!vcha_emo_almacen_origen
            rsaux2.Close
         End If
         lbl_enviados = Format(Str(var_suma_cantidad_enviada), "###,###,##0.00")
         lbl_recibidos = Format(Str(var_suma_cantidad_recibida), "###,###,##0.00")
         If var_solo_lectura = True Then
            txt_codigo.Enabled = False
         Else
            ok = False
            var_global_bloqueado = 1
            ok = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, txt_archivo, Now, var_clave_usuario_global, fun_NombrePc)
            txt_codigo.Enabled = True
         End If
         txt_archivo.Enabled = False
         If var_tipo_lecturA_2 = "HH" Then
            Me.cmd_tipo_lectura.Visible = True
            Me.txt_codigo.Enabled = False
            Me.cmd_tipo_lectura.Enabled = True
         Else
            If var_tipo_lecturA_2 = "F" Then
               Me.cmd_tipo_lectura.Enabled = False
               Me.cmd_tipo_lectura.Visible = False
               Me.txt_codigo.Enabled = False
               If rsaux11.State = 1 Then
                  rsaux11.Close
               End If
               rsaux11.Open "SELECT * FROM TB_ENCABEZADO_MOVIMIENTOS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND VCHA_EMO_REFERENCIA = '" + Me.txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux11.EOF Then
                  MsgBox "La nota ya fue leida en el movimiento " + CStr(rsaux11!inte_emo_numero), vbOKOnly, "ATENCION"
               Else
                  MsgBox "La nota ya no puede ser modificada", vbOKOnly, "ATENCION"
               End If
               rsaux11.Close
            Else
               If var_tipo_lecturA_2 = "H" Then
                  Me.txt_codigo.Enabled = False
                  Me.cmd_tipo_lectura.Visible = False
                  MsgBox "No se puede abrir la nota ya que se esta leyendo en una maquina portatil", vbOKOnly, "ATENCION"
               Else
                  Me.cmd_tipo_lectura.Visible = False
               End If
            End If
         End If
      Else
         rs.Close
         'On Error GoTo ersalir:
         rs.Open "select * from tb_principal", cnn, adOpenDynamic, adLockOptimistic
         If var_clave_movimiento = "EP" Then
            var_ruta = rs!VCHA_PRI_RUTA_NOTAS_ENVIO
         End If
         If var_clave_movimiento = "DT" Then
            var_ruta = rs!VCHA_PRI_RUTA_DEVOLUCIONES_TIENDA
         End If
         rs.Close
         If var_tabla.State = 1 Then
            var_tabla.Close
         End If
         
         If var_clave_movimiento = "DT" Then
            var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=S?;Collate=Machine;" + """"
            rs.Open "select destino,tipo_prov,proveedor,numnota as folio,codigo,cant1 as cantidad,costo,transporte as transporto,fecha,lote,anocosto, '' as var_caja, 0 as var_peso from " + Trim(var_ruta) + "\" + Right(Trim(txt_archivo), 8), var_tabla, adOpenDynamic, adLockOptimistic
         End If
         If var_clave_movimiento = "EP" Then
            var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=S?;Collate=Machine;" + """"
            'MsgBox var_ruta
            rs.Open "select destino,tipo_prov,proveedor,numnota as folio,codigo,cant1 as cantidad,costo,transporte as transporto,fecha,lote,'2005' as anocosto, var_caja, var_peso from " + Trim(var_ruta) + "\" + Right(Trim(txt_archivo), 8), var_tabla, adOpenDynamic, adLockOptimistic
         End If
         If var_clave_movimiento = "DCOM" Then
            var_cadena = Me.txt_archivo
            var_contado = 0
            tda_codigo = ""
            folest_codigo = ""
            foldoc_codigo = ""
            folconsecutivo = ""
            For var_j = 1 To Len(var_cadena)
                If Mid(var_cadena, var_j, 1) <> "-" Then
                   If var_contador = 0 Then
                      tda_codigo = tda_codigo + Mid(var_cadena, var_j, 1)
                   End If
                   If var_contador = 1 Then
                      folest_codigo = folest_codigo + Mid(var_cadena, var_j, 1)
                   End If
                   If var_contador = 2 Then
                      foldoc_codigo = foldoc_codigo + Mid(var_cadena, var_j, 1)
                   End If
                   If var_contador = 3 Then
                      folconsecutivo = folconsecutivo + Mid(var_cadena, var_j, 1)
                   End If
                Else
                   var_contador = var_contador + 1
                End If
            Next var_j
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_almacen_Destino = IIf(IsNull(rs!vcha_per_almacen_1), "", rs!vcha_per_almacen_1)
               'MsgBox var_almacen_Destino
               If var_almacen_Destino = "" Then
                  var_almacen_Destino = "PTVH"
               End If
            Else
               var_almacen_Destino = "PTVH"
            End If
            rs.Close
            'MsgBox var_almacen_Destino
            rs.Open "select '" + var_almacen_Destino + "' as destino, 'U' as tipo_prov,'" + var_unidad_organizacional + "' as proveedor, a.folconsecutivo as folio, art_codigo as codigo, dns_cantidad as cantidad, dns_costo as costo,'' as transporto, getdate() as fecha, 0 as lote, '2005' as anocosto, '' as var_caja, 0 as var_peso, A.alm_codigo from detallenotassalida a, Notassalida n Where TMA_Codigo = 105 And n.foltda_codigo = " + CStr(tda_codigo) + " And n.folest_codigo = " + CStr(folest_codigo) + " And n.foldoc_codigo = " + CStr(foldoc_codigo) + " And n.folconsecutivo = " + CStr(folconsecutivo) + " and n.foltda_codigo = a.foltda_codigo and n.folest_codigo = a.folest_codigo and n.foldoc_codigo = a.foldoc_codigo and n.folconsecutivo = a.folconsecutivo AND A.alm_codigo <> '2'", cnn_compucaja, adOpenDynamic, adLockOptimistic
         End If
         x = 0
         var_requisicion = 0
         If x = 0 Then
            If var_clave_movimiento = "EC" Then
               txt_archivo = var_txt_archivo
               'Cadena = "SELECT sysdate as fecha, '' AS VAR_CAJA, 0 AS VAR_PESO, TO_NUMBER(PH.SEGMENT1) AS FOLIO, '' AS LOTE, '' as transporto, 'P' AS TIPO_PROV, '8' AS DESTINO, PO.LINE_NUM, IT.SEGMENT3 as codigo, PO.ITEM_DESCRIPTION, PO.UNIT_MEAS_LOOKUP_CODE, (PO.QUANTITY - NVL (REC.NUMB_RC_CANTIDAD, 0)) AS CANTIDAD, PO.LIST_PRICE_PER_UNIT, PO.UNIT_PRICE as costo , PO.PO_HEADER_ID, PV.VENDOR_NAME, PO.PO_LINE_ID, PH.CURRENCY_CODE, PH.SEGMENT1, PV.VENDOR_ID as proveedor, d.destination_organization_id DEST"
               'Cadena = Cadena + "         FROM PO_LINES_ALL@PERPVIA.VIANNEY.COM.MX PO, PO_VENDORS@PERPVIA.VIANNEY.COM.MX PV,"
               'Cadena = Cadena + "         PO_HEADERS_ALL@PERPVIA.VIANNEY.COM.MX PH,"
               'Cadena = Cadena + "         MTL_SYSTEM_ITEMS_B@PERPVIA.VIANNEY.COM.MX IT, po_distributions_all@perpvia.vianney.com.mx d,"
               'Cadena = Cadena + "                        (SELECT   NUMB_RC_ORDEN_COMPRA_ID, NUMB_RC_NUMERO_LINEA,"
               'Cadena = Cadena + "                        NUMB_RC_LINEA_ID, SUM (NUMB_RC_CANTIDAD) NUMB_RC_CANTIDAD"
               'Cadena = Cadena + "                        From RC_TB_RECEPCIONES"
               'Cadena = Cadena + "                        Where NUMB_RC_ORDEN_COMPRA_ID = " + txt_archivo + " and numb_rc_org_id = " + var_unidad_OC
               'Cadena = Cadena + "                        GROUP BY NUMB_RC_ORDEN_COMPRA_ID, NUMB_RC_NUMERO_LINEA, NUMB_RC_LINEA_ID) REC"
               'Cadena = Cadena + "                        Where po.po_line_id = d.po_line_id AND PH.VENDOR_ID = PV.VENDOR_ID"
               'Cadena = Cadena + "                        AND PH.PO_HEADER_ID = PO.PO_HEADER_ID"
               'Cadena = Cadena + "                        AND IT.ORGANIZATION_ID = 83"
               'Cadena = Cadena + "                        AND IT.INVENTORY_ITEM_ID = PO.ITEM_ID"
               'Cadena = Cadena + "                        AND PH.APPROVED_FLAG = 'Y'"
               'Cadena = Cadena + "                        AND PO.PO_LINE_ID = REC.NUMB_RC_LINEA_ID(+)"
               'Cadena = Cadena + "                        AND (PO.QUANTITY - NVL (REC.NUMB_RC_CANTIDAD, 0)) > 0"
               'Cadena = Cadena + "                        AND (PO.CANCEL_FLAG = 'N' OR PO.CANCEL_FLAG IS NULL) "
               'Cadena = Cadena + "                        AND PO.PO_HEADER_ID = (SELECT PO_HEADER_ID"
               'Cadena = Cadena + "                        FROM PO_HEADERS_ALL@PERPVIA.VIANNEY.COM.MX"
               'Cadena = Cadena + "                        Where SEGMENT1 = " + txt_archivo
               'Cadena = Cadena + "                        AND ORG_ID = " + var_unidad_OC + ") ORDER BY PO.LINE_NUM"
               
               
               If var_clave_usuario_global = "U0000000174" Then
                  Cadena = "SELECT SYSDATE AS fecha, '' AS var_caja, 0 AS var_peso, TO_NUMBER (h.segment1) AS folio, '' AS lote, '' AS transporto, 'P' AS tipo_prov, '8' AS destino, l.line_num, i.segment2||i.segment3 AS codigo,"
                  Cadena = Cadena + " l.item_description, l.unit_meas_lookup_code, (l.quantity - NVL (rec.numb_rc_cantidad, 0)) AS cantidad, l.list_price_per_unit, l.unit_price  AS costo, l.po_header_id, v.vendor_name, l.po_line_id, h.currency_code, h.segment1, v.vendor_id AS proveedor, ll.ship_to_organization_id dest, nvl(rate,1) as tipo_Cambio"
                  Cadena = Cadena + " FROM mtl_system_items_b@perpvia.vianney.com.mx i, po_line_locations_all@perpvia.vianney.com.mx ll, po_lines_all@perpvia.vianney.com.mx l, po_headers_all@perpvia.vianney.com.mx h, po_vendors@perpvia.vianney.com.mx v, (SELECT   numb_rc_orden_compra_id, numb_rc_numero_linea, numb_rc_linea_id, SUM (numb_rc_cantidad) numb_rc_cantidad From rc_tb_recepciones "
                  Cadena = Cadena + " Where numb_rc_orden_compra_id = " + txt_archivo + " And numb_rc_org_id = " + var_unidad_OC + " GROUP BY numb_rc_orden_compra_id, numb_rc_numero_linea, numb_rc_linea_id) rec Where ll.po_line_id = l.po_line_id AND ll.po_header_id = l.po_header_id AND ll.po_header_id = h.po_header_id AND l.po_header_id = h.po_header_id AND ll.org_id = l.org_id AND ll.org_id = h.org_id AND l.org_id = h.org_id "
                  Cadena = Cadena + " AND i.inventory_item_id = l.item_id AND v.vendor_id = h.vendor_id AND l.po_line_id = rec.numb_rc_linea_id(+) AND (l.quantity - NVL (rec.numb_rc_cantidad, 0)) > 0 AND i.organization_id = 83 aND h.approved_flag = 'Y' AND (h.cancel_flag = 'N' OR h.cancel_flag IS NULL) AND h.po_header_id IN (SELECT po_header_id FROM po_headers_all@perpvia.vianney.com.mx "
                  Cadena = Cadena + " WHERE segment1 = " + txt_archivo + " AND org_id = " + var_unidad_OC + ")"
               Else
                  'MsgBox var_unidad_OC
                  Cadena = "SELECT SYSDATE AS fecha, '' AS var_caja, 0 AS var_peso, TO_NUMBER (h.segment1) AS folio, '' AS lote, '' AS transporto, 'P' AS tipo_prov, '8' AS destino, l.line_num, i.segment3 AS codigo,"
                  Cadena = Cadena + " l.item_description, l.unit_meas_lookup_code, (l.quantity - NVL (rec.numb_rc_cantidad, 0)) AS cantidad, l.list_price_per_unit, l.unit_price AS costo, l.po_header_id, v.vendor_name, l.po_line_id, h.currency_code, h.segment1, v.vendor_id AS proveedor, ll.ship_to_organization_id dest, nvl(rate,1) as tipo_Cambio, i.attribute7, i.attribute9 "
                  Cadena = Cadena + " FROM mtl_system_items_b@perpvia.vianney.com.mx i, po_line_locations_all@perpvia.vianney.com.mx ll, po_lines_all@perpvia.vianney.com.mx l, po_headers_all@perpvia.vianney.com.mx h, po_vendors@perpvia.vianney.com.mx v, (SELECT   numb_rc_orden_compra_id, numb_rc_numero_linea, numb_rc_linea_id, SUM (numb_rc_cantidad) numb_rc_cantidad From rc_tb_recepciones "
                  Cadena = Cadena + " Where numb_rc_orden_compra_id = " + txt_archivo + " And numb_rc_org_id = " + var_unidad_OC + " GROUP BY numb_rc_orden_compra_id, numb_rc_numero_linea, numb_rc_linea_id) rec Where ll.po_line_id = l.po_line_id AND ll.po_header_id = l.po_header_id AND ll.po_header_id = h.po_header_id AND l.po_header_id = h.po_header_id AND ll.org_id = l.org_id AND ll.org_id = h.org_id AND l.org_id = h.org_id "
                  Cadena = Cadena + " AND i.inventory_item_id = l.item_id AND v.vendor_id = h.vendor_id AND l.po_line_id = rec.numb_rc_linea_id(+) AND (l.quantity - NVL (rec.numb_rc_cantidad, 0)) > 0 AND i.organization_id = 83 aND h.approved_flag = 'Y' AND (h.cancel_flag = 'N' OR h.cancel_flag IS NULL) AND h.po_header_id IN (SELECT po_header_id FROM po_headers_all@perpvia.vianney.com.mx "
                  Cadena = Cadena + " WHERE segment1 = " + txt_archivo + " AND org_id = " + var_unidad_OC + ")"
               
               
               
               End If
               Text1 = Cadena
               
               rs.Open Cadena, cnnoracle, adOpenDynamic, adLockOptimistic
               '''
               If rsaux5.State = 1 Then
                  rsaux5.Close
               End If
               rsaux5.Open "select get_request(" + Trim(Me.txt_archivo) + ", '" + var_unidad_OC + "') from dual", cnnoracle, adOpenDynamic, adLockOptimistic
               txt_archivo = Trim(CStr(var_unidad_OC)) + Trim(txt_archivo)
               If Not rsaux5.EOF Then
                  var_requisicion = rsaux5(0).Value
               End If
               rsaux5.Close
            End If
         End If
         var_posible_existen = True
        
         If Not rs.EOF Then
            While Not rs.EOF
                  If var_empresa = "06" Then
                  
                  Else
                     If var_empresa = "18" And Trim(rs!proveedor) = "2458" Then
                        If Len(Trim(rs!codigo)) = 10 Then
                           rsaux4.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Trim(rs!codigo) + "0" + "'", cnn, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux4.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Trim(rs!codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
                        End If
                     Else
                        If rsaux4.State = 1 Then
                           rsaux4.Close
                        End If
                        'MsgBox rs!codigo
                        rsaux4.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Trim(rs!codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     If rsaux4.EOF Then
                        
                        If var_empresa = 31 And var_clave_movimiento = "EC" Then
                           'MsgBox rs!item_description
                           If rsaux5.State = 1 Then
                              rsaux5.Close
                           End If
                           rsaux5.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Trim(rs!attribute9) + "'", cnn, adOpenDynamic, adLockOptimistic
                           'rsaux5.Open "select * from tb_Articulos where vcha_Art_articulo_id = 'A3416'", cnn, adOpenDynamic, adLockOptimistic
                        Else
                          rsaux5.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Trim(rs!codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        If rsaux5.EOF Then
                           
                           'Aqui se dan de alta los articulos que vienen en la orden de compra
                           If var_empresa = "31" And var_clave_movimiento = "EC" Then
                              var_cadena = "insert into tb_Articulos (vcha_Art_articulo_id, vcha_art_nombre_espa?ol, mone_art_precio_base, mone_art_costo_estandar, dtim_art_fecha_alta, vcha_art_catalogo_vigente, vcha_lic_licencia_id, vcha_art_numero_lic,  vcha_lin_linea_id, vcha_uni_unidad_id, vcha_art_codigo_Externo, inte_art_detenido)"
                              var_cadena = var_cadena + " values ('" + Trim(rs!attribute9) + "','" + rs!item_description + "',0," + CStr(rs!Costo) + ",getdate(),'S/C','SIN LICENCIA','SIN LICENCIA','SL','01','" + Trim(rs!attribute9) + "',1)"
                              'var_cadena = var_cadena + " values ('A3417','" + rs!item_description + "',0," + CStr(rs!Costo) + ",getdate(),'S/C','SIN LICENCIA','SIN LICENCIA','SL','01','A3417',1)"
                              rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              rsaux11.Open "select * from tb_equivalencias where vcha_art_articulo_id = '" + rs!attribute9 + "' and vcha_equ_codigo_equivalente = '" + CStr(rs!attribute7) + "'", cnn, adOpenDynamic, adLockOptimistic
                              If rsaux11.EOF Then
                                 rsaux10.Open "insert into tb_equivalencias (vcha_art_articulo_id, vcha_equ_codigo_equivalente) values ('" + rs!attribute9 + "','" + CStr(rs!attribute7) + "')", cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux11.Close
                              rsaux11.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + rs!codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                              If rsaux11.EOF Then
                                 rsaux10.Open "insert into tb_equivalencias (vcha_art_articulo_id, vcha_equ_codigo_equivalente) values ('" + rs!attribute9 + "','" + rs!codigo + "')", cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux11.Close
                              'rsaux10.Open "insert into tb_equivalencias (vcha_art_articulo_id, vcha_equ_codigo_equivalente) values ('1053760000109','" + rs!codigo + "')", cnn, adOpenDynamic, adLockOptimistic
                           Else
                              var_posible_existen = False
                           End If
                        End If
                        rsaux5.Close
                     End If
                     rsaux4.Close
                  End If
                  rs.MoveNext
            Wend
         End If
         rs.MoveFirst
         If var_posible_existen = True Then
         var_entrada_calidad = False
         If var_causa_devolucion = True Then
            rsaux2.Open "select * from tb_almacenes where inte_alm_calidad = 1 and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               var_almacen_Destino = rsaux2!VCHA_ALM_ALMACEN_ID
               rsaux3.Open "select * from tb_almacenes where inte_alm_costeo =  1 and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux3.EOF Then
                  var_entrada_calidad = True
                  If IsNull(rsaux3!VCHA_ALM_ALMACEN_ID) Then
                     var_entrada_calidad = False
                  Else
                     var_almacen_costeo = rsaux3!VCHA_ALM_ALMACEN_ID
                  End If
               Else
                  var_entrada_calidad = False
               End If
               rsaux3.Close
            Else
               If IsNull(rs!Destino) Then
                  var_almacen_Destino = ""
               Else
                  If UCase(parametros(0)) = "SQLHOUSTON" Then
                     var_almacen_Destino = "CDH"
                  Else
                     var_almacen_Destino = Trim(rs!Destino)
                  End If
                  If Trim(var_almacen_Destino) = "" Then
                     If var_tipo_permiso = 1 Then
                        rsaux4.Open "select vcha_alm_almacen_id, vcha_alm_nombre from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
                        numero_items = 0
                        While Not rsaux4.EOF
                              Set list_item = frmlista.lv_lista.ListItems.Add(, , rsaux4(0).Value)
                              list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rsaux4(1).Value)
                              rsaux4.MoveNext:
                              numero_items = numero_items + 1
                        Wend
                        rsaux4.Close
                        If numero_items > 8 Then
                           frmlista.lv_lista.ColumnHeaders(2).Width = 6600
                        Else
                           frmlista.lv_lista.ColumnHeaders(2).Width = 6800
                        End If
                        frmlista.Caption = "Lista de Proveedores"
                     Else
                        rsaux4.Open "select  vcha_alm_almacen_id, vcha_alm_nombre from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
                        numero_items = 0
                        While Not rs.EOF
                              Set list_item = frmlista.lv_lista.ListItems.Add(, , rsaux4(0).Value)
                              list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rsaux4(1).Value)
                              rsaux4.MoveNext:
                              numero_items = numero_items + 1
                        Wend
                        rsaux4.Close
                        If numero_items > 8 Then
                           frmlista.lv_lista.ColumnHeaders(1).Width = 6600
                        Else
                           frmlista.lv_lista.ColumnHeaders(1).Width = 6800
                        End If
                        frmlista.Caption = "Lista de Almacenes"
                     End If
                     frmlista.Show 1
                     var_almacen_Destino = var_clave_lista_global
                  End If
               End If
            End If
            rsaux2.Close
         Else
            If IsNull(rs!Destino) Then
               var_almacen_Destino = ""
            Else
               var_almacen_Destino = Trim(rs!Destino)
               If var_clave_movimiento = "EC" Then
               
                  If rs!dest = 114 Then
                     var_almacen_Destino = "AG"
                  End If
                  If var_empresa = "18" Then
                     If rs!dest = 524 Then
                        var_almacen_Destino = "AG"
                     Else
                        If var_empresa = "31" Then
                           var_almacen_Destino = "INVH"
                        Else
                           var_almacen_Destino = "PTTEX"
                        End If
                     End If
                   Else
                     If var_empresa = "31" Then
                        var_almacen_Destino = "INVH"
                     End If
                     
                  End If
                  If var_clave_usuario_global = "U0000000061" Then
                     var_almacen_Destino = "AC"
                  End If
                  If var_clave_usuario_global = "U0000000174" Then
                     var_almacen_Destino = "MUEBLES"
                  End If
                  If var_empresa = "06" Then
                     var_almacen_Destino = "Q0Z"
                  End If
                  If var_clave_movimiento = "DCOM" Then
                     If var_clave_usuario_global = "U0000000145" Or var_clave_usuario_global = "U0000000150" Then
                        var_almacen_Destino = "CANCA"
                     Else
                        var_almacen_Destino = "PTVH"
                     End If
                  End If

               End If
               If var_clave_movimiento = "DCOM" Then
                  If var_clave_usuario_global = "U0000000145" Or var_clave_usuario_global = "U0000000150" Then
                     'var_almacen_Destino = "CANCA"
                  Else
                     'var_almacen_Destino = "PTVH"
                  End If
               End If
               If var_empresa = "02" Then
                  If var_clave_usuario_global = "U0000000058" Then
                     var_almacen_Destino = "AG"
                  End If
               End If
            End If
         End If
         If IsNull(rs!proveedor) Then
            var_origen = ""
         Else
            var_origen = Trim(rs!proveedor)
            If var_clave_movimiento = "EC" Then
               var_nombre_proveedor = IIf(IsNull(rs!vendor_name), "", rs!vendor_name)
            End If
         End If
         If IsNull(rs!tipo_prov) Then
            var_tipo_proveedor = ""
         Else
            var_tipo_proveedor = rs!tipo_prov
         End If
         If IsNull(rs!transporto) Then
            var_transporto = ""
            txt_transporto = Trim(var_transporto)
         Else
            var_transporto = rs!transporto
            txt_transporto = var_transporto
         End If
         If var_almacen_Destino <> "" Then
            rsaux.Open "select vcha_alm_almacen_id,vcha_alm_nombre from tb_almacenes where vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               txt_destino = rsaux(1).Value
               var_almacen_Destino = rsaux(0).Value
               rsaux.Close
            Else
               rsaux.Close
               If var_tipo_permiso = 1 Then
                  rs.Open "select vcha_alm_nombre, vcha_alm_almacen_id from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
                  numero_items = 0
                  While Not rs.EOF
                        Set list_item = frmlista.lv_lista.ListItems.Add(, , rs(0).Value)
                        list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
                        rs.MoveNext:
                        numero_items = numero_items + 1
                  Wend
                  rs.Close
                  If numero_items > 8 Then
                     frmlista.lv_lista.ColumnHeaders(1).Width = 6600
                  Else
                     frmlista.lv_lista.ColumnHeaders(1).Width = 6800
                  End If
                  frmlista.Caption = "Lista de Proveedores"
               Else
                  rs.Open "select  vcha_alm_nombre, vcha_alm_almacen_id from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
                  numero_items = 0
                  While Not rs.EOF
                        Set list_item = frmlista.lv_lista.ListItems.Add(, , rs(0).Value)
                        list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
                        rs.MoveNext:
                        numero_items = numero_items + 1
                  Wend
                  rs.Close
                  If numero_items > 8 Then
                     frmlista.lv_lista.ColumnHeaders(1).Width = 6600
                  Else
                     frmlista.lv_lista.ColumnHeaders(1).Width = 6800
                  End If
                  frmlista.Caption = "Lista de Almacenes"
               End If
               frmlista.Show 1
               txt_destino = var_nombre_lista_global
               var_almacen_Destino = var_clave_lista_global
               If Trim(var_almacen_Destino) = "" Then
                  GoTo noalmacen:
               End If
            End If
         End If
         If Trim(var_tipo_proveedor_movimiento) <> Trim(var_tipo_proveedor) Then
            GoTo notipoproveedor:
         End If
         If var_tipo_proveedor <> "" Then
            If var_tipo_proveedor = "U" Then
               rsaux.Open "select vcha_uor_unidad_id,vcha_uor_nombre from tb_unidadesorganizacionales where vcha_UOR_unidad_id = '" + var_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_proveedor = rsaux!VCHA_UOR_UNIDAD_ID
                  txt_origen = rsaux!VCHA_UOR_NOMBRE
               Else
                  GoTo noproveedor:
               End If
               rsaux.Close
            End If
            If var_tipo_proveedor = "P" Then
               rsaux.Open "select vcha_pro_proveedor_id,vcha_pro_nombre from tb_proveedores where vcha_pro_proveedor_id = '" + var_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_proveedor = rsaux!VCHA_PRO_PROVEEDOR_ID
                  txt_origen = rsaux!VCHA_PRO_NOMBRE
               Else
                  var_nombre_proveedor_2 = ""
                  For var_jj = 1 To Len(var_nombre_proveedor)
                      If Mid(var_nombre_proveedor, var_jj, 1) <> "'" Then
                         var_nombre_proveedor_2 = var_nombre_proveedor_2 + Mid(var_nombre_proveedor, var_jj, 1)
                      End If
                  Next var_jj
                  var_nombre_proveedor = var_nombre_proveedor_2
                  rsaux5.Open "insert into tb_proveedores (vcha_pro_proveedor_id,vcha_pro_nombre) values ('" + var_origen + "','" + var_nombre_proveedor + "')", cnn, adOpenDynamic, adLockOptimistic
                  'GoTo noproveedor:
               End If
               rsaux.Close
            End If
            If var_tipo_proveedor = "T" Then
               rsaux.Open "select vcha_cli_clave_id,vcha_cli_nombre from TB_clientes where vcha_cli_clave_id = '" + var_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_proveedor = rsaux!vcha_cli_clave_id
                  txt_origen = rsaux!VCHA_CLI_NOMBRE
               Else
                  rsaux4.Open "select vcha_cli_clave_id,vcha_cli_nombre from tb_clientes where VCHA_TCL_TIPO_CLIENTE_ID = 'T' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
                  numero_items = 0
                  While Not rsaux4.EOF
                        Set list_item = frmlista.lv_lista.ListItems.Add(, , rsaux4(0).Value)
                        list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rsaux4(1).Value)
                        rsaux4.MoveNext:
                        numero_items = numero_items + 1
                  Wend
                  rsaux4.Close
                  If numero_items > 8 Then
                     frmlista.lv_lista.ColumnHeaders(2).Width = 6600
                  Else
                     frmlista.lv_lista.ColumnHeaders(2).Width = 6800
                  End If
                  frmlista.Caption = "Lista de Proveedores"
                  frmlista.Show 1
                  var_proveedor = var_clave_lista_global
                  var_origen = var_clave_lista_global
                  txt_origen = var_nombre_lista_global
                  If Trim(var_clave_lista_global) = "" Then
                     GoTo noproveedor:
                  End If
               End If
               rsaux.Close
            End If
            If var_tipo_proveedor = "A" Then
               rsaux.Open "select vcha_alm_almacen_id,vcha_alm_nombre from tb_almacenes where vcha_alm_almacen_id = '" + var_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_proveedor = rsaux!VCHA_ALM_ALMACEN_ID
                  txt_origen = rsaux!VCHA_ALM_NOMBRE
               Else
                  GoTo noproveedor:
               End If
               rsaux.Close
            End If
            If var_tipo_proveedor = "G" Then
               rsaux.Open "select vcha_age_agente_id,vcha_age_nombre from tb_agentes where vcha_age_agente_id = '" + var_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_proveedor = rsaux!VCHA_AGE_AGENTE_ID
                  txt_origen = rsaux!VCHA_AGE_NOMBRE
               Else
                  GoTo noproveedor:
               End If
               rsaux.Close
            End If
         Else
            GoTo noproveedor:
         End If
         rs.MoveFirst
         var_consecutivo = 0
         var_fecha_envio = rs!fecha
         If var_clave_movimiento = "EC" Then
               var_clave_moneda_oracle = IIf(IsNull(rs!currency_code), "MXP", rs!currency_code)
               If var_clave_moneda_oracle = "MXP" Then
                  VAR_CLAVE_MONEDA_STR = "1"
               Else
                  If var_clave_moneda_oracle = "USD" Then
                     VAR_CLAVE_MONEDA_STR = "2"
                  Else
                     If var_clave_moneda_oracle = "EUR" Then
                        VAR_CLAVE_MONEDA_STR = "3"
                     Else
                        If var_clave_moneda_oracle = "YEN" Then
                           VAR_CLAVE_MONEDA_STR = "4"
                        Else
                           VAR_CLAVE_MONEDA_STR = "1"
                        End If
                     End If
                  End If
               End If
               If VAR_CLAVE_MONEDA_STR <> "1" Then
                  var_tipo_cambio_global = 1
                  var_cadena = "select from_currency, to_currency, conversion_date, conversion_type, conversion_rate from gl_daily_rates@perpvia.vianney.com.mx where from_currency = 'USD' and to_currency = 'MXP' and conversion_type = 1000 AND TO_CHAR(CONVERSION_DATE, 'DD/MM/YYYY') = TO_CHAR(SYSDATE, 'DD/MM/YYYY')"
                  rsaux10.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                  If Not rsaux10.EOF Then
                     var_tipo_cambio_global = IIf(IsNull(rsaux10!conversion_rate), 1, rsaux10!conversion_rate)
                  Else
                     frmtipo_Cambio.Show 1
                  End If
                  rsaux10.Close
                  
                  
               End If
         End If
         rs.MoveFirst
         While Not rs.EOF
            var_lote = IIf(IsNull(rs!lote), 0, rs!lote)
            var_consecutivo = var_consecutivo + 1
            If IsNull(rs!codigo) Then
               var_articulo_enviado = ""
            Else
               var_articulo_enviado = Trim(rs!codigo)
            End If
            If var_empresa = 31 And var_clave_movimiento = "EC" Then
               If rs!codigo = "0003" Then
                  var_articulo_enviado = "0003"
               Else
                  
                  var_articulo_enviado = Trim(rs!attribute9)
                  'var_articulo_eviado = "A1722"
               End If
               '"
               
            End If
            If IsNull(rs!Costo) Then
               var_costo_enviado = 0
            Else
               var_costo_enviado = rs!Costo
            End If
            'MsgBox var_costo_enviado
            If IsNull(rs!Cantidad) Then
               VAR_CANTIDAD_ENIVADA = 0
            Else
               var_cantidad_enviada = rs!Cantidad
            End If
            If IsNull(rs!FOLIO) Then
               var_folio_enviado = 0
            Else
               var_folio_enviado = rs!FOLIO
            End If
            If var_entrada_calidad = True Then
                rsaux2.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_costeo + "' and vcha_art_articulo_id = '" + Trim(var_articulo_enviado) + "'", cnn, adOpenDynamic, adLockOptimistic
                If Not rsaux2.EOF Then
                   var_costo_enviado = rsaux2!FLOA_eXI_COSTO
                   rsaux2.Close
                Else
                   rsaux2.Close
                   rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + Trim(var_articulo_enviado) + "'", cnn, adOpenDynamic, adLockOptimistic
                   If Not rsaux2.EOF Then
                      var_costo_enviado = rsaux2!mone_Art_costo_estandar
                   End If
                   rsaux2.Close
                End If
            End If
            rsaux.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + Trim(var_articulo_enviado) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               rsaux.Close
            Else
               rsaux.Close
               If var_empresa = "18" And Trim(rs!proveedor) = "2458" Then
                  rsaux.Open "select vcha_art_articulo_id from tb_equivalencias where VCHA_EQU_CODIGO_EQUIVALENTE = '" + Trim(var_articulo_enviado) + "0'", cnn, adOpenDynamic, adLockOptimistic
               Else
                  rsaux.Open "select vcha_art_articulo_id from tb_equivalencias where VCHA_EQU_CODIGO_EQUIVALENTE = '" + Trim(var_articulo_enviado) + "'", cnn, adOpenDynamic, adLockOptimistic
               End If
               If Not rsaux.EOF Then
                  var_articulo_enviado = rsaux!VCHA_ART_ARTICULO_ID
                  rsaux.Close
               Else
                  rsaux.Close
               End If
            End If
            If var_clave_movimiento = "DT" Then
               var_a?o = CInt(rs!ANOCOSTO)
            End If
            var_caja = IIf(IsNull(rs!var_caja), "", rs!var_caja)
            var_peso = IIf(IsNull(rs!var_peso), 0, rs!var_peso)
            If var_clave_movimiento = "EC" And var_empresa = "18" Then
               rsaux5.Open "select * from tb_requisiciones where inte_req_numero = " + CStr(var_requisicion), cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux5.EOF Then
                  var_costo_enviado = var_costo_enviado + IIf(IsNull(rsaux5!floa_req_costo), 0, rsaux5!floa_req_costo)
               End If
               rsaux5.Close
               
            End If
            If var_tipo_cambio_global = 0 Then
               var_tipo_cambio_global = 1
            End If
            If var_clave_movimiento = "EC" Then
               'MsgBox var_costo_enviado
               ok = TB_ARCH_COMPARACION_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), var_fecha_envio, var_tipo_proveedor, var_origen, var_articulo_enviado, var_costo_enviado * var_tipo_cambio_global, var_cantidad_enviada, 0, var_transporto, txt_archivo, var_lote, var_consecutivo, var_a?o, var_caja, var_peso)
            Else
               ok = TB_ARCH_COMPARACION_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, CDbl(var_folio_enviado), var_fecha_envio, var_tipo_proveedor, var_origen, var_articulo_enviado, var_costo_enviado, var_cantidad_enviada, 0, var_transporto, txt_archivo, var_lote, var_consecutivo, var_a?o, var_caja, var_peso)
            End If
            If var_clave_movimiento = "EC" Then
               var_clave_moneda_oracle = IIf(IsNull(rs!currency_code), "MXP", rs!currency_code)
               If var_clave_moneda_oracle = "MXP" Then
                  VAR_CLAVE_MONEDA_STR = "1"
               Else
                  If var_clave_moneda_oracle = "USD" Then
                     VAR_CLAVE_MONEDA_STR = "2"
                  Else
                     If var_clave_moneda_oracle = "EUR" Then
                        VAR_CLAVE_MONEDA_STR = "3"
                     Else
                        If var_clave_moneda_oracle = "YEN" Then
                           VAR_CLAVE_MONEDA_STR = "4"
                        Else
                           VAR_CLAVE_MONEDA_STR = "1"
                        End If
                     End If
                  End If
               End If
               VAR_CLAVE_ALMACEN_COMPUCAJA = ""
               If var_tipo_cambio_global = 0 Then
                  var_tipo_cambio_global = 1
               End If
               If var_tipo_cambio_global = 1 Then
                  rsaux5.Open "update tb_archivo_comparacion set VCHA_COM_REFERENCIA_ALMACEN = '" + VAR_CLAVE_ALMACEN_COMPUCAJA + "', floa_com_tipo_cambio_oracle = " + CStr(var_tipo_cambio_global) + ", vcha_mon_moneda_id = '" + VAR_CLAVE_MONEDA_STR + "',P_RC_ORD_ID = " + var_unidad_OC + ", P_RC_LINEA_ID = " + CStr(rs!po_line_id) + ", P_RC_NUMERO_LINEA = " + CStr(rs!LINE_NUM) + ", floa_com_costo = floa_com_costo * 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_com_numero = " + CStr(CDbl(var_folio_enviado)) + " and vcha_com_referencia = '" + txt_archivo + "' and vcha_art_articulo_id = '" + var_articulo_enviado + "' and inte_com_consecutivo  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               Else
                  rsaux5.Open "update tb_archivo_comparacion set VCHA_COM_REFERENCIA_ALMACEN = '" + VAR_CLAVE_ALMACEN_COMPUCAJA + "', floa_com_tipo_cambio_oracle = " + CStr(var_tipo_cambio_global) + ", vcha_mon_moneda_id = '" + VAR_CLAVE_MONEDA_STR + "',P_RC_ORD_ID = " + var_unidad_OC + ", P_RC_LINEA_ID = " + CStr(rs!po_line_id) + ", P_RC_NUMERO_LINEA = " + CStr(rs!LINE_NUM) + ", floa_com_costo = floa_com_costo where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_com_numero = " + CStr(CDbl(var_folio_enviado)) + " and vcha_com_referencia = '" + txt_archivo + "' and vcha_art_articulo_id = '" + var_articulo_enviado + "' and inte_com_consecutivo  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               End If
            Else
               If var_clave_movimiento = "DCOM" Then
                  'Aqui debe de ir la seleccion del almacen de acuerdo a la que tenga permiso el usuario
                  VAR_CLAVE_ALMACEN_COMPUCAJA = ""
                  If var_clave_movimiento = "DCOM" Then
                     VAR_CLAVE_ALMACEN_COMPUCAJA = IIf(IsNull(rs!ALM_CODIGO), "", rs!ALM_CODIGO)
                  End If
                  rsaux5.Open "update tb_archivo_comparacion set VCHA_COM_REFERENCIA_ALMACEN = '" + CStr(VAR_CLAVE_ALMACEN_COMPUCAJA) + "' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_com_numero = " + CStr(CDbl(var_folio_enviado)) + " and vcha_com_referencia = '" + txt_archivo + "' and vcha_art_articulo_id = '" + var_articulo_enviado + "' and inte_com_consecutivo  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               End If
            End If
            rs.MoveNext
         Wend
         rs.Close
         var_suma_cantidad_enviada = 0
         var_suma_cantidad_recibida = 0
         lbl_enviados.Caption = "0"
         lbl_recibidos.Caption = "0"
         lv_entradas.ListItems.Clear
         rs.Open "select * from tb_archivo_comparacion where vcha_emp_empresa_id ='" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_com_referencia = '" + txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Len(Trim(IIf(IsNull(rs!VCHA_COM_CAJA), "", rs!VCHA_COM_CAJA))) > 0 Then
            var_cajas = True
            lbl_tipo = "C?digo de la caja:"
         Else
            lbl_tipo = "C?digo del art?culo:"
            var_cajas = False
         End If
         While Not rs.EOF
            rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Set list_item = lv_entradas.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
                   list_item.SubItems(1) = IIf(IsNull(rsaux(1).Value), "", rsaux(1).Value)
                   list_item.SubItems(2) = Format(IIf(IsNull(rs!FLOA_COM_CANTIDAD_ENVIADA), 0, rs!FLOA_COM_CANTIDAD_ENVIADA), "###,###,##0.00")
                   list_item.SubItems(3) = Format(IIf(IsNull(rs!FLOA_com_cANTIDAD_RECIBIDA), 0, rs!FLOA_com_cANTIDAD_RECIBIDA), "###,###,##0.00")
                   list_item.SubItems(4) = Format(0, "###,###,##0.00")
                   list_item.SubItems(5) = Format(list_item.SubItems(2) - list_item.SubItems(3), "###,###,##0.00")
                   list_item.SubItems(6) = IIf(IsNull(rs!FLOA_COM_COSTO), "", rs!FLOA_COM_COSTO)
                   list_item.SubItems(7) = IIf(IsNull(rsaux(2).Value), "", rsaux(2).Value)
                   list_item.SubItems(8) = IIf(IsNull(rs!INTE_COM_LOTE), "", rs!INTE_COM_LOTE)
                   list_item.SubItems(9) = IIf(IsNull(rs!INTE_COM_CONSECUTIVO), "", rs!INTE_COM_CONSECUTIVO)
                   list_item.SubItems(10) = IIf(IsNull(rs!INTE_COM_A?O), "", rs!INTE_COM_A?O)
                   list_item.SubItems(11) = IIf(IsNull(rs!VCHA_COM_CAJA), "", rs!VCHA_COM_CAJA)
                   list_item.SubItems(12) = IIf(IsNull(rs!FLOA_COM_PESO), "", rs!FLOA_COM_PESO)
                   If var_clave_movimiento = "EC" Then
                      list_item.SubItems(13) = IIf(IsNull(rs!P_RC_LINEA_ID), 0, rs!P_RC_LINEA_ID)
                      list_item.SubItems(14) = IIf(IsNull(rs!p_rc_numero_linea), 0, rs!p_rc_numero_linea)
                   End If
                   list_item.SubItems(15) = IIf(IsNull(rs!VCHA_cOM_TIPO_LECTURA), "", rs!VCHA_cOM_TIPO_LECTURA)
                   list_item.SubItems(16) = IIf(IsNull(rs!FLOA_COM_PRECIO), 0, rs!FLOA_COM_PRECIO)
                   
                   var_suma_cantidad_enviada = var_suma_cantidad_enviada + rs!FLOA_COM_CANTIDAD_ENVIADA
                   var_suma_cantidad_recibida = var_suma_cantidad_recibida + rs!FLOA_com_cANTIDAD_RECIBIDA
                   var_n = lv_entradas.ListItems.Count
                   If var_n > 10 Then
                      lv_entradas.ColumnHeaders(2).Width = 4700.01
                   Else
                      lv_entradas.ColumnHeaders(2).Width = 4900.01
                   End If
            End If
            rsaux.Close
            rs.MoveNext:
         Wend
         rs.Close
         If var_tipo_documento = "V" Then
            rsaux2.Open "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_ID ='" + var_unidad_organizacional + "' and inte_emo_numero = " + Str(var_folio_enviado) + " and vcha_mov_movimiento_id = 'AV'", cnn, adOpenDynamic, adLockOptimistic
            var_almacen_origen = rsaux2!VCHA_EMO_ALMACEN_DESTINO
            var_almacen_Destino = rsaux2!vcha_emo_almacen_origen
            rsaux2.Close
         End If
         lbl_enviados = Format(Str(var_suma_cantidad_enviada), "###,###,##0.00")
         lbl_recibidos = Format(Str(var_suma_cantidad_recibida), "###,###,##0.00")
         If var_solo_lectura = True Then
            txt_codigo.Enabled = False
         Else
            ok = False
            var_global_bloqueado = 1
            ok = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, txt_archivo, Now, var_clave_usuario_global, fun_NombrePc)
            txt_codigo.Enabled = True
         End If
         txt_archivo.Enabled = False
         Else
            rs.Close
            MsgBox "El archivo contiene c?digos de art?culos que no existen en el almacen", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   var_posible = 1
   'permiso diana
   If var_tipo_permiso = 1 Then
      rsaux.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_per_almacen_1 = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockBatchOptimistic
      If Not rsaux.EOF Then
         var_posible = 0
         'If var_empresa = "16" Then
         '   Me.cmd_pasar_todos.Visible = True
         '   Me.cmd_pasar_todos.Enabled = True
         '   Me.txt_codigo.Enabled = False
         'Else
         '   Me.cmd_pasar_todos.Visible = False
         'End If
      End If
      rsaux.Close
   Else
      'If var_empresa = "16" Then
      '   Me.cmd_pasar_todos.Visible = True
      '   Me.cmd_pasar_todos.Enabled = True
      '   Me.txt_codigo.Enabled = False
      'Else
      '   Me.cmd_pasar_todos.Visible = False
      'End If
      var_posible = 0
   End If
   If var_posible = 1 Then
      Me.cmd_pasar_todos.Enabled = False
      MsgBox "No esta autorizado para modificar el movimiento", vbOKOnly, "ATENCION"
      txt_factura.Enabled = False
      txt_codigo.Enabled = False
   End If
   Else
      MsgBox "No es posible cargar este tipo de ordenes de compra", vbOKOnly, "ATENCION"
   End If
Exit Sub
ernoalmacen_int:
   MsgBox "No existe un almac?n para hacer las entradas por compra intercompa?ia", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
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
   txt_archivo.Enabled = True
   If var_solo_lectura = False Then
      Set TB_BLOQUEOS = New TB_BLOQUEOS
      var_global_bloqueado = 0
      ok = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, txt_archivo, Now, "", "")
   End If
   txt_destino = ""
   txt_origen = ""
   txt_archivo = ""
   txt_factura = ""
   Me.txt_pedimento = ""
   txt_transporto = ""
Exit Sub


ernocodigos:
   If var_codigos_faltantes = "" Then
      MsgBox "La factura contiene codigos que no estan dados de alta en la planta", vbOKOnly, "ATENCION"
   Else
      MsgBox "Los siguientes c?digos no tienen equivalencia en el almacen: " + var_codigos_faltantes, vbOKOnly, "ATENCION"
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
   txt_archivo.Enabled = True
   If var_solo_lectura = False Then
      Set TB_BLOQUEOS = New TB_BLOQUEOS
      var_global_bloqueado = 0
      ok = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, txt_archivo, Now, "", "")
   End If
   txt_destino = ""
   txt_origen = ""
   txt_archivo = ""
   txt_factura = ""
   Me.txt_pedimento = ""
   txt_transporto = ""
Exit Sub

ernofactura:
   MsgBox "No existe conexion hacia la planta de la factura seleccionada", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
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
   txt_archivo.Enabled = True
   If var_solo_lectura = False Then
      Set TB_BLOQUEOS = New TB_BLOQUEOS
      var_global_bloqueado = 0
      ok = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, txt_archivo, Now, "", "")
   End If
   txt_destino = ""
   txt_origen = ""
   txt_archivo = ""
   txt_factura = ""
   Me.txt_pedimento = ""
   txt_transporto = ""
Exit Sub

ernounidadorganziacional:
   MsgBox "No existe conexion hacia la planta de la factura seleccionada", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
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
   txt_archivo.Enabled = True
   If var_solo_lectura = False Then
      Set TB_BLOQUEOS = New TB_BLOQUEOS
      var_global_bloqueado = 0
      ok = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, txt_archivo, Now, "", "")
   End If
   txt_destino = ""
   txt_origen = ""
   txt_archivo = ""
   txt_factura = ""
   Me.txt_pedimento = ""
   txt_transporto = ""
Exit Sub

ersalir:
   MsgBox "A surgido un error al leer el archivo, puede que el archivo no exista o este siendo utilizado por otro usuario", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   txt_archivo.Enabled = True
   If var_solo_lectura = False Then
      Set TB_BLOQUEOS = New TB_BLOQUEOS
      var_global_bloqueado = 0
      ok = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, txt_archivo, Now, "", "")
   End If
   txt_destino = ""
   txt_origen = ""
   txt_archivo = ""
   txt_factura = ""
   Me.txt_pedimento = ""
   txt_transporto = ""
Exit Sub
noalmacen:
   MsgBox "El almac?n destino no existe", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If var_solo_lectura = False Then
      Set TB_BLOQUEOS = New TB_BLOQUEOS
      var_global_bloqueado = 0
      ok = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, txt_archivo, Now, "", "")
   End If
   txt_destino = ""
   txt_origen = ""
   txt_archivo = ""
   txt_factura = ""
   Me.txt_pedimento = ""
   txt_transporto = ""
Exit Sub
noproveedor:
   MsgBox "El proveedor no existe", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If var_solo_lectura = False Then
      Set TB_BLOQUEOS = New TB_BLOQUEOS
      var_global_bloqueado = 0
      ok = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, txt_archivo, Now, "", "")
   End If
   txt_destino = ""
   txt_origen = ""
   txt_archivo = ""
   txt_factura = ""
   Me.txt_pedimento = ""
   txt_transporto = ""
Exit Sub
notipoproveedor:
   MsgBox "El archivo leido no corresponde con el tipo de proveedor del movimiento"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If var_solo_lectura = False Then
      Set TB_BLOQUEOS = New TB_BLOQUEOS
      var_global_bloqueado = 0
      ok = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, txt_archivo, Now, "", "")
   End If
   txt_destino = ""
   txt_origen = ""
   txt_archivo = ""
   txt_factura = ""
   Me.txt_pedimento = ""
   txt_transporto = ""
Exit Sub
ernoempresa_ei:
   MsgBox "El cliente seleccionado pertenece a la empresa de TELAS DEL HOGAR, favor de hacer la entrada en esa empresa", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
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
  
Exit Sub
e:
    a = Err.Description
End Sub

Private Sub txt_numero_serie_GotFocus()
   Me.txt_numero_serie = ""
End Sub

Private Sub txt_numero_serie_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Cadena = "select * from TB_EXISTENCIAS_SERIES WHERE vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_eMO_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "' and vcha_art_numero_serie = '" + Me.txt_numero_serie + "'"
      rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         rs.Close
         MsgBox "Ya existe el n?mero de serie", vbOKOnly, "ATENCION"
         Me.txt_foco.Enabled = False
      Else
         rs.Close
         Me.txt_foco.Enabled = True
         If Trim(Me.txt_numero_serie) <> "" Then
            If Me.txt_Cantidad.Visible = True Then
               Me.txt_Cantidad.SetFocus
            Else
               If Me.txt_foco.Enabled = True Then
                  Me.txt_foco.SetFocus
               End If
            End If
         End If
      End If
   End If
   If KeyAscii = 27 Then
      Me.txt_foco.Enabled = False
      Me.txt_codigo.Enabled = True
      Me.txt_Cantidad.Visible = False
      Me.lbl_Cantidad.Visible = False
      var_costo_tela = 0
      Me.txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_numero_serie_LostFocus()
   frmnumero_serie.Visible = False
End Sub

Private Sub txt_pedimento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_codigo.Enabled = True
      Me.txt_codigo.SetFocus
      
   End If
End Sub

Private Sub txt_pedimento_LostFocus()
   Me.txt_pedimento.Enabled = False
End Sub

Private Sub txt_peso_caja_KeyPress(KeyAscii As Integer)
   Dim var_porcentaje_peso As Double
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If IsNumeric(txt_peso_caja) Then
         Dim var_peso_caja_leido As Double
         Dim var_posible_peso_caja As Boolean
         var_peso_caja_leido = CDbl(txt_peso_caja)
         var_posible_peso_caja = False
         If var_peso_caja_leido = var_peso_caja Then
            var_posible_peso_caja = True
         Else
            var_porcentaje_peso = (var_peso_caja * 100) / var_peso_caja_leido
            If var_porcentaje_peso > 100 Then
               var_porcentaje_peso = var_porcentaje_peso - 100
            Else
               var_porcentaje_peso = 100 - var_porcentaje_peso
            End If
            If var_porcentaje_peso <= var_tolerancia_peso_caja Then
              var_posible_peso_caja = True
            Else
              var_posible_peso_caja = False
            End If
            'If var_peso_caja_leido < var_peso_caja Then
            '   var_peso_caja_leido = var_peso_caja_leido + var_tolerancia_peso_caja
            '   If var_peso_caja_leido >= var_peso_caja Then
            '      var_posible_peso_caja = True
            '   Else
            '      var_posible_peso_caja = False
            '   End If
            'Else
            '   var_peso_caja_leido = var_peso_caja_leido - var_tolerancia_peso_caja
            '   If var_peso_caja_leido <= var_peso_caja Then
            '      var_posible_peso_caja = True
            '   Else
            '      var_posible_peso_caja = False
            '   End If
            'End If
         End If
         If var_posible_peso_caja = True Then
            var_peso_correcto = True
            var_cantidad_leida = var_cantidad_caja_peso
            Me.txt_foco.Enabled = True
            Me.txt_foco.SetFocus
         Else
            var_peso_correcto = False
            Me.lbl_articulo_caja.Caption = "Art?culo contenido en la caja " + Trim(txt_codigo)
            frm_articulo_caja.Visible = True
            var_ventana = 1
            txt_codigo_caja.SetFocus
            Me.txt_codigo.Enabled = False
         End If
      Else
         MsgBox "Peso incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_peso_caja_LostFocus()
   frm_peso_caja.Visible = False
End Sub

Private Function pro_DatosTransito(numFolio As String, strAlmacenOri As String, strTipoMov As String) As Boolean
    
    Dim strSelTransito As String
    Dim strFromTransito As String
    Dim strWhereTransito As String
    Dim clsTransito As New clsTransito
On Error GoTo errorDatasSIP:
    'rs.Open "select VCHA_NOT_CLASIFICACION " & _
    '        "from tb_notas with(nolock) " & _
    '        "where bint_not_nota_id =" & numFolio, _
    '    cnn, _
    '    adOpenDynamic, _
    '    adLockOptimistic
    
        
    strSelTransito = "SELECT 'RECIBO' VCHA_TIPO_MOVIMIENTO, com.INTE_COM_CONSECUTIVO NUMB_TRA_CONSECUTIVO, " & _
                    "com.VCHA_COM_REFERENCIA VCHA_TRA_NOTA_ENVIO, '" & var_almacen_Destino & "' VCHA_TRA_PLANTA_DESTINO," & _
                    "convert(int , com.VCHA_COM_PROVEEDOR ) VCHA_TRA_PLANTA_ORIGEN, '" & var_almacen_Destino & "' VCHA_TRA_ALMACEN_DESTINO," & _
                    "com.VCHA_COM_REFERENCIA_ALMACEN VCHA_TRA_ALMACEN_ORIGEN, '' VCHA_EMP_EMPRESA_ORIGEN, " & _
                    "com.VCHA_EMP_EMPRESA_ID VCHA_EMP_EMPRESA_DESTINO,  '' VCHA_UOR_UNIDAD_ORIGEN," & _
                    "com.VCHA_UOR_UNIDAD_ID  VCHA_UOR_UNIDAD_DESTINO, '" & var_clave_movimiento & "' VCHA_MOV_MOVIMIENTO_ORIGEN," & _
                    "'EP' VCHA_MOV_MOVIMIENTO_DESTINO, '' VCHA_SER_SERIE_ORIGEN, '' VCHA_SER_SERIE_DESTINO," & _
                    "ENT.VCHA_ART_ARTICULO_ID VCHA_ART_ARTICULO_ORIGEN, ENT.VCHA_ART_ARTICULO_ID VCHA_ART_ARTICULO_DESTINO, " & _
                    "'' VCHA_ART_DESCRIPCION, COM.FLOA_COM_CANTIDAD_ENVIADA  NUMB_TRA_CANTIDAD_ENVIADA, ENT.FLOA_ENT_CANTIDAD NUMB_TRA_CANTIDAD_RECIBIDA, " & _
                    "0 NUMB_TRA_COSTO, '1' VCHA_TRA_CALIDAD,'' VCHA_TRA_SISTEMA_ORIGEN, 'SID' VCHA_TRA_SISTEMA_DESTINO," & _
                    "'A' VCHA_TRA_STATUS_ID, '' VCHA_TRA_USUARIO_ORIGEN, '" & var_usuario_global & "'  VCHA_TRA_USUARIO_DESTINO," & _
                    "'' VCHA_TRA_MAQUINA_origen, '" & fun_NombrePc & "' VCHA_TRA_MAQUINA_DESTINO, '' VCHA_TRA_REFERENCIA1, '' VCHA_TRA_REFERENCIA2," & _
                    "COM.CHAR_COM_TIPO_PROVEEDOR  CHAR_COM_TIPO_PROVEEDOR, '' VCHA_TRA_TRANSPORTO," & _
                    "COM.VCHA_COM_CAJA  VCHA_TRA_CONTENEDOR_ID, COM.FLOA_COM_PESO  NUMB_TRA_PESO," & _
                    "'MXP' VCHA_MON_MONEDA_ID, 0 NUMB_TRA_PRECIO, ENT.INTE_ENT_NUMERO numb_DTR_FOLIO_RECEPCION "
                        
    strFromTransito = "FROM TB_ARCHIVO_COMPARACION com with(nolock), " & _
                        "TB_ENTRADAS ENT WITH(NOLOCK) "
                            
    strWhereTransito = "WHERE ENT.INTE_ENT_NUMERO  = " & numFolio & _
                        "AND ENT.VCHA_ALM_ALMACEN_ID = '" & strAlmacenOri & "' " & _
                        "AND COM.VCHA_ART_ARTICULO_ID = ENT.VCHA_ART_ARTICULO_ID " & _
                        "AND ENT.VCHA_ALM_ALMACEN_ID  = COM.VCHA_ALM_ALMACEN_ID " & _
                        "AND ENT.VCHA_EMP_EMPRESA_ID  = COM.VCHA_EMP_EMPRESA_ID " & _
                        "AND ENT.VCHA_UOR_UNIDAD_ID  = COM.VCHA_UOR_UNIDAD_ID  " & _
                        "AND ENT.INTE_ENT_CONSECUTIVO = COM.INTE_COM_CONSECUTIVO "

    
    'rs.Close
    ok = clsTransito.fun_GuardaTransito(strSelTransito & strFromTransito & strWhereTransito)
Exit Function
errorDatasSIP:
    MsgBox Err.Description, vbCritical, "SID"
    If rs.State = 1 Then: rs.Close
    
End Function


