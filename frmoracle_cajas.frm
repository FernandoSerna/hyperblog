VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_cajas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empacado de mercancia"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   1380
      Top             =   30
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11130
      Picture         =   "frmoracle_cajas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   630
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      Picture         =   "frmoracle_cajas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Buscar Movimiento"
      Top             =   630
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   690
      Picture         =   "frmoracle_cajas.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cerrar Caja e Imprimir las Etiquetas"
      Top             =   630
      Width           =   330
   End
   Begin VB.CommandButton cmd_cerrar_embarque 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1020
      Picture         =   "frmoracle_cajas.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cerrar Embarque"
      Top             =   630
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmoracle_cajas.frx":0940
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   630
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir_reporte_faltantes 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1350
      Picture         =   "frmoracle_cajas.frx":0A42
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Imprimir reporte de faltantes"
      Top             =   630
      Width           =   330
   End
   Begin VB.CommandButton cmd_cerrar 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1680
      Picture         =   "frmoracle_cajas.frx":0B44
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "Cerrar Alt + C"
      Top             =   645
      Width           =   315
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   0
      TabIndex        =   26
      Top             =   870
      Width           =   11505
   End
   Begin VB.Frame frm_busqueda 
      Height          =   975
      Left            =   570
      TabIndex        =   53
      Top             =   930
      Width           =   3150
      Begin VB.TextBox txt_busqueda_caja 
         Height          =   315
         Left            =   1290
         TabIndex        =   54
         Top             =   495
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Busqueda de Caja"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   6
         Left            =   30
         TabIndex        =   56
         Top             =   120
         Width           =   3090
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Caja:"
         Height          =   195
         Left            =   195
         TabIndex        =   55
         Top             =   555
         Width           =   360
      End
   End
   Begin VB.CommandButton cmd_mensaje_2 
      Caption         =   "mensaje 2"
      Height          =   195
      Left            =   2160
      TabIndex        =   16
      Top             =   675
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton cmd_mensaje_4 
      Caption         =   "mensaje 4"
      Height          =   195
      Left            =   2325
      TabIndex        =   15
      Top             =   675
      Visible         =   0   'False
      Width           =   75
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   30
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
            Picture         =   "frmoracle_cajas.frx":0C46
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_cajas.frx":1520
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_cajas.frx":1DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_cajas.frx":2396
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_cajas.frx":2C72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_cajas.frx":354C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_cajas.frx":3E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_cajas.frx":3F38
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_cajas.frx":404A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_cajas.frx":415C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_cajas.frx":426E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_cajas.frx":4380
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_cajas.frx":4492
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_cajas.frx":4634
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_cajas.frx":5486
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_cajas.frx":565C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_cajas.frx":576E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   3585
      Left            =   60
      TabIndex        =   17
      Top             =   3660
      Width           =   11475
      Begin VB.CommandButton cmd_pasar_movimiento 
         Height          =   330
         Left            =   10515
         Picture         =   "frmoracle_cajas.frx":5880
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   480
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
         Left            =   1575
         TabIndex        =   12
         Top             =   435
         Width           =   3390
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   4440
         TabIndex        =   18
         Top             =   1575
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   75
            TabIndex        =   19
            Top             =   375
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H000000C0&
            Caption         =   "Cantidad a eliminar"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   20
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
         TabIndex        =   13
         Top             =   495
         Width           =   1890
      End
      Begin VB.TextBox txt_foco 
         Height          =   315
         Left            =   11655
         TabIndex        =   14
         Top             =   525
         Width           =   1650
      End
      Begin MSComctlLib.ListView lv_salidas 
         Height          =   2460
         Left            =   0
         TabIndex        =   22
         Top             =   1065
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   4339
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "   Código"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Posibles    "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Empacado "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Caja"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Faltan    "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "inventory item id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "delivery detail id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "source line number"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "delivery_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "customer_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Agente"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_impresa 
         Caption         =   "IMPRESA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Left            =   7980
         TabIndex        =   57
         Top             =   465
         Width           =   3180
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   5115
         TabIndex        =   25
         Top             =   615
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   615
         Width           =   1395
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Lectura de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   0
         Left            =   30
         TabIndex        =   23
         Top             =   120
         Width           =   11400
      End
   End
   Begin VB.Frame Frame3 
      Height          =   900
      Index           =   4
      Left            =   8520
      TabIndex        =   35
      Top             =   930
      Width           =   1620
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Cantidad empacada"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   5
         Left            =   30
         TabIndex        =   37
         Top             =   120
         Width           =   1545
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
         TabIndex        =   36
         Top             =   420
         Width           =   1320
      End
   End
   Begin VB.Frame Frame3 
      Height          =   915
      Index           =   3
      Left            =   6765
      TabIndex        =   32
      Top             =   915
      Width           =   1740
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Cantidad enviada"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   4
         Left            =   30
         TabIndex        =   34
         Top             =   120
         Width           =   1665
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
         Left            =   135
         TabIndex        =   33
         Top             =   420
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   900
      Index           =   2
      Left            =   10170
      TabIndex        =   49
      Top             =   930
      Width           =   1365
      Begin VB.Label lbl_cantidad_caja 
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
         Left            =   165
         TabIndex        =   51
         Top             =   420
         Width           =   1110
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Cantidad en caja"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   2
         Left            =   30
         TabIndex        =   50
         Top             =   120
         Width           =   1290
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1830
      Index           =   1
      Left            =   75
      TabIndex        =   38
      Top             =   1830
      Width           =   11460
      Begin VB.TextBox txt_orden_lectura 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         TabIndex        =   70
         Top             =   1410
         Width           =   1140
      End
      Begin VB.TextBox txt_entrega 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         TabIndex        =   68
         Top             =   1080
         Width           =   4230
      End
      Begin VB.TextBox txt_nombre_caja 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8730
         TabIndex        =   59
         Top             =   750
         Width           =   2580
      End
      Begin VB.TextBox txt_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7080
         TabIndex        =   10
         Top             =   420
         Width           =   4230
      End
      Begin VB.TextBox txt_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         TabIndex        =   11
         Top             =   750
         Width           =   4230
      End
      Begin VB.TextBox txt_archivo 
         Height          =   315
         Left            =   7080
         TabIndex        =   0
         Top             =   750
         Width           =   1170
      End
      Begin VB.TextBox txt_origen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         TabIndex        =   9
         Top             =   420
         Width           =   4230
      End
      Begin VB.Label lbl_bascula 
         Caption         =   "0000.000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   10440
         TabIndex        =   73
         Top             =   1530
         Width           =   915
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Peso en bascula:"
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
         Left            =   8580
         TabIndex        =   72
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Orden de lectura:"
         Height          =   195
         Left            =   165
         TabIndex        =   71
         Top             =   1470
         Width           =   1230
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Entrega:"
         Height          =   195
         Left            =   195
         TabIndex        =   69
         Top             =   1140
         Width           =   600
      End
      Begin VB.Label lbl_peso 
         AutoSize        =   -1  'True
         Caption         =   "0000.000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   7560
         TabIndex        =   65
         Top             =   1530
         Width           =   915
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Peso en bulto:"
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
         Left            =   5970
         TabIndex        =   64
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label lbl_maximo 
         AutoSize        =   -1  'True
         Caption         =   "0000.000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   4860
         TabIndex        =   63
         Top             =   1530
         Width           =   915
      End
      Begin VB.Label Label9 
         Caption         =   "Peso máximo:"
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
         Left            =   3360
         TabIndex        =   62
         Top             =   1500
         Width           =   1620
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Caja:"
         Height          =   195
         Left            =   8355
         TabIndex        =   58
         Top             =   810
         Width           =   360
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   6075
         TabIndex        =   43
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   195
         TabIndex        =   42
         Top             =   810
         Width           =   555
      End
      Begin VB.Label label 
         BackColor       =   &H000000C0&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   30
         TabIndex        =   41
         Top             =   120
         Width           =   11385
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   40
         Top             =   480
         Width           =   660
      End
      Begin VB.Label lbl_archivo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "O. de Surtido:"
         Height          =   195
         Left            =   6075
         TabIndex        =   39
         Top             =   810
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   15
      TabIndex        =   27
      Top             =   495
      Width           =   11505
   End
   Begin VB.Frame Frame3 
      Height          =   915
      Index           =   0
      Left            =   75
      TabIndex        =   28
      Top             =   915
      Width           =   6645
      Begin VB.TextBox txt_caja_pedido 
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
         Left            =   5790
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   375
         Width           =   780
      End
      Begin VB.TextBox txt_caja 
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
         Left            =   4005
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   375
         Width           =   780
      End
      Begin VB.TextBox txt_jaula 
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
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   375
         Width           =   705
      End
      Begin VB.TextBox txt_embarque 
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
         Left            =   885
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   375
         Width           =   1425
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Caja Pedido:"
         Height          =   195
         Left            =   4890
         TabIndex        =   61
         Top             =   518
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Caja:"
         Height          =   195
         Left            =   3615
         TabIndex        =   52
         Top             =   518
         Width           =   360
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Jaula:"
         Height          =   195
         Left            =   2355
         TabIndex        =   31
         Top             =   518
         Width           =   420
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   75
         TabIndex        =   30
         Top             =   518
         Width           =   765
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   1
         Left            =   30
         TabIndex        =   29
         Top             =   120
         Width           =   6570
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp4 
      Height          =   75
      Left            =   10170
      TabIndex        =   48
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
      Left            =   4695
      TabIndex        =   47
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
      Left            =   8475
      TabIndex        =   46
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
      Caption         =   "EMPACADO DE MERCANCIA"
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
      Left            =   60
      TabIndex        =   45
      Top             =   30
      Width           =   11445
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp2 
      Height          =   135
      Left            =   1470
      TabIndex        =   44
      Top             =   750
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
Attribute VB_Name = "frmoracle_cajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_lectura_flete As Integer
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_encontro As Integer
Dim var_cantidad_leida As Double
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim var_primera_vez As Integer
Dim var_renglon As Integer
Dim var_caja_pedido As Integer
Dim var_peso As Double
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter
Dim var_almacen_motor_logistico As String
Dim var_almacen_destino_caja As String
Dim var_caja_motor As String



Sub ilumina_grid()
   var_n = lv_salidas.ListItems.Count
   For var_i = 1 To var_n
       If var_i = var_renglon Then
          lv_salidas.ListItems.Item(var_i).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(3).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(4).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(5).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(6).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(7).Bold = True
          lv_salidas.ListItems.Item(var_i).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H8000&
       Else
          If (lv_salidas.ListItems.Item(var_i).ListSubItems(5) * 1) = 0 Then
             lv_salidas.ListItems.Item(var_i).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(1).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(2).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(3).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(4).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(5).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(6).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(7).Bold = False
             lv_salidas.ListItems.Item(var_i).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HFF&
          Else
             lv_salidas.ListItems.Item(var_i).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(1).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(2).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(3).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(4).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(5).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(6).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(7).Bold = False
             lv_salidas.ListItems.Item(var_i).ForeColor = &H80000012
             lv_salidas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000012
             lv_salidas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000012
             lv_salidas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000012
             lv_salidas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000012
             lv_salidas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H80000012
             lv_salidas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H80000012
             lv_salidas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H80000012
          End If
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_salidas.ListItems.Item(var_renglon).Selected = True
      lv_salidas.selectedItem.EnsureVisible
   End If
   lv_salidas.Refresh
End Sub


Private Sub ejecuta()
   Dim list_item As ListItem
   Dim var_tipo_pedido_embarque As String
   Dim var_posible_pedido_embarque As Integer
   Dim var_embarque_asignado
   Dim var_posible_continuar As Integer
   If IsNumeric(Me.txt_archivo) Then
      var_lectura_flete = 0
      If rs.State = 1 Then
         rs.Close
      End If
      var_posible_pedido_embarque = 1
      var_posibe_maquina = 1
      var_posible_continuar = 1
      var_almacen_motor_logistico = ""
      If var_bandera_asignacion = 0 Then
         rs.Open "select * from tb_oracle_pedidos_asignados_embarques where pedido = " + CStr(CDbl(Me.txt_archivo)), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If rs!Embarque <> Me.txt_embarque Then
               var_posible_pedido_embarque = 0
               var_embarque_asignado = rs!Embarque
            End If
         End If
         rs.Close
         rs.Open "SELECT * FROM tb_oracle_pedidos_maquinas where pedido = " + CStr(CDbl(Me.txt_archivo)), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If rs!maquina = fun_NombrePc Then
               var_posibe_maquina = 1
            Else
               var_posibe_maquina = 0
            End If
         Else
            var_posibe_maquina = 1
         End If
         rs.Close
      End If
      If var_posible_pedido_embarque = 1 Then
         If var_posibe_maquina = 1 Then
            'rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
            strconsulta = "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
                 .Parameters.Append parametro
            End With
            Set rs = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            
            If Not rs.EOF Then
               var_estatus_embarque = IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus)
               var_tipo_pedido_embarque = IIf(IsNull(rs!tipo_pedido), "", rs!tipo_pedido)
            Else
               var_estatus_embarque = "I"
            End If
            rs.Close
            If var_estatus_embarque = "" Then
               var_orden = CDbl(Me.txt_archivo)
               var_requisicion = ""
               VAR_ESTABLECIMIENTO = ""
               'rsaux7.Open "SELECT HEADER_ID, source_document_id, SHIP_TO_ORG_ID FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = " + CStr(var_orden), cnnoracle_4, adOpenDynamic, adLockOptimistic
               'var_orden = 548654
               strconsulta = "SELECT HEADER_ID, source_document_id, SHIP_TO_ORG_ID FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_orden)
                    .Parameters.Append parametro
               End With
               Set rsaux7 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               
               If Not rsaux7.EOF Then
                  VAR_HEADER_ID = IIf(IsNull(rsaux7!header_id), 0, rsaux7!header_id)
                  var_requisicion = IIf(IsNull(rsaux7!source_document_id), "", rsaux7!source_document_id)
                  VAR_ESTABLECIMIENTO = IIf(IsNull(rsaux7!ship_to_org_id), "0", rsaux7!ship_to_org_id)
               Else
                  VAR_HEADER_ID = 0
               End If
               rsaux7.Close
               
               rsaux7.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'SHIP_TO' AND csu.LOCATION = ps.party_site_number and csu.site_use_id = " + CStr(VAR_ESTABLECIMIENTO), cnnoracle_4, adOpenDynamic, adLockOptimistic
               
               If Not rsaux7.EOF Then
                  Me.txt_entrega = Replace(IIf(IsNull(rsaux7!vcha_esb_nombre), "", rsaux7!vcha_esb_nombre), "'", " ")
               Else
                  Me.txt_entrega = ""
               End If
               rsaux7.Close
              'lento
              'cambio bind
               'var_cadena = " SELECT a.source_header_type_name, HCAS.CUST_ACCOUNT_ID, HCAS.CUST_ACCT_SITE_ID as customer_id, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, oha.attribute8, oha.attribute9 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) = " + CStr(var_orden)
               'var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID  AND released_status = 'Y' AND A.SOURCE_HEADER_ID = " + CStr(VAR_HEADER_ID)
                              
               var_cadena = " SELECT a.source_header_type_name, HCAS.CUST_ACCOUNT_ID, HCAS.CUST_ACCT_SITE_ID as customer_id, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, oha.attribute8, oha.attribute9 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) = ?"
               var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID  AND released_status = 'Y' AND A.SOURCE_HEADER_ID = ?"
               
               strconsulta = var_cadena
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(var_orden))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(VAR_HEADER_ID))
                    .Parameters.Append parametro
                End With
                Set rs = comandoORA.execute
                Set comandoORA = Nothing
                Set parametro = Nothing
               
               
               
               'rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If var_tipo_pedido_embarque = "" Then
                  var_tipo_pedido_embarque = rs!source_header_type_name
                  rsaux.Open "update XXVIA_TB_ENCABEZADO_EMBARQUES set tipo_pedido = '" + rs!source_header_type_name + "' WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               End If
               If Not rs.EOF Then
                  If rs!source_header_type_name = var_tipo_pedido_embarque Then
                     If rsaux.State = 1 Then
                        rsaux.Close
                     End If
                     
                      If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Then
                         rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(var_requisicion) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                         If Not rsaux2.EOF Then
                            var_almacen_motor_logistico = rsaux2!attribute1
                         Else
                            var_almacen_motor_logistico = ""
                         End If
                         rsaux2.Close
                      End If
                     
                     
                     
                     rsaux.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     VAR_AGENTE_str = IIf(IsNull(rsaux!collector_id), "", rsaux!collector_id)
                     var_nombre_agente_str = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                     rsaux.Close
                     'cambio bind
                     'rsaux.Open "SELECT oha.header_id, oha.ordered_date, oha.order_number,  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  E.NAME, f.orig_system_reference from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors E, hz_cust_acct_sites_all f Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id and oha.order_type_id in (1106) and HCSU.site_use_code = 'BILL_TO' and f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID and order_number  = '" + Me.txt_archivo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_cadena = "SELECT oha.header_id, oha.ordered_date, oha.order_number,  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  E.NAME, f.orig_system_reference from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors E, hz_cust_acct_sites_all f Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id and oha.order_type_id in (1106) and HCSU.site_use_code = 'BILL_TO' and f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID and order_number  = ?"
                     strconsulta = var_cadena
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(CDbl(Me.txt_archivo)))
                          .Parameters.Append parametro
                     End With
                     Set rsaux = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     
                     If Not rsaux.EOF Then
                        rsaux1.Open "select * from OE_ORDER_HOLDS_ALL where header_id = " + CStr(rsaux!header_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux1.EOF Then
                           var_estatus_vxt = IIf(IsNull(rsaux1!released_flag), "N", rsaux1!released_flag)
                        Else
                           var_estatus_vxt = "Y"
                        End If
                        rsaux1.Close
                        If var_estatus_vxt <> "Y" Then
                           var_posible_ventas_x_telefono = 0
                        Else
                           var_posible_ventas_x_telefono = 1
                        End If
                     Else
                        var_posible_ventas_x_telefono = 1
                     End If
                     rsaux.Close
                     If var_posible_ventas_x_telefono = 1 Then
                        rsaux.Open "SELECT * FROM TB_ORACLE_EMBARQUES_ORDENES WHERE source_header_number = " + CStr(var_orden), cnn, adOpenDynamic, adLockOptimistic
                        If rsaux.EOF Then
                           var_primera_vez = 1
                           Me.txt_agente = var_nombre_agente_str
                           If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rs!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                              If var_pedido_tienda = 0 Then
                              
                                 Me.txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                                 rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(var_requisicion) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux2.EOF Then
                                    Me.txt_entrega = Replace(IIf(IsNull(rsaux2!Description), "", rsaux2!Description), "'", " ")
                                 End If
                                 rsaux2.Close
                              Else
                                 Me.txt_cliente = IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9)
                              End If
                           Else
                              Me.txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                           End If
                           Me.txt_origen = IIf(IsNull(rs!subinventory), "", rs!subinventory)
                           Me.lv_salidas.ListItems.Clear
                           var_cantidad_enviada = 0
                           While Not rs.EOF
                                  Set list_item = lv_salidas.ListItems.Add(, , rs!SEGMENT1)
                                 list_item.SubItems(1) = IIf(IsNull(rs!Description), "", rs!Description)
                                 list_item.SubItems(2) = Format(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity), "###,###,##0.00")
                                 var_cantidad_enviada = var_cantidad_enviada + IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)
                                 list_item.SubItems(3) = 0
                                 list_item.SubItems(4) = 0
                                 list_item.SubItems(5) = 0
                                 list_item.SubItems(6) = IIf(IsNull(rs!inventory_item_id), 0, rs!inventory_item_id)
                                 list_item.SubItems(7) = IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)
                                 list_item.SubItems(8) = IIf(IsNull(rs!SOURCE_LINE_NUMBER), 0, rs!SOURCE_LINE_NUMBER)
                                 list_item.SubItems(9) = IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)
                                 list_item.SubItems(10) = IIf(IsNull(rs!CUSTOMER_ID), 0, rs!CUSTOMER_ID)
                                 list_item.SubItems(11) = VAR_AGENTE_str
                                 rs.MoveNext
                           Wend
                           
                           rsaux2.Open "select * from tb_oracle_pedidos_asignados_embarques where pedido = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux2.EOF Then
                              Me.txt_orden_lectura = IIf(IsNull(rsaux2!orden_pedido), "", rsaux2!orden_pedido)
                           Else
                              Me.txt_orden_lectura = ""
                           End If
                           rsaux2.Close
                           
                           Me.lbl_enviados = Format(var_cantidad_enviada, "###,###,##0.00")
                           Me.lbl_recibidos = Format(0, "###,###,##0.00")
                           Me.lbl_cantidad_caja = Format(0, "###,###,##0.00")
                           Me.txt_archivo.Enabled = False
                           var_cantidad_recibida = 0
                           'rsaux2.Open "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND  source_header_number = " + CStr(CDbl(Me.txt_archivo)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           strconsulta = "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = ? AND  source_header_number = ?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_archivo))
                                .Parameters.Append parametro
                           End With
                           Set rsaux2 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           
                           If Not rsaux2.EOF Then
                              While Not rsaux2.EOF
                                    var_codigo = rsaux2!SEGMENT1
                                    For var_j = 1 To Me.lv_salidas.ListItems.Count
                                        Me.lv_salidas.ListItems.Item(var_j).Selected = True
                                        If CDbl(Me.lv_salidas.selectedItem.SubItems(7)) = CDbl(rsaux2!delivery_detail_id) Then
                                           Me.lv_salidas.selectedItem.SubItems(3) = CDbl(Me.lv_salidas.selectedItem.SubItems(3)) + Format(rsaux2!FLOA_SAL_CANTIDAD_LEIDA, "###,###,##0.00")
                                           Me.lv_salidas.selectedItem.SubItems(5) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(2)) - CDbl(Me.lv_salidas.selectedItem.SubItems(3)), "###,###,##0.00")
                                        End If
                                    Next var_j
                                    var_cantidad_recibida = var_cantidad_recibida + rsaux2!FLOA_SAL_CANTIDAD_LEIDA
                                    rsaux2.MoveNext
                              Wend
                           Else
                              For var_j = 1 To Me.lv_salidas.ListItems.Count
                                  Me.lv_salidas.ListItems.Item(var_j).Selected = True
                                  Me.lv_salidas.selectedItem.SubItems(5) = Format(Me.lv_salidas.selectedItem.SubItems(2), "###,###,##0.00")
                              Next var_j
                           End If
                           rsaux2.Close
                           Me.lbl_recibidos = Format(var_cantidad_recibida, "###,###,##0.00")
                           'pago flete
                           For var_j = 1 To Me.lv_salidas.ListItems.Count
                               Me.lv_salidas.ListItems.Item(var_j).Selected = True
                               If Me.lv_salidas.selectedItem = "00004434" And CDbl(Me.lv_salidas.selectedItem.SubItems(3)) > 0 Then
                                  var_lectura_flete = 1
                               Else
                                  var_lectura_flete = 0
                               End If
                           Next var_j
                           frmoracle_tipo_cajas.Show 1
                           Me.txt_nombre_caja = var_nombre_caja
                           rsaux7.Open "select * from tb_oracle_empaques where empaque = '" + Me.txt_nombre_caja + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux7.EOF Then
                              Me.lbl_maximo = Format(IIf(IsNull(rsaux7!PESO), 0, rsaux7!PESO), "###,###,##0.000")
                           Else
                              Me.lbl_maximo = "0.000"
                           End If
                           rsaux7.Close
                        
                           Me.txt_codigo.Enabled = True
                           Me.txt_codigo.SetFocus
                        Else
                           var_primera_vez = 1
                           If rsaux!inte_Emb_Embarque = CDbl(Me.txt_embarque) Or rsaux.EOF Then
                              Me.txt_agente = var_nombre_agente_str
                              If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rs!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                                 If var_pedido_tienda = 0 Then
                                    Me.txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                                 Else
                                    Me.txt_cliente = IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9)
                                 End If
                              Else
                                 Me.txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                              End If
                              Me.txt_origen = IIf(IsNull(rs!subinventory), "", rs!subinventory)
                              Me.lv_salidas.ListItems.Clear
                              var_cantidad_enviada = 0
                              While Not rs.EOF
                                    Set list_item = lv_salidas.ListItems.Add(, , rs!SEGMENT1)
                                    list_item.SubItems(1) = IIf(IsNull(rs!Description), "", rs!Description)
                                    list_item.SubItems(2) = Format(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity), "###,###,##0.00")
                                    var_cantidad_enviada = var_cantidad_enviada + IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)
                                    list_item.SubItems(3) = 0
                                    list_item.SubItems(4) = 0
                                    list_item.SubItems(5) = IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)
                                    list_item.SubItems(6) = IIf(IsNull(rs!inventory_item_id), 0, rs!inventory_item_id)
                                    list_item.SubItems(7) = IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)
                                    list_item.SubItems(8) = IIf(IsNull(rs!SOURCE_LINE_NUMBER), 0, rs!SOURCE_LINE_NUMBER)
                                    list_item.SubItems(9) = IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)
                                    list_item.SubItems(10) = IIf(IsNull(rs!CUSTOMER_ID), 0, rs!CUSTOMER_ID)
                                    list_item.SubItems(11) = VAR_AGENTE_str
                                    rs.MoveNext
                              Wend
                              var_cantidad_recibida = 0
                              'rsaux2.Open "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND  source_header_number = " + CStr(CDbl(Me.txt_archivo)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              strconsulta = "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = ? AND  source_header_number = ?"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_archivo))
                                   .Parameters.Append parametro
                              End With
                              Set rsaux2 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                              
                              While Not rsaux2.EOF
                                    VAR_DELIVERY_ID = rsaux2!delivery_detail_id
                                    For var_j = 1 To Me.lv_salidas.ListItems.Count
                                        Me.lv_salidas.ListItems.Item(var_j).Selected = True
                                        If CDbl(Me.lv_salidas.selectedItem.SubItems(7)) = CDbl(rsaux2!delivery_detail_id) Then
                                           Me.lv_salidas.selectedItem.SubItems(3) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(3)) + rsaux2!FLOA_SAL_CANTIDAD_LEIDA, "###,###,##0.00")
                                           Me.lv_salidas.selectedItem.SubItems(5) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(2)) - CDbl(Me.lv_salidas.selectedItem.SubItems(3)), "###,###,##0.00")
                                        End If
                                    Next var_j
                                    var_cantidad_recibida = var_cantidad_recibida + rsaux2!FLOA_SAL_CANTIDAD_LEIDA
                                    rsaux2.MoveNext
                              Wend
                              rsaux2.Close
                              'pago flete
                              For var_j = 1 To Me.lv_salidas.ListItems.Count
                                  Me.lv_salidas.ListItems.Item(var_j).Selected = True
                                  If Me.lv_salidas.selectedItem = "00004434" And CDbl(Me.lv_salidas.selectedItem.SubItems(3)) > 0 Then
                                     var_lectura_flete = 1
                                  Else
                                     var_lectura_flete = 0
                                  End If
                              Next var_j
                              
                              rsaux2.Open "select * from tb_oracle_pedidos_asignados_embarques where pedido = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 Me.txt_orden_lectura = IIf(IsNull(rsaux2!orden_pedido), "", rsaux2!orden_pedido)
                              Else
                                 Me.txt_orden_lectura = ""
                              End If
                              rsaux2.Close
                              Me.lbl_recibidos = Format(var_cantidad_recibida, "###,###,##0.00")
                              Me.lbl_enviados = Format(var_cantidad_enviada, "###,###,##0.00")
                              Me.lbl_cantidad_caja = Format(0, "###,###,##0.00")
                              Me.txt_archivo.Enabled = False
                              frmoracle_tipo_cajas.Show 1
                              Me.txt_nombre_caja = var_nombre_caja
                              
                              rsaux7.Open "select * from tb_oracle_empaques where empaque = '" + Me.txt_nombre_caja + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux7.EOF Then
                                 Me.lbl_maximo = Format(IIf(IsNull(rsaux7!PESO), 0, rsaux7!PESO), "###,###,##0.000")
                              Else
                                 Me.lbl_maximo = "0.000"
                              End If
                              rsaux7.Close
                              
                              Me.txt_codigo.Enabled = True
                              Me.txt_codigo.SetFocus
                           Else
                              rsaux1.Open "SELECT * FROM TB_ORACLE_EMBARQUES_ORDENES WHERE SOURCE_HEADER_NUMBER = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
                              'rsaux1.Open "SELECT dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_EMBARQUES.INTE_JAU_JAULA_ID, dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID, dbo.TB_USUARIOS.VCHA_USU_NOMBRE, dbo.TB_ENCABEZADO_EMBARQUES.VCHA_AUD_MAQUINA, dbo.Tb_usuarios.VCHA_USU_APELLIDOS FROM dbo.TB_ENCABEZADO_EMBARQUES INNER JOIN dbo.TB_USUARIOS ON dbo.TB_ENCABEZADO_EMBARQUES.VCHA_AUD_USUARIO = dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID Where (dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE = " + CStr(rsaux!INTE_EMB_EMBARQUE) + ")", cnn, adOpenDynamic, adLockOptimistic
                              'MsgBox "La orden de surtido se encuentra en el embarque " + CStr(rsaux1!INTE_EMB_EMBARQUE) + " en la máquina " + IIf(IsNull(rsaux1!vcha_aud_maquina), "", rsaux1!vcha_aud_maquina) + " con el usuario " + IIf(IsNull(rsaux1!VCHA_USU_NOMBRE), "", rsaux1!VCHA_USU_NOMBRE) + " " + IIf(IsNull(rsaux1!VCHA_USU_APELLIDOS), "", rsaux1!VCHA_USU_APELLIDOS), vbOKOnly, "ATENCION"
                              MsgBox "La orden de surtido se encuentra en el embarque " + CStr(rsaux1!inte_Emb_Embarque), vbOKOnly, "ATENCION"
                              rsaux1.Close
                              Me.txt_agente = ""
                              Me.txt_archivo = ""
                              Me.txt_cliente = ""
                              Me.txt_origen = ""
                              Me.lbl_enviados = ""
                              Me.lbl_recibidos = ""
                              Me.txt_entrega = ""
                              Me.txt_orden_lectura = ""
                              Me.txt_codigo.Enabled = False
                              Me.lv_salidas.ListItems.Clear
                           End If
                        End If
                        rsaux.Close
                     Else
                        MsgBox "La orden de surtido pertenece a ventas por teléfono y no a sido liberada", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "No es posible mezclar tipos de pedidos diferentes", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "La orden de surtido no existe", vbOKOnly, "ATENCION"
                  Me.txt_agente = ""
                  Me.txt_archivo = ""
                  Me.txt_cliente = ""
                  Me.txt_origen = ""
                  Me.lbl_enviados = ""
                  Me.lbl_recibidos = ""
                  Me.txt_entrega = ""
                  Me.txt_orden_lectura = ""
                  Me.txt_codigo.Enabled = False
                  Me.lv_salidas.ListItems.Clear
               End If
               rs.Close
            Else
            End If
         Else
            MsgBox "El pedido ya esta siendo utilizado en la máquina", vbOKOnly, "ATENCION"
         End If
      Else
         If var_embarque_asignado = 0 Then
            MsgBox "El pedido no a sido asignado", vbOKOnly, "ATENCION"
         Else
            MsgBox "El pedido se encuentra asignado al embarque " + CStr(var_embarque_asignado), vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "Número de orden de surtido incorrecta", vbOKOnly, "ATENCION"
      Me.txt_agente = ""
      Me.txt_archivo = ""
      Me.txt_cliente = ""
      Me.txt_origen = ""
      Me.txt_entrega = ""
      Me.txt_orden_lectura = ""
      Me.txt_codigo.Enabled = False
      Me.lv_salidas.ListItems.Clear
   End If
End Sub



Private Sub cmd_buscar_Click()
   Me.frm_busqueda.Visible = True
   Me.txt_busqueda_caja = ""
   Me.txt_busqueda_caja.SetFocus
End Sub

Private Sub cmd_cerrar_Click()
   If Me.lv_salidas.ListItems.Count > 0 Then
      strconsulta = "select distinct inte_paq_caja from xxvia_tb_salidas_cajas where source_header_number = ? and char_paq_estatus is null and floa_sal_cantidad_leida > 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.txt_archivo))
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If rsaux9.EOF Then
         strconsulta = "select nvl(estatus_pedido,0) from xxvia_tb_salidas_cajas where source_header_number = ? "
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.txt_archivo))
              .Parameters.Append parametro
         End With
         Set rsaux8 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rsaux8.EOF Then
            If rsaux8(0).Value = 0 Then
               'If var_bandera_asignacion = 0 Then
                  var_si = MsgBox("     ¿Desea cerrar el pedido?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     var_si = MsgBox("Confirmar el cerrado del pedido", vbYesNo, "ATENCION")
                     If var_si = 6 Then
                        'var_cantidad_negar = 0
                        'For var_j = 1 To Me.lv_salidas.ListItems.Count
                        '    Me.lv_salidas.ListItems.Item(var_j).Selected = True
                        '    var_cantidad_negar = var_cantidad_negar + CDbl(Me.lv_salidas.selectedItem.SubItems(5))
                        'Next var_j
                        'If var_cantidad_negar > 0 Then
                           strconsulta = "delete from xxvia_tb_negado_distribucion where SOURCE_HEADER_NUMBER = ?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.txt_archivo))
                                .Parameters.Append parametro
                           End With
                           Set rsaux8 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           
                           For var_j = 1 To Me.lv_salidas.ListItems.Count
                               Me.lv_salidas.ListItems.Item(var_j).Selected = True
                               'If CDbl(Me.lv_salidas.selectedItem.SubItems(5)) > 0 Then
                                  strconsulta = "insert into xxvia_tb_negado_distribucion (DELIVERY_DETAIL_ID, INVENTORY_ITEM_ID, SOURCE_HEADER_NUMBER, SEGMENT1, FECHA_NEGADO, Cantidad, ORGANIZATION_ID, CANTIDAD_PEDIDA, CANTIDAD_SURTIDA) values (?, ?, ?, ?, sysdate, ?, ?, ?, ?)"
                                  With comandoORA
                                       .ActiveConnection = cnnoracle_4
                                       .CommandType = adCmdText
                                       .CommandText = strconsulta
                                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.lv_salidas.selectedItem.SubItems(7)))
                                       .Parameters.Append parametro
                                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.lv_salidas.selectedItem.SubItems(6)))
                                       .Parameters.Append parametro
                                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.txt_archivo))
                                       .Parameters.Append parametro
                                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.lv_salidas.selectedItem)
                                       .Parameters.Append parametro
                                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.lv_salidas.selectedItem.SubItems(5)))
                                       .Parameters.Append parametro
                                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_unidad_organizacional))
                                       .Parameters.Append parametro
                                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.lv_salidas.selectedItem.SubItems(2)))
                                       .Parameters.Append parametro
                                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.lv_salidas.selectedItem.SubItems(3)))
                                       .Parameters.Append parametro
                                  End With
                                  Set rsaux8 = comandoORA.execute
                                  Set comandoORA = Nothing
                                  Set parametro = Nothing
                               'End If
                           Next var_j
                           
REPETIR:
                           var_orden_depurar = CDbl(Me.txt_archivo)
                           var_lote_depurar = 0
                           strconsulta = "select * from xxvia_tb_negado_distribucion where SOURCE_HEADER_NUMBER = ? and nvl(causa_negado,' ') = ' ' and cantidad > 0"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_orden_depurar))
                                .Parameters.Append parametro
                           End With
                           Set rsaux10 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux10.EOF Then
                              frmoracle_lineas_depurar.Show 1
                           End If
                        'End If
                        If var_lote_depurar = 0 Then
                           strconsulta = "select a.DELIVERY_DETAIL_ID, a.INVENTORY_ITEM_ID, a.SOURCE_HEADER_NUMBER, a.SEGMENT1 as codigo, a.FECHA_NEGADO, nvl(a.CAUSA_NEGADO,' ') as causa_negado, a.NOMBRE_CAUSA_NEGADO, a.Cantidad, a.ORGANIZATION_ID, a.LOTE, b.description as descripcion from xxvia_tb_negado_distribucion a, xxvia_system_items_b b where SOURCE_HEADER_NUMBER = ? and a.inventory_item_id = b.inventory_item_id and a.organization_id = b.organization_id and nvl(causa_negado,' ') = ' ' and cantidad > 0"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_orden_depurar))
                                .Parameters.Append parametro
                           End With
                           Set rsaux8 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                        End If
                        If rsaux8.EOF Then
                           rsaux.Open "update XXVIA_TB_SALIDAS_CAJAS set estatus_pedido = 1 WHERE SOURCE_HEADER_NUMBER = " + Me.txt_archivo, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           rsaux.Open "UPDATE tb_oracle_pedidos_asignados_embarques SET ESTATUS_PEDIDO = 1 WHERE PEDIDO = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
                           Me.txt_codigo.Enabled = False
                           MsgBox "Se a cambiado el estatus del pedido", vbOKOnly, "ATENCION"
                        Else
                           var_si = MsgBox("No se han asignado todas las causas de negado, ¿Desea terminar de asignar las causas de negado?", vbYesNo, "ATENCION")
                           If var_si = 6 Then
                              GoTo REPETIR:
                           Else
                              MsgBox "Se han eliminado las causas de negado", vbOKOnly, "ATENCION"
                           End If
                        End If
                     End If
                  End If
               'End If
            Else
               MsgBox "El pedido ya habia sido cerrado", vbOKOnly, "ATENCION"
            End If
         End If
      Else
         var_cadena_cajas = ""
         var_i = 0
         While Not rsaux9.EOF
               If var_cadena_cajas = "" Then
                  var_cadena_cajas = CStr(rsaux9(0).Value)
               Else
                  var_cadena_cajas = var_cadena_cajas + ", " + CStr(rsaux9(0).Value)
               End If
               var_i = var_i + 1
               rsaux9.MoveNext
         Wend
         If var_i > 1 Then
            MsgBox "Las siguientes cajas no estan cerradas: " + var_cadena_cajas, vbOKOnly, "ATENCION"
         Else
            MsgBox "La siguiente caja no esta cerrada: " + var_cadena_cajas, vbOKOnly, "ATENCION"
         End If
      End If
      rsaux9.Close
   Else
      MsgBox "Debe de seleccionar el pedido", vbOKOnly, "ATENCION"
   End If
   If rs.State = 1 Then
      rs.Close
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
   If rsaux12.State = 1 Then
      rsaux12.Close
   End If
   If rsaux13.State = 1 Then
      rsaux13.Close
   End If
   If rsaux14.State = 1 Then
      rsaux14.Close
   End If
   If rsaux15.State = 1 Then
      rsaux15.Close
   End If
   
End Sub

Private Sub cmd_cerrar_embarque_Click()
   If var_estatus_embarque = "" Then
      rsaux10.Open "select * from TB_ORACLE_EMBARQUES_ORDENES where INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
      var_Cadena_pedidos = ""
      While Not rsaux10.EOF
            If var_Cadena_pedidos = "" Then
               var_Cadena_pedidos = CStr(rsaux10!source_header_number)
            Else
               var_Cadena_pedidos = var_Cadena_pedidos + ", " + CStr(rsaux10!source_header_number)
            End If
            rsaux10.MoveNext
      Wend
      rsaux10.Close
      rsaux10.Open "select distinct source_header_number from xxvia_tb_salidas_cajas where source_header_number in (" + var_Cadena_pedidos + ") and nvl(estatus_pedido,0) <> 1", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If rsaux10.EOF Then
         var_si = MsgBox("¿Deseas ya cerrar el embarque?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar el cerrado de el embarque", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_estatus_embarque = "E"
               If rs.State = 1 Then
                  rs.Close
               End If
               If var_bandera_asignacion <> 0 Then
                  rs.Open "update XXVIA_TB_SALIDAS_CAJAS set char_paq_estatus = 'I' where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               Else
                  rs.Open "update XXVIA_TB_SALIDAS_CAJAS set char_paq_estatus = 'I' where inte_emb_embarque = " + Me.txt_embarque + " and char_paq_estatus<>'S'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               End If
               'rs.Open "select * from XXVIA_TB_SALIDAS_CAJAS where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               strconsulta = "select * from XXVIA_TB_SALIDAS_CAJAS where inte_emb_embarque = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_embarque))
                    .Parameters.Append parametro
               End With
               Set rs = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
            
               var_posible_cerrado = True
               While Not rs.EOF
                     VAR_ESTATUS = IIf(IsNull(rs!char_paq_estatus), "", rs!char_paq_estatus)
                     If VAR_ESTATUS = "" Then
                        var_posible_cerrado = False
                     End If
                     rs.MoveNext
               Wend
               rs.Close
               If var_posible_cerrado = True Then
                  rs.Open "UPDATE XXVIA_TB_ENCABEZADO_EMBARQUEs SET CHAR_EMB_ESTATUS = 'E' WHERE EMBARQUE = " + txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  MsgBox "El embarque a sido cerrado", vbOKOnly, "ATENCION"
               Else
                  MsgBox "No se puede cerrar el embarque ya que existen cajas sin cerrar", vbOKOnly, "ATENCION"
               End If
            End If
         End If
      Else
         var_Cadena_pedidos = ""
         var_i = 0
         While Not rsaux10.EOF
               If var_Cadena_pedidos = "" Then
                  var_Cadena_pedidos = CStr(rsaux10(0).Value)
               Else
                  var_Cadena_pedidos = var_Cadena_pedidos + ", " + CStr(rsaux10(0).Value)
               End If
               var_i = var_i´ + 1
               rsaux10.MoveNext
         Wend
         If var_i > 1 Then
            MsgBox "Falta cerrar los siguientes pedidos: " + var_Cadena_pedidos
         Else
            MsgBox "Falta cerrar el siguiente pedido: " + var_Cadena_pedidos
         End If
      End If
      rsaux10.Close
   Else
      MsgBox "El embarque ya no puede ser cerrado", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
   Dim clnt As New SoapClient30
   Dim var_referencia_vi As String
   Dim var_contador_renglones As Integer
   Dim var_numero_etiqueta As Integer
   Dim var_longitud As Integer
   Dim var_articulo As String
   Dim var_referencia_caja As String
   Dim var_contador As Integer
   Dim var_cantidad_total As String
   Dim var_cantidad_caja_impresion As Double
   Dim var_cliente_coppel As String
   Dim var_posible_sello As Boolean
   Dim var_bari As Integer
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
   If rsaux12.State = 1 Then
      rsaux12.Close
   End If
   If rsaux13.State = 1 Then
      rsaux13.Close
   End If
   If rsaux14.State = 1 Then
      rsaux14.Close
   End If
   If rsaux15.State = 1 Then
      rsaux15.Close
   End If
   
   
   If IsNumeric(Me.txt_caja) Then
      var_numero_caja = CDbl(Me.txt_caja)
      var_referencia_caja = ""
      var_contador = 0
      If Len(Trim(Str(var_numero_caja))) = 1 Then
         var_referencia_caja = "00" + Trim(Str(var_numero_caja))
      End If
      If Len(Trim(Str(var_numero_caja))) = 2 Then
         var_referencia_caja = "0" + Trim(Str(var_numero_caja))
      End If
      If Len(Trim(Str(var_numero_caja))) = 3 Then
         var_referencia_caja = Trim(Str(var_numero_caja))
      End If
      If Len(Trim(Str(txt_embarque))) = 1 Then
         var_referencia_embarque = "00000" + Trim(Str(txt_embarque))
      End If
      If Len(Trim(Str(txt_embarque))) = 2 Then
         var_referencia_embarque = "0000" + Trim(Str(txt_embarque))
      End If
      If Len(Trim(Str(txt_embarque))) = 3 Then
         var_referencia_embarque = "000" + Trim(Str(txt_embarque))
      End If
      If Len(Trim(Str(txt_embarque))) = 4 Then
         var_referencia_embarque = "00" + Trim(Str(txt_embarque))
      End If
      If Len(Trim(Str(txt_embarque))) = 5 Then
         var_referencia_embarque = "0" + Trim(Str(txt_embarque))
      End If
      If Len(Trim(Str(txt_embarque))) = 6 Then
         var_referencia_embarque = Trim(Str(txt_embarque))
      End If
      VAR_CAJA_AUTORIZA = "C" + var_referencia_embarque + var_referencia_caja
   
      rsaux12.Open "select * from tb_oracle_impresion_etiquetas where caja = '" + VAR_CAJA_AUTORIZA + "'", cnn, adOpenDynamic, adLockOptimistic
      var_autoriza_REIMPRESION = 0
      If Not rsaux12.EOF Then
         frmoracle_autoriza_reimpresion_etiquetas_cajas.Show 1
      Else
         rsaux11.Open "INSERT INTO tb_oracle_impresion_etiquetas (CAJA) VALUES ('" + VAR_CAJA_AUTORIZA + "')", cnn, adOpenDynamic, adLockOptimistic
         var_autoriza_REIMPRESION = 1
      End If
      rsaux12.Close
      If var_autoriza_REIMPRESION = 1 Then
         var_numero_caja = CDbl(Me.txt_caja)
         var_cantidad_caja_impresion = 0
         var_estatus_movimiento = "I"
         var_bari = 0
         If var_estatus_movimiento = "I" Then
            'rs.Open "select * from XXVIA_TB_SALIDAS_CAJAS where source_header_number = " + txt_archivo + " and inte_paq_caja = " + Me.txt_caja + " AND INTE_EMB_EMBARQUE = " + Me.txt_embarque + " and floa_Sal_cantidad_leida > 0 ", cnnoracle_4, adOpenDynamic, adLockOptimistic
            strconsulta = "select * from XXVIA_TB_SALIDAS_CAJAS where source_header_number = ? and inte_paq_caja = ? AND INTE_EMB_EMBARQUE = ? and floa_Sal_cantidad_leida > 0 "
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_archivo))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_caja))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
                 .Parameters.Append parametro
            End With
            Set rs = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
         
            
            If IsNumeric(Me.lbl_recibidos) Then
               var_cantidad_total = CStr(CDbl(Me.lbl_recibidos))
            Else
               var_cantidad_total = ""
            End If
            If Not rs.EOF Then
               While Not rs.EOF
                     strconsulta = "select * from xxvia_vw_categorias_item_b where organization_id = ? and cat_mex like '%BIASI%' AND CODIGO = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_unidad_organizacional))
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 10, rs!SEGMENT1)
                          .Parameters.Append parametro
                     End With
                     Set rsaux1 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If Not rsaux1.EOF Then
                        var_bari = 1
                     End If
                     rsaux1.Close
                     rs.MoveNext
               Wend
               rs.Close
               strconsulta = "select * from XXVIA_TB_SALIDAS_CAJAS where source_header_number = ? and inte_paq_caja = ? AND INTE_EMB_EMBARQUE = ? and floa_Sal_cantidad_leida > 0 "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_archivo))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_caja))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
                    .Parameters.Append parametro
               End With
               Set rs = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               
               var_sello_caja = IIf(IsNull(rs!sello), "", rs!sello)
               var_si = 6
               var_estatus_movimiento = IIf(IsNull(rs!char_paq_estatus), "", rs!char_paq_estatus)
               If var_estatus_movimiento = "" Then
                  var_si = MsgBox("¿Desea cerrar la caja e imprimir el movimiento?", vbYesNo, "ATENCION")
               End If
               If var_si = 6 Then
                  If var_sello_caja = "" Then
                     frmoracle_sello_caja.Show 1
                  End If
                  If Me.txt_nombre_caja = "COSTAL GRANDE" Or Me.txt_nombre_caja = "COSTAL CHICO" Or Me.txt_nombre_caja = "COSTAL EXTRAGRANDE" Then
                     If Len(var_sello_caja) >= 6 Then
                        var_tipo_caja_sello = var_sello_caja
                     Else
                        var_tipo_caja_sello = ""
                     End If
                  Else
                     var_tipo_caja_sello = "x"
                  End If
                  If var_tipo_caja_sello <> "" Then
                  
                  rsaux.Open "UPDATE XXVIA_TB_SALIDAS_CAJAS SET CHAR_PAQ_ESTATUS = 'I', sello = '" + var_sello_caja + "' where source_header_number = " + txt_archivo + " and inte_paq_caja = " + Me.txt_caja + " AND INTE_EMB_EMBARQUE = " + Me.txt_embarque + " and floa_Sal_cantidad_leida > 0 ", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  Me.txt_codigo.Enabled = False
                  'If var_bandera_asignacion <> 0 Then
                  Set fs = CreateObject("Scripting.FileSystemObject")
                  Set A = fs.CreateTextFile(App.Path + "\etiquetas.txt", True)
                  var_numero_caja = rs!INTE_PAQ_CAJA
                  var_referencia_caja = ""
                  var_contador = 0
                  If Len(Trim(Str(var_numero_caja))) = 1 Then
                     var_referencia_caja = "00" + Trim(Str(var_numero_caja))
                  End If
                  If Len(Trim(Str(var_numero_caja))) = 2 Then
                     var_referencia_caja = "0" + Trim(Str(var_numero_caja))
                  End If
                  If Len(Trim(Str(var_numero_caja))) = 3 Then
                     var_referencia_caja = Trim(Str(var_numero_caja))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 1 Then
                     var_referencia_embarque = "00000" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 2 Then
                     var_referencia_embarque = "0000" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 3 Then
                     var_referencia_embarque = "000" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 4 Then
                     var_referencia_embarque = "00" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 5 Then
                     var_referencia_embarque = "0" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 6 Then
                     var_referencia_embarque = Trim(Str(txt_embarque))
                  End If
                  var_numero_etiqueta = 1
                  var_mm = 0
                  'CAMBIO BIND
                  'var_cadena = "select B.NAME from oe_order_headers_all A, OE_TRANSACTION_TYPES_TL B where order_number = " + Me.txt_archivo + " AND A.ORDER_TYPE_ID = B.TRANSACTION_TYPE_ID AND LANGUAGE = 'ESA'"
                  If rsaux6.State = 1 Then
                     rsaux6.Close
                  End If
                  var_cadena = "select B.NAME from oe_order_headers_all A, OE_TRANSACTION_TYPES_TL B where order_number = ? AND A.ORDER_TYPE_ID = B.TRANSACTION_TYPE_ID AND LANGUAGE = 'ESA'"
                  strconsulta = var_cadena
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(CDbl(Me.txt_archivo)))
                       .Parameters.Append parametro
                  End With
                  Set rsaux6 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  
                  'rsaux6.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               
                  var_tipo_pedido = IIf(IsNull(rsaux6(0).Value), "", rsaux6(0).Value)
                  rsaux6.Close
                  If var_tipo_pedido <> "VIA_PEDIDO_INTERNO" Then ' para que ya no imprima el contenido
                     If var_tipo_pedido <> "VIA_MAYOREO_NACIONAL" Then
                     While Not rs.EOF
                           var_articulo = ""
                           If var_numero_etiqueta = 7 Then
                              var_numero_etiqueta = 1
                           End If
                           If var_numero_etiqueta = 1 Then
                              A.writeline ("")
                              A.writeline ("US")
                              A.writeline ("N")
                              A.writeline ("q816")
                              A.writeline ("Q1015,20+0")
                              A.writeline ("S2")
                              A.writeline ("D8")
                              A.writeline ("ZT")
                              A.writeline ("TTh:m")
                              A.writeline ("TDy2.mn.dd")
                           End If
'''' coppel
                           rsaux3.Open "SELECT description as vcha_Art_nombre_español FROM xxvia_system_items_b WHERE segment1 = '" + rs!SEGMENT1 + "' AND ORGANIZATION_ID = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           var_longitud = Len(Trim(rsaux3!vcha_art_nombre_español))
                           If var_longitud >= 35 Then
                              var_articulo = Replace(Left(Trim(rsaux3!vcha_art_nombre_español), 35), """", " ") + "  "
                           End If
                           If var_longitud < 35 Then
                              var_articulo = Replace(Trim(rsaux3!vcha_art_nombre_español), """", " ")
                              While Not var_longitud = 38
                                    var_articulo = var_articulo + " "
                                    var_longitud = var_longitud + 1
                              Wend
                           End If
                           rsaux3.Close
                           If Me.txt_entrega = "TIENDA HOUSTON" Or Me.txt_entrega = "TIENDA LOS ANGELES" Or Me.txt_entrega = "CEDIS KATY" Or Me.txt_entrega = "CEDIS CALIFORNIA" Then
                              
                              If cnn_icg_usa.State = 1 Then
                                 cnn_icg_usa.Close
                              End If
                              If Me.txt_entrega = "TIENDA HOUSTON" Or Me.txt_entrega = "CEDIS KATY" Then
                                 'cnn_icg_usa.Open "Provider=SQLOLEDB.1;Password=ICGUsa2014;Persist Security Info=True;User ID=ICGAdmin;Initial Catalog=bd1;Data Source=sqlcedishou.VIANNEYcatalog.COM"
                                 'rsaux3.Open "select REFPROVEEDOR, UBICACION from sqlcedishou.bd1.dbo.stocks st, sqlcedishou.bd1.dbo.ARTICULOS a where a.CODARTICULO = st.CODARTICULO and a.refproveedor = '" + rs!SEGMENT1 + "' AND ISNULL(UBICACION,'')<>''", cnn, adOpenDynamic, adLockOptimistic
                                 'If Not rsaux3.EOF Then
                                 '   var_ubicacion = "UBICACION: " + IIf(IsNull(rsaux3!ubicacion), "", rsaux3!ubicacion)
                                 'Else
                                 '   var_ubicacion = "UBICACION: "
                                 'End If
                                 'rsaux3.Close
                              
                              Else
                                 'cnn_icg_usa.Open "Provider=SQLOLEDB.1;Password=ICGUsa2014;Persist Security Info=True;User ID=ICGAdmin;Initial Catalog=bd1;Data Source=sqlcedisla.VIANNEYcatalog.COM"
                                 'rsaux3.Open "select REFPROVEEDOR, UBICACION from sqlcedisla.bd1.dbo.stocks st, sqlcedisla.bd1.dbo.ARTICULOS a where a.CODARTICULO = st.CODARTICULO and a.refproveedor = '" + rs!SEGMENT1 + "' AND ISNULL(UBICACION,'')<>''", cnn, adOpenDynamic, adLockOptimistic
                                 'If Not rsaux3.EOF Then
                                 '   var_ubicacion = "UBICACION: " + IIf(IsNull(rsaux3!ubicacion), "", rsaux3!ubicacion)
                                 'Else
                                 '   var_ubicacion = "UBICACION: "
                                 'End If
                                 'rsaux3.Close
                              End If
                              var_ubicacion = "UBICACION: "
                              
                              'cnn_icg_usa.Close
                              var_cantidad_caja_impresion = var_cantidad_caja_impresion + rs!FLOA_SAL_CANTIDAD_LEIDA
                              var_articulo = var_articulo + Trim(Str(rs!FLOA_SAL_CANTIDAD_LEIDA))
                              
                              If var_numero_etiqueta = 1 Then
                                 A.writeline ("A782,20,1,4,2,1,N,""" + var_articulo + """")
                                 A.writeline ("A696,20,1,4,2,1,N,""" + var_ubicacion + """")
                              End If
                              If var_numero_etiqueta = 3 Then
                                 A.writeline ("A627,20,1,4,2,1,N,""" + var_articulo + """")
                                 A.writeline ("A554,20,1,4,2,1,N,""" + var_ubicacion + """")
                              End If
                              If var_numero_etiqueta = 5 Then
                                 A.writeline ("A475,20,1,4,2,1,N,""" + var_articulo + """")
                                 A.writeline ("A390,20,1,4,2,1,N,""" + var_ubicacion + """")
                              End If
                              var_articulo = ""
                              rs.MoveNext
                              If rs.EOF Then
                                 var_numero_etiqueta = 5
                              End If
                              If var_numero_etiqueta = 5 Then
                                 A.writeline ("A270,20,1,5,1,1,N,""CAJA     :""")
                                 A.writeline ("A168,20,1,5,1,1,N,""EMBARQUE :""")
                                 A.writeline ("A116,20,1,4,2,1,N,""" + Mid(txt_cliente, 1, 47) + """")
                                 A.writeline ("A282,459,1,5,1,1,N,""" + var_referencia_caja + "/" + CStr(var_cantidad_caja_impresion) + "/" + var_cantidad_total + """")
                                 A.writeline ("A187,459,1,5,1,1,N,""" + var_referencia_embarque + """")
                                 If var_contador = 0 Then
                                    'cambio de caja en caso de ser exportaciones
                                    A.writeline ("B77,782,0,3,4,8,101,B,""C" + var_referencia_embarque + var_referencia_caja + """")
                                 End If
                                 var_contador = var_contador + 1
                                 A.writeline ("P1")
                              End If
                              var_numero_etiqueta = var_numero_etiqueta + 2
                           Else
                              var_cantidad_caja_impresion = var_cantidad_caja_impresion + rs!FLOA_SAL_CANTIDAD_LEIDA
                              var_articulo = var_articulo + Trim(Str(rs!FLOA_SAL_CANTIDAD_LEIDA))
                              If var_numero_etiqueta = 1 Then
                                 A.writeline ("A782,20,1,4,2,1,N,""" + var_articulo + """")
                              End If
                              If var_numero_etiqueta = 2 Then
                                 A.writeline ("A696,20,1,4,2,1,N,""" + var_articulo + """")
                              End If
                              If var_numero_etiqueta = 3 Then
                                 A.writeline ("A627,20,1,4,2,1,N,""" + var_articulo + """")
                              End If
                              If var_numero_etiqueta = 4 Then
                                 A.writeline ("A554,20,1,4,2,1,N,""" + var_articulo + """")
                              End If
                              If var_numero_etiqueta = 5 Then
                                 A.writeline ("A475,20,1,4,2,1,N,""" + var_articulo + """")
                              End If
                              If var_numero_etiqueta = 6 Then
                                 A.writeline ("A390,20,1,4,2,1,N,""" + var_articulo + """")
                              End If
                              var_articulo = ""
                              rs.MoveNext
                              If rs.EOF Then
                                 var_numero_etiqueta = 6
                              End If
                              If var_numero_etiqueta = 6 Then
                                 A.writeline ("A270,20,1,5,1,1,N,""CAJA     :""")
                                 A.writeline ("A168,20,1,5,1,1,N,""EMBARQUE :""")
                                 A.writeline ("A116,20,1,4,2,1,N,""" + Mid(txt_cliente, 1, 47) + """")
                                 A.writeline ("A282,459,1,5,1,1,N,""" + var_referencia_caja + "/" + CStr(var_cantidad_caja_impresion) + "/" + var_cantidad_total + """")
                                 A.writeline ("A187,459,1,5,1,1,N,""" + var_referencia_embarque + """")
                                 If var_contador = 0 Then
                                    'cambio de caja en caso de ser exportaciones
                                    A.writeline ("B77,782,0,3,4,8,101,B,""C" + var_referencia_embarque + var_referencia_caja + """")
                                 End If
                                 var_contador = var_contador + 1
                                 A.writeline ("P1")
                              End If
                              var_numero_etiqueta = var_numero_etiqueta + 1
                           End If
                     Wend
                     If var_numero_etiqueta < 6 Then
                        While Not var_numero_etiqueta = 7
                              If var_numero_etiqueta = 1 Then
                                 A.writeline ("A782,20,1,4,2,1,N,""" + var_articulo + """")
                              End If
                              If var_numero_etiqueta = 2 Then
                                 A.writeline ("A696,20,1,4,2,1,N,""" + var_articulo + """")
                              End If
                              If var_numero_etiqueta = 3 Then
                                 A.writeline ("A627,20,1,4,2,1,N,""" + var_articulo + """")
                              End If
                              If var_numero_etiqueta = 4 Then
                                 A.writeline ("A554,20,1,4,2,1,N,""" + var_articulo + """")
                              End If
                              If var_numero_etiqueta = 5 Then
                                 A.writeline ("A475,20,1,4,2,1,N,""" + var_articulo + """")
                              End If
                              If var_numero_etiqueta = 6 Then
                                 A.writeline ("A390,20,1,4,2,1,N,""" + var_articulo + """")
                              End If
                              var_articulo = ""
                              If var_numero_etiqueta = 6 Then
                                 A.writeline ("A270,20,1,5,1,1,N,""CAJA     :""")
                                 A.writeline ("A168,20,1,5,1,1,N,""EMBARQUE :""")
                                 A.writeline ("A116,20,1,4,2,1,N,""" + Mid(txt_cliente, 1, 47) + """")
                                 A.writeline ("A282,459,1,5,1,1,N,""" + var_referencia_caja + "/" + CStr(var_cantidad_caja_impresion) + "/" + var_cantidad_total + """")
                                 A.writeline ("A187,459,1,5,1,1,N,""" + var_referencia_embarque + """")
                                 If var_contador = 0 Then
                                    'cambio de caja en caso de ser exportaciones
                                    A.writeline ("B77,782,0,3,4,8,101,B,""C" + var_referencia_embarque + var_referencia_caja + """")
                                 End If
                                 var_contador = var_contador + 1
                                 A.writeline ("P1")
                              End If
                              If var_numero_etiqueta = 6 Then
                                 'a.writeline ("")
                                 'a.writeline ("O")
                                 'a.writeline ("q816<")
                                 'a.writeline ("Q1015,20+0")
                                 'a.writeline ("S2")
                                 'a.writeline ("D8")
                                 'a.writeline ("ZT")
                                 'a.writeline ("TTh: m")
                                 'a.writeline ("TDy2.mn.dd")
                              End If
                              var_numero_etiqueta = var_numero_etiqueta + 1
                        Wend
                     End If
                     End If
                  End If
                  
                  rsaux7.Open "SELECT HEADER_ID FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = " + Me.txt_archivo, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  
                  If Not rsaux7.EOF Then
                     VAR_HEADER_ID = IIf(IsNull(rsaux7!header_id), 0, rsaux7!header_id)
                  Else
                     VAR_HEADER_ID = 0
                  End If
                  rsaux7.Close
                  'var_cadena = "SELECT  a.source_header_type_name, oha.attribute8, oha.attribute9 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERY_DETAILS A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) IN (" + Me.txt_archivo + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID  AND ROWNUM = 1 and A.SOURCE_HEADER_ID = " + CStr(VAR_HEADER_ID) + " AND A.ORGANIZATION_ID = " + var_unidad_organizacional
                  If rsaux6.State = 1 Then
                     rsaux6.Close
                  End If
                  'lento
                  'cambio bind
                  var_cadena = "SELECT  a.source_header_type_name, oha.attribute8, oha.attribute9 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERY_DETAILS A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) = ? AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID  AND ROWNUM = 1 and A.SOURCE_HEADER_ID = ? AND A.ORGANIZATION_ID = ?"
                  strconsulta = var_cadena
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_archivo))
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(VAR_HEADER_ID))
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(var_unidad_organizacional))
                       .Parameters.Append parametro
                   End With
                   Set rsaux6 = comandoORA.execute
                   Set comandoORA = Nothing
                   Set parametro = Nothing
                  
                  'rsaux6.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  var_tipo_pedido = IIf(IsNull(rsaux6(0).Value), "", rsaux6(0).Value)
                  var_clave_almacen_tienda = IIf(IsNull(rsaux6!attribute8), "", rsaux6!attribute8)
                  var_clave_nombre_tienda = IIf(IsNull(rsaux6!ATTRIBUTE9), "", rsaux6!ATTRIBUTE9)
                  rsaux6.Close
                  var_nombre_cliente = Me.txt_cliente
                  If var_tipo_pedido = "VIA_PEDIDO_INTERNO" Or var_tipo_pedido = "TEX_PEDIDO_INTERNO" Then
                     If var_pedido_tienda = 0 Then
                        rsaux6.Open "select source_document_id from oe_order_headers_all where order_number in (" + Me.txt_archivo + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux7.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux6!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux7.EOF Then
                           var_almacen_tienda = rsaux7!attribute1
                           var_nombre_cliente = rsaux7!Description
                        Else
                           var_almacen_tienda = var_clave_almacen_tienda
                           var_nombre_cliente = var_clave_nombre_tienda
                        End If
                        rsaux7.Close
                        rsaux6.Close
                     Else
                        var_almacen_tienda = var_clave_almacen_tienda
                        var_nombre_cliente = var_clave_nombre_tienda
                     End If
                     If var_almacen_tienda <> "" Then
                        rsaux3.Open "select * from mtl_secondary_inventories where secondary_inventory_name = '" + var_almacen_tienda + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           var_location_id = IIf(IsNull(rsaux3!LOCATION_ID), 0, rsaux3!LOCATION_ID)
                           If var_location_id > 0 Then
                              rsaux4.Open "select ADDRESS_LINE_1, ADDRESS_LINE_2, TOWN_OR_CITY, REGION_1, COUNTRY, POSTAL_CODE  from hr_locations_all where location_id = '" + CStr(CDbl(var_location_id)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              VAR_DIRECCION = Mid(IIf(IsNull(rsaux4!ADDRESS_LINE_1), "", rsaux4!ADDRESS_LINE_1), 1, 50)
                              VAR_COLONIA = IIf(IsNull(rsaux4!ADDRESS_LINE_2), "", rsaux4!ADDRESS_LINE_2)
                              var_ciudad = IIf(IsNull(rsaux4!TOWN_OR_CITY), "", rsaux4!TOWN_OR_CITY)
                              var_estado = IIf(IsNull(rsaux4!REGION_1), "", rsaux4!REGION_1)
                              var_pais = IIf(IsNull(rsaux4!COUNTRY), "", rsaux4!COUNTRY)
                              VAR_CP = IIf(IsNull(rsaux4!POSTAL_CODE), "", rsaux4!POSTAL_CODE)
                              rsaux4.Close
                           Else
                              'lento
                              'cambio bind
                              var_cadena = "SELECT  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.invoice_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                              strconsulta = var_cadena
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(CDbl(Me.txt_archivo)))
                                   .Parameters.Append parametro
                              End With
                              Set rsaux6 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                           
                              'rsaux6.Open "SELECT  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.invoice_TO_ORG_ID AND oha.order_number = '" + CStr(CDbl(Me.txt_archivo)) + "' AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux6.EOF Then
                                 'lento
                                 'cambio bind
                                 var_cadena = "SELECT  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                                 strconsulta = var_cadena
                                 With comandoORA
                                      .ActiveConnection = cnnoracle_4
                                      .CommandType = adCmdText
                                      .CommandText = strconsulta
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(CDbl(Me.txt_archivo)))
                                      .Parameters.Append parametro
                                 End With
                                 Set rsaux5 = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing
                                 
                                 'rsaux5.Open "SELECT  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = '" + CStr(CDbl(Me.txt_archivo)) + "' AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux5.EOF Then
                                    VAR_DIRECCION = Mid(IIf(IsNull(rsaux5!calle), "", rsaux5!calle) + " " + IIf(IsNull(rsaux5!NUMERO), "", rsaux5!NUMERO), 1, 50)
                                    VAR_COLONIA = IIf(IsNull(rsaux5!colonia), "", rsaux5!colonia)
                                    var_ciudad = IIf(IsNull(rsaux5!ciudad), "", rsaux5!ciudad)
                                    VAR_MUNICIPIO = IIf(IsNull(rsaux5!municipio), "", rsaux5!municipio)
                                    var_estado = IIf(IsNull(rsaux5!estado), "", rsaux5!estado)
                                    var_pais = IIf(IsNull(rsaux5!pais), "", rsaux5!pais)
                                    VAR_CP = IIf(IsNull(rsaux5!cp), "", rsaux5!cp)
                                    rsaux5.Close
                                 Else
                                    rsaux5.Close
                                    VAR_DIRECCION = Mid(IIf(IsNull(rsaux6!calle), "", rsaux6!calle) + " " + IIf(IsNull(rsaux6!NUMERO), "", rsaux6!NUMERO), 1, 50)
                                    VAR_COLONIA = IIf(IsNull(rsaux6!colonia), "", rsaux6!colonia)
                                    var_ciudad = IIf(IsNull(rsaux6!ciudad), "", rsaux6!ciudad)
                                    VAR_MUNICIPIO = IIf(IsNull(rsaux6!municipio), "", rsaux6!municipio)
                                    var_estado = IIf(IsNull(rsaux6!estado), "", rsaux6!estado)
                                    var_pais = IIf(IsNull(rsaux6!pais), "", rsaux6!pais)
                                    VAR_CP = IIf(IsNull(rsaux6!cp), "", rsaux6!cp)
                                 End If
                              Else
                                 VAR_DIRECCION = ""
                                 VAR_COLONIA = ""
                                 var_ciudad = ""
                                 VAR_MUNICIPIO = ""
                                 var_estado = ""
                                 var_pais = ""
                                 VAR_CP = ""
                              End If
                              rsaux6.Close
                           End If
                        End If
                        rsaux3.Close
                     Else
                        'cambio bind
                        'rsaux6.Open "SELECT  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.invoice_TO_ORG_ID AND oha.order_number = '" + CStr(CDbl(Me.txt_archivo)) + "' AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        var_cadena = "SELECT  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.invoice_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                        strconsulta = var_cadena
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(CDbl(Me.txt_archivo)))
                             .Parameters.Append parametro
                        End With
                        Set rsaux6 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                     
                        If Not rsaux6.EOF Then
                           'cambio bind
                           'rsaux5.Open "SELECT  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = '" + CStr(CDbl(Me.txt_archivo)) + "' AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                            var_cadena = "SELECT  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                            strconsulta = var_cadena
                            With comandoORA
                                 .ActiveConnection = cnnoracle_4
                                 .CommandType = adCmdText
                                 .CommandText = strconsulta
                                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(CDbl(Me.txt_archivo)))
                                 .Parameters.Append parametro
                            End With
                            Set rsaux5 = comandoORA.execute
                            Set comandoORA = Nothing
                            Set parametro = Nothing
                           
                           If Not rsaux5.EOF Then
                              VAR_DIRECCION = Mid(IIf(IsNull(rsaux5!calle), "", rsaux5!calle) + " " + IIf(IsNull(rsaux5!NUMERO), "", rsaux5!NUMERO), 1, 50)
                              VAR_COLONIA = IIf(IsNull(rsaux5!colonia), "", rsaux5!colonia)
                              var_ciudad = IIf(IsNull(rsaux5!ciudad), "", rsaux5!ciudad)
                              VAR_MUNICIPIO = IIf(IsNull(rsaux5!municipio), "", rsaux5!municipio)
                              var_estado = IIf(IsNull(rsaux5!estado), "", rsaux5!estado)
                              var_pais = IIf(IsNull(rsaux5!pais), "", rsaux5!pais)
                              VAR_CP = IIf(IsNull(rsaux5!cp), "", rsaux5!cp)
                              rsaux5.Close
                           Else
                              rsaux5.Close
                              VAR_DIRECCION = Mid(IIf(IsNull(rsaux6!calle), "", rsaux6!calle) + " " + IIf(IsNull(rsaux6!NUMERO), "", rsaux6!NUMERO), 1, 50)
                              VAR_COLONIA = IIf(IsNull(rsaux6!colonia), "", rsaux6!colonia)
                              var_ciudad = IIf(IsNull(rsaux6!ciudad), "", rsaux6!ciudad)
                              VAR_MUNICIPIO = IIf(IsNull(rsaux6!municipio), "", rsaux6!municipio)
                              var_estado = IIf(IsNull(rsaux6!estado), "", rsaux6!estado)
                              var_pais = IIf(IsNull(rsaux6!pais), "", rsaux6!pais)
                              VAR_CP = IIf(IsNull(rsaux6!cp), "", rsaux6!cp)
                           End If
                        Else
                           VAR_DIRECCION = ""
                           VAR_COLONIA = ""
                           var_ciudad = ""
                           VAR_MUNICIPIO = ""
                           var_estado = ""
                           var_pais = ""
                           VAR_CP = ""
                        End If
                        rsaux6.Close
                     End If
                  Else
                     'cambio bind
                     'rsaux6.Open "SELECT  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.invoice_TO_ORG_ID AND oha.order_number = '" + CStr(CDbl(Me.txt_archivo)) + "' AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_cadena = "SELECT  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.invoice_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                     strconsulta = var_cadena
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(CDbl(Me.txt_archivo)))
                          .Parameters.Append parametro
                     End With
                     Set rsaux6 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     
                     If Not rsaux6.EOF Then
                        'cambio bind
                        'rsaux5.Open "SELECT  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = '" + CStr(CDbl(Me.txt_archivo)) + "' AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        var_cadena = "SELECT  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                        strconsulta = var_cadena
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(CDbl(Me.txt_archivo)))
                             .Parameters.Append parametro
                        End With
                        Set rsaux5 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        
                        If Not rsaux5.EOF Then
                           VAR_DIRECCION = Mid(IIf(IsNull(rsaux5!calle), "", rsaux5!calle) + " " + IIf(IsNull(rsaux5!NUMERO), "", rsaux5!NUMERO), 1, 50)
                           VAR_COLONIA = IIf(IsNull(rsaux5!colonia), "", rsaux5!colonia)
                           var_ciudad = IIf(IsNull(rsaux5!ciudad), "", rsaux5!ciudad)
                           VAR_MUNICIPIO = IIf(IsNull(rsaux5!municipio), "", rsaux5!municipio)
                           var_estado = IIf(IsNull(rsaux5!estado), "", rsaux5!estado)
                           var_pais = IIf(IsNull(rsaux5!pais), "", rsaux5!pais)
                           VAR_CP = IIf(IsNull(rsaux5!cp), "", rsaux5!cp)
                           rsaux5.Close
                        Else
                           rsaux5.Close
                           VAR_DIRECCION = Mid(IIf(IsNull(rsaux6!calle), "", rsaux6!calle) + " " + IIf(IsNull(rsaux6!NUMERO), "", rsaux6!NUMERO), 1, 50)
                           VAR_COLONIA = IIf(IsNull(rsaux6!colonia), "", rsaux6!colonia)
                           var_ciudad = IIf(IsNull(rsaux6!ciudad), "", rsaux6!ciudad)
                           VAR_MUNICIPIO = IIf(IsNull(rsaux6!municipio), "", rsaux6!municipio)
                           var_estado = IIf(IsNull(rsaux6!estado), "", rsaux6!estado)
                           var_pais = IIf(IsNull(rsaux6!pais), "", rsaux6!pais)
                           VAR_CP = IIf(IsNull(rsaux6!cp), "", rsaux6!cp)
                        End If
                     Else
                        VAR_DIRECCION = ""
                        VAR_COLONIA = ""
                        var_ciudad = ""
                        VAR_MUNICIPIO = ""
                        var_estado = ""
                        var_pais = ""
                        VAR_CP = ""
                     End If
                     rsaux6.Close
                  End If
                
                     
                  A.writeline ("")
                  A.writeline ("US")
                  A.writeline ("N")
                  A.writeline ("q816")
                  A.writeline ("Q1015,20+0")
                  A.writeline ("S2")
                  A.writeline ("D8")
                  A.writeline ("ZT")
                  A.writeline ("TTh:m")
                  A.writeline ("TDy2.mn.dd")
                  A.writeline ("A740,20,1,4,2,1,N,""Cliente: " + Mid(var_nombre_cliente, 1, 35) + """")
                  A.writeline ("A698,20,1,4,2,1,N,""Dirección: " + Mid(VAR_DIRECCION, 1, 35) + """")
                  A.writeline ("A656,20,1,4,2,1,N,""Colonia: " + VAR_COLONIA + """")
                  A.writeline ("A604,20,1,4,2,1,N,""C.P: " + VAR_CP + """")
                  A.writeline ("A552,20,1,4,2,1,N,""Ciudad: " + var_ciudad + """")
                  'A.writeline ("A562,20,1,4,2,1,N,""Municipio : " + VAR_MUNICIPIO + """")
                  A.writeline ("A500,20,1,4,2,1,N,""Estado: " + var_estado + ", " + var_pais + """")
                  If Len(Trim(Me.txt_caja_pedido)) = 1 Then
                     A.writeline ("A430,400,1,4,8,9,N,""" + Me.txt_caja_pedido + """")
                  Else
                     A.writeline ("A430,300,1,4,8,9,N,""" + Me.txt_caja_pedido + """")
                  End If
                  A.writeline ("A145,20,1,5,2,2,N,""   " + CStr(CDbl(Me.txt_archivo)) + """")
                  A.writeline ("B70,850,0,3,4,8,101,B,""C" + var_referencia_embarque + var_referencia_caja + """")
                  
                  
                  
                  rsaux8.Open "select shipping_method_code from oe_order_headers_all where order_number = " + CStr(CDbl(Me.txt_archivo)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  var_paqueteria = ""
                  If Not rsaux8.EOF Then
                     var_tipo_metodo = IIf(IsNull(rsaux8(0).Value), "", rsaux8(0).Value)
                     If var_tipo_metodo <> "" Then
                        rsaux9.Open "SELECT description FROM fnd_lookup_values where lookup_type = 'SHIP_METHOD' and lookup_code = '" + var_tipo_metodo + "' AND LANGUAGE = 'ESA'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux9.EOF Then
                           var_paqueteria = IIf(IsNull(rsaux9(0).Value), "", rsaux9(0).Value)
                        End If
                        rsaux9.Close
                     End If
                  End If
                  rsaux8.Close
                  If Len(var_paqueteria) > 20 Then
                     A.writeline ("A220,20,1,4,6,1,N,""" + var_paqueteria + """")
                  Else
                     A.writeline ("A220,20,1,4,8,3,N,""" + var_paqueteria + """")
                  End If
                  If var_clave_usuario_global <> "U0000000307" Then
                     'A.writeline ("A130,20,1,4,2,1,N,""CAJA     :""")
                     A.writeline ("A50,20,1,4,2,1,N,""EMBARQUE : " + var_referencia_embarque + " CAJA: " + var_referencia_caja + " PEDIDO: " + CStr(CDbl(Me.txt_archivo)) + """")
                     'A.writeline ("A130,300,1,4,2,1,N,""" + var_referencia_caja + """")
                     'A.writeline ("A50,300,1,4,2,1,N,""" + var_referencia_embarque + """" + " Caja: " + """ + var_referencia_caja + """ + var_referencia_caja + """")
                     'cambio de caja en caso de ser exportaciones
                     'A.writeline ("B77,782,0,3,4,8,101,B,""C" + var_referencia_embarque + var_referencia_caja + """")
                  End If
                  
                  A.writeline ("P1")
                  
                  'A.writeline ("")
                  'A.writeline ("US")
                  'A.writeline ("N")
                  'A.writeline ("q816")
                  'A.writeline ("Q1015,20+0")
                  'A.writeline ("S2")
                  'A.writeline ("D8")
                  'A.writeline ("ZT")
                  'A.writeline ("TTh:m")
                  'A.writeline ("TDy2.mn.dd")
                  'A.writeline ("A782,20,1,4,2,1,N,""Cliente: " + Mid(var_nombre_cliente, 1, 47) + """")
                  'cambio de caja en caso de ser exportaciones
                  
                  'If var_tipo_pedido = "VIA_EXPORTACION" Or var_tipo_pedido = "VIA_CATALOG_USA" Or var_tipo_pedido = "VIA_VTA_GUATEMALA" Then
                  '   If Len(Trim(Me.txt_caja)) = 1 Then
                  '      A.writeline ("A605,400,1,5,9,4,N,""" + Me.txt_caja + """")
                  '   Else
                  '      A.writeline ("A605,300,1,5,9,4,N,""" + Me.txt_caja + """")
                  '   End If
                  'Else
                  '   If Len(Trim(Me.txt_caja_pedido)) = 1 Then
                  '      A.writeline ("A605,400,1,5,9,4,N,""" + Me.txt_caja_pedido + """")
                  '   Else
                  '      A.writeline ("A605,300,1,5,9,4,N,""" + Me.txt_caja_pedido + """")
                  '   End If
                  'End If
                  'A.writeline ("A130,20,1,4,2,2,N,""PEDIDO: """)
                  'A.writeline ("A130,20,1,5,2,2,N,""   " + CStr(CDbl(Me.txt_archivo)) + """")
                  'A.writeline ("B77,782,0,3,4,8,101,B,""C" + var_referencia_embarque + var_referencia_caja + """")
                  'A.writeline ("P1")
                  If var_bari = 1 Then
                     A.writeline ("")
                     A.writeline ("US")
                     A.writeline ("N")
                     A.writeline ("q816")
                     A.writeline ("Q1015,20+0")
                     A.writeline ("S2")
                     A.writeline ("D8")
                     A.writeline ("ZT")
                     A.writeline ("TTh: m")
                     A.writeline ("TDy2.mn.dd")
                     A.writeline ("A605,80,1,5,9,4,N,""FRAGIL""")
                     A.writeline ("P1")
                  End If
               
                  A.Close
               
                  Open (App.Path & "\net_use.bat") For Output As #3
                  var_archivo = App.Path & "\net_use.bat"
                  Print #3, "net use lpt1 \\" + fun_NombrePc + "\zebra"
                  Close #3
                  x = Shell(var_archivo, vbHide)
               
               
                  Open (App.Path & "\etiquetas.bat") For Output As #2
                  var_archivo = App.Path & "\etiquetas.bat"
                  'Print #2, "copy " + App.Path + "\etiquetas.txt lpt1"
                  Print #2, "copy " + App.Path + "\etiquetas.txt \\" + fun_NombrePc + "\zebra"
                  
                  Close #2
                  x = Shell(var_archivo, vbHide)
                  Me.lbl_impresa.Visible = True
                  rsaux10.Open "select sum(floa_sal_Cantidad_leida) as cantidad from xxvia_tb_salidas_cajas where inte_emb_embarque = " + Me.txt_embarque + " and source_header_number = " + Me.txt_archivo + " and inte_paq_caja = " + Me.txt_caja, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux10.EOF Then
                     var_cantidad_caja_impresion = IIf(IsNull(rsaux10!cantidad), 0, rsaux10!cantidad)
                  Else
                     var_cantidad_caja_impresion = 0
                  End If
                  rsaux10.Close
                  rsaux10.Open "SELECT * FROM tb_oracle_cajas_aduana WHERE EMBARQUE = " + Me.txt_embarque + " AND NUMERO_CAJA = " + Me.txt_caja, cnn, adOpenDynamic, adLockOptimistic
                  If rsaux10.EOF Then
                     rsaux11.Open "insert into tb_oracle_cajas_aduana (embarque, pedido, numero_caja, caja, agente, cliente, establecimiento, piezas, estatus, TIPO_EMPAQUE, caja_pedido) values (" + Me.txt_embarque + "," + Me.txt_archivo + "," + Me.txt_caja + ",'C" + var_referencia_embarque + var_referencia_caja + "','" + Me.txt_agente + "','" + var_nombre_cliente + "',''," + CStr(var_cantidad_caja_impresion) + ",'','" + Me.txt_nombre_caja + "'," + Me.txt_caja_pedido + ")", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     If rsaux11.State = 1 Then
                        rsaux11.Close
                     End If
                     rsaux11.Open "UPDATE TB_ORACLE_CAJAS_ADUANA SET PIEZAS = " + CStr(var_cantidad_caja_impresion) + "  WHERE EMBARQUE = " + Me.txt_embarque + " AND NUMERO_CAJA = " + Me.txt_caja, cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux10.Close
                  'End If 'de var_bandera_asignacion = 0
                  
                  rsaux10.Open "select * from tb_video", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux10.EOF Then
                     V = IIf(IsNull(rsaux10(0).Value), 0, rsaux10(0).Value)
                  Else
                     V = 0
                  End If
                  rsaux10.Close
                  If V = 1 Then
                     If var_modo_texto_ip = 1 Then
                        On Error GoTo salir2
                        var_cadena = "@B@ " + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "EMBARQUE:     " + Me.txt_embarque + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "PEDIDO:       " + Me.txt_archivo + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "CAJA:         " + Me.txt_caja + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "CAJA PEDIDO:  " + Me.txt_caja_pedido + Chr(13) + Chr(10)
                        'rsaux10.Open "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS where source_header_number = " + CStr(CDbl(Me.txt_archivo)) + " and inte_paq_caja = " + Me.txt_caja + " AND INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        strconsulta = "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS where source_header_number = ? and inte_paq_caja = ? AND INTE_EMB_EMBARQUE = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_archivo))
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_caja))
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
                             .Parameters.Append parametro
                        End With
                        Set rsaux10 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                     
                        var_cadena = var_cadena + "TIPO DE CAJA: " + IIf(IsNull(rsaux10!tipo_caja), "", rsaux10!tipo_caja) + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "SELLO:        " + IIf(IsNull(rsaux10!sello), "", rsaux10!sello) + Chr(13) + Chr(10)
                        rsaux9.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_nombre_usuario = IIf(IsNull(rsaux9!vcha_usu_nombre), "", rsaux9!vcha_usu_nombre) + " " + IIf(IsNull(rsaux9!vcha_usu_apellidos), "", rsaux9!vcha_usu_apellidos)
                        rsaux9.Close
                        var_cadena = var_cadena + "USUARIO:      " + var_nombre_usuario + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "FECHA Y HORA: " + CStr(Now) + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "" + Chr(13) + Chr(10)
                               
                        var_cadena = var_cadena + "======================================================================" + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "CODIGO     DESCRIPCION                                      Cantidad  " + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "______________________________________________________________________" + Chr(13) + Chr(10)
                        var_cantidad_total = 0
                        While Not rsaux10.EOF
                              var_nombre_articulo = Mid(IIf(IsNull(rsaux10!item_description), "", rsaux10!item_description), 1, 57)
                              For var_j = Len(var_nombre_articulo) To 50
                                  var_nombre_articulo = var_nombre_articulo + " "
                              Next var_j
                              var_cantidad_total = var_cantidad_total + IIf(IsNull(rsaux10!FLOA_SAL_CANTIDAD_LEIDA), 0, rsaux10!FLOA_SAL_CANTIDAD_LEIDA)
                              VAR_CANTIDAD_TOTAL_STR = Format(CStr(var_cantidad_total), "###,###,##0.00")
                              For var_j = Len(VAR_CANTIDAD_TOTAL_STR) To 14
                                  VAR_CANTIDAD_TOTAL_STR = " " + VAR_CANTIDAD_TOTAL_STR
                              Next var_j
                         
                              VAR_CANTIDAD_ETIQUETA = Format(" " + CStr(IIf(IsNull(rsaux10!FLOA_SAL_CANTIDAD_LEIDA), 0, rsaux10!FLOA_SAL_CANTIDAD_LEIDA)), "###,###,##0.00")
                              For var_j = Len(VAR_CANTIDAD_ETIQUETA) To 5
                                  VAR_CANTIDAD_ETIQUETA = " " + VAR_CANTIDAD_ETIQUETA
                              Next var_j
                              var_cadena = var_cadena + IIf(IsNull(rsaux10!SEGMENT1), "", rsaux10!SEGMENT1) + "   " + var_nombre_articulo + VAR_CANTIDAD_ETIQUETA + Chr(13) + Chr(10)
                              rsaux10.MoveNext
                        Wend
                        var_cadena = var_cadena + "______________________________________________________________________" + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "                                               TOTAL:" + VAR_CANTIDAD_TOTAL_STR + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "======================================================================" + Chr(13) + Chr(10)
'--   --
                        rsaux10.MoveFirst
                        var_cadena = var_cadena + "EMBARQUE:     " + Me.txt_embarque + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "PEDIDO:       " + Me.txt_archivo + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "CAJA:         " + Me.txt_caja + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "CAJA PEDIDO:  " + Me.txt_caja_pedido + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "TIPO DE CAJA: " + IIf(IsNull(rsaux10!tipo_caja), "", rsaux10!tipo_caja) + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "SELLO:        " + IIf(IsNull(rsaux10!sello), "", rsaux10!sello) + Chr(13) + Chr(10)
                        rsaux9.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_nombre_usuario = IIf(IsNull(rsaux9!vcha_usu_nombre), "", rsaux9!vcha_usu_nombre) + " " + IIf(IsNull(rsaux9!vcha_usu_apellidos), "", rsaux9!vcha_usu_apellidos)
                       rsaux9.Close
                        var_cadena = var_cadena + "USUARIO:      " + var_nombre_usuario + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "FECHA Y HORA: " + CStr(Now) + Chr(13) + Chr(10)
                        var_cadena = var_cadena + "" + Chr(13) + Chr(10)
                        rsaux10.Close
'--   --
                        var_cadena = var_cadena + "@E@ "
                        On Error GoTo SALIR:
                        Set clnt = Nothing
                        clnt.MSSoapInit var_webservice_texto
                        var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), var_cadena + Chr(13))
                        Set clnt = Nothing
                     Else
                        On Error GoTo salir2
                        'If 'mscomm1.PortOpen = True Then
                        '   'mscomm1.PortOpen = False
                        'End If
                        'mscomm1.CommPort = 1
                        'mscomm1.settings = var_baudios
                        'mscomm1.PortOpen = True
                        'mscomm1.Output = "@B@ " + Chr(13) + Chr(10)
                        'mscomm1.Output = "EMBARQUE:     " + Me.txt_embarque + "^]EOL"
                        'mscomm1.Output = "PEDIDO:       " + Me.txt_archivo + "^]EOL"
                        'mscomm1.Output = "CAJA:         " + Me.txt_caja + "^]EOL"
                        'mscomm1.Output = "CAJA PEDIDO:  " + Me.txt_caja_pedido + "^]EOL"
                        'rsaux10.Open "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS where source_header_number = " + CStr(CDbl(Me.txt_archivo)) + " and inte_paq_caja = " + Me.txt_caja + " AND INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        strconsulta = "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS where source_header_number = ? and inte_paq_caja = ? AND INTE_EMB_EMBARQUE = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_archivo))
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_caja))
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
                             .Parameters.Append parametro
                        End With
                        Set rsaux10 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                     
                        'mscomm1.Output = "TIPO DE CAJA: " + IIf(IsNull(rsaux10!tipo_caja), "", rsaux10!tipo_caja) + "^]EOL"
                        'mscomm1.Output = "SELLO:        " + IIf(IsNull(rsaux10!sello), "", rsaux10!sello) + "^]EOL"
                        rsaux9.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_nombre_usuario = IIf(IsNull(rsaux9!vcha_usu_nombre), "", rsaux9!vcha_usu_nombre) + " " + IIf(IsNull(rsaux9!vcha_usu_apellidos), "", rsaux9!vcha_usu_apellidos)
                        rsaux9.Close
                        'mscomm1.Output = "USUARIO:      " + var_nombre_usuario + "^]EOL"
                        'mscomm1.Output = "FECHA Y HORA: " + CStr(Now) + "^]EOL"
                        'mscomm1.Output = "" + "^]EOL"
                               
                        'mscomm1.Output = "======================================================================" + "^]EOL"
                        'mscomm1.Output = "CODIGO     DESCRIPCION                                      Cantidad  " + "^]EOL"
                        'mscomm1.Output = "______________________________________________________________________" + "^]EOL"
                        var_cantidad_total = 0
                        While Not rsaux10.EOF
                              var_nombre_articulo = Mid(IIf(IsNull(rsaux10!item_description), "", rsaux10!item_description), 1, 57)
                              For var_j = Len(var_nombre_articulo) To 50
                                  var_nombre_articulo = var_nombre_articulo + " "
                              Next var_j
                              var_cantidad_total = var_cantidad_total + IIf(IsNull(rsaux10!FLOA_SAL_CANTIDAD_LEIDA), 0, rsaux10!FLOA_SAL_CANTIDAD_LEIDA)
                              VAR_CANTIDAD_TOTAL_STR = Format(CStr(var_cantidad_total), "###,###,##0.00")
                              For var_j = Len(VAR_CANTIDAD_TOTAL_STR) To 14
                                  VAR_CANTIDAD_TOTAL_STR = " " + VAR_CANTIDAD_TOTAL_STR
                              Next var_j
                      
                              VAR_CANTIDAD_ETIQUETA = Format(" " + CStr(IIf(IsNull(rsaux10!FLOA_SAL_CANTIDAD_LEIDA), 0, rsaux10!FLOA_SAL_CANTIDAD_LEIDA)), "###,###,##0.00")
                              For var_j = Len(VAR_CANTIDAD_ETIQUETA) To 5
                                  VAR_CANTIDAD_ETIQUETA = " " + VAR_CANTIDAD_ETIQUETA
                              Next var_j
                              'mscomm1.Output = IIf(IsNull(rsaux10!SEGMENT1), "", rsaux10!SEGMENT1) + "   " + var_nombre_articulo + VAR_CANTIDAD_ETIQUETA + "^]EOL"
                              rsaux10.MoveNext
                        Wend
                        rsaux10.Close
                        'mscomm1.Output = "______________________________________________________________________" + "^]EOL"
                        'mscomm1.Output = "                                               TOTAL:" + VAR_CANTIDAD_TOTAL_STR + "^]EOL"
                        'mscomm1.Output = "======================================================================" + "^]EOL"
                              
                        'mscomm1.Output = "@E@ "
                        'mscomm1.OutBufferCount = 0
                        'mscomm1.PortOpen = False
   
                     End If
                  
                  End If
                  Else
                     MsgBox "Sello incorrecto", vbOKOnly, "ATENCION"
                  End If
               End If
            End If
            rs.Close
            
         End If
      Else
         MsgBox "No esta autorizado para reimprimir etiquetas.", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
   End If
   Exit Sub
SALIR:
   If Err.Number = 70 Then
      MsgBox "Faltan permisos para imprimir", vbOKOnly, "ATENCION"
   Else
      MsgBox "La etiqueta no se pudo imprimir", vbOKOnly, "ATENCION"
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
   Exit Sub
salir2:
   'MsgBox Err.Description
   
   Resume Next
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
   
End Sub

Private Sub cmd_imprimir_reporte_faltantes_Click()
   If Me.lv_salidas.ListItems.Count > 0 Then
      rs.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_usuario_reporte = IIf(IsNull(rs!vcha_usu_nombre), "", rs!vcha_usu_nombre) + " " + IIf(IsNull(rs!vcha_usu_apellidos), "", rs!vcha_usu_apellidos)
      End If
      rs.Close
      For var_j = 1 To lv_salidas.ListItems.Count
          If var_j = 1 Then
             cnn.BeginTrans
             rs.Open "select max(inte_tem_consecutivo) from tb_temp_oracle_pedido_piezas_faltantes", cnn, adOpenDynamic, adLockOptimistic
             If Not rs.EOF Then
                var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
             Else
                var_consecutivo = 1
             End If
             rs.Close
             rs.Open "insert into tb_temp_oracle_pedido_piezas_faltantes (inte_tem_Consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
             cnn.CommitTrans
          End If
          lv_salidas.ListItems.Item(var_j).Selected = True
          'rsaux.Open "SELECT ATTRIBUTE2 FROM xxvia_system_items_b WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND SEGMENT1 = '" + Me.lv_salidas.selectedItem + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
          strconsulta = "SELECT ATTRIBUTE2 FROM xxvia_system_items_b WHERE ORGANIZATION_ID = ? AND SEGMENT1 = ?"
          With comandoORA
               .ActiveConnection = cnnoracle_4
               .CommandType = adCmdText
               .CommandText = strconsulta
               Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
               .Parameters.Append parametro
               Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.lv_salidas.selectedItem)
               .Parameters.Append parametro
          End With
          Set rsaux = comandoORA.execute
          Set comandoORA = Nothing
          Set parametro = Nothing
          
          strconsulta = "select * from Xxvia_vw_existencias_inv where organization_id = ? and subinventory_code = ? and segment1 = ?"
          With comandoORA
               .ActiveConnection = cnnoracle_4
               .CommandType = adCmdText
               .CommandText = strconsulta
               Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_unidad_organizacional))
               .Parameters.Append parametro
               Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, "CDI_ALMPT")
               .Parameters.Append parametro
               Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.lv_salidas.selectedItem)
              .Parameters.Append parametro
          End With
          Set rsaux9 = comandoORA.execute
          Set comandoORA = Nothing
          Set parametro = Nothing
          VAR_EXISTEN = 0
          var_disponible = 0
          If Not rsaux9.EOF Then
             VAR_EXISTEN = IIf(IsNull(rsaux9!CANTMANO), 0, rsaux9!CANTMANO)
             var_disponible = IIf(IsNull(rsaux9!Disponible), 0, rsaux9!Disponible)
          End If
          
          
          
          var_ubicacion = ""
          If Not rsaux.EOF Then
             var_ubicacion = IIf(IsNull(rsaux!attribute2), "", rsaux!attribute2)
          End If
          rsaux.Close
          rs.Open "insert into tb_temp_oracle_pedido_piezas_faltantes (inte_tem_consecutivo, embarque, pedido, agente, cliente, codigo, descripcion, cantidad_pedida, cantidad_surtida, cantidad_faltante, usuario, maquina, UBICACION) values (" + CStr(var_consecutivo) + "," + Me.txt_embarque + "," + Me.txt_archivo + ",'" + Me.txt_agente + "', '" + Me.txt_cliente + "','" + Me.lv_salidas.selectedItem + "','" + Mid(Me.lv_salidas.selectedItem.SubItems(1), 1, 50) + "'," + Me.lv_salidas.selectedItem.SubItems(2) + "," + Me.lv_salidas.selectedItem.SubItems(3) + "," + Me.lv_salidas.selectedItem.SubItems(5) + ",'" + var_usuario_reporte + "','" + fun_NombrePc + "','" + var_ubicacion + "')", cnn, adOpenDynamic, adLockOptimistic
      Next var_j
      rsaux.Open "delete from tb_temp_oracle_pedido_piezas_faltantes where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and codigo is null", cnn, adOpenDynamic, adLockOptimistic
      rsaux.Open "select * from tb_temp_oracle_pedido_piezas_faltantes where inte_tem_consecutivo = " + CStr(var_consecutivo) + "and cantidad_faltante > 0", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         Set reporte = appl.OpenReport(App.Path + "\rep_oracle_pedido_piezas_faltantes.rpt")
         reporte.RecordSelectionFormula = "{VW_ORACLE_PEDIDO_PIEZAS_FALTANTES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_PEDIDO_PIEZAS_FALTANTES.CANTIDAD_FALTANTE} > 0"
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Piezas faltantes"
         frmvistasprevias.Show 1
         Set reporte = Nothing
      Else
         MsgBox "No hay faltantes en el pedido", vbOKOnly, "ATENCION"
      End If
      rsaux.Close
   Else
      MsgBox "No se a seleccionado un pedido", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_mensaje_2_Click()
   Me.wmp2.Controls.play
End Sub

Private Sub cmd_mensaje_4_Click()
   Me.wmp4.Controls.play
End Sub

Private Sub cmd_nuevo_Click()
   frmoracle_tipo_cajas.Show 1
   Me.txt_nombre_caja = var_nombre_caja
   rsaux7.Open "select * from tb_oracle_empaques where empaque = '" + Me.txt_nombre_caja + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux7.EOF Then
      Me.lbl_maximo = Format(IIf(IsNull(rsaux7!PESO), 0, rsaux7!PESO), "###,###,##0.000")
   Else
      Me.lbl_maximo = "0.000"
   End If
   rsaux7.Close
   Me.lbl_peso = "0.000"
   Me.txt_caja = ""
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
   VAR_ESTATUS = IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus)
   rs.Close
   If Trim(VAR_ESTATUS) = "" Then
      For var_j = 1 To Me.lv_salidas.ListItems.Count
          Me.lv_salidas.ListItems(var_j).Selected = True
          Me.lv_salidas.selectedItem.SubItems(4) = 0
      Next var_j
      Me.txt_archivo.Enabled = False
      Me.lbl_cantidad_caja = 0
      Me.txt_caja_pedido = ""
      var_primera_vez = 1
      Me.txt_codigo.Enabled = True
      Me.txt_codigo.SetFocus
      Me.lbl_impresa.Visible = False
   Else
      Me.txt_archivo.Enabled = False
      Me.txt_codigo.Enabled = False
      MsgBox "El embarque ya no puede ser modificado", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   
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
      
   End If
   If Shift = 4 And KeyCode = 77 Then
   End If
End Sub

Private Sub Form_Load()
   On Error GoTo NO_PUERTO:
   rs.Open "select * from tb_oracle_maquinas where maquina = '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_puerto = IIf(IsNull(rs!COM_BASCULA), 0, rs!COM_BASCULA)
      If var_puerto > 0 Then
         puerto.CommPort = var_puerto
         puerto.PortOpen = True
         Me.Timer1.Enabled = True
      End If
   Else
      Me.Timer1.Enabled = False
   End If
   rs.Close
NO_PUERTO:
   Me.lbl_impresa.Visible = False
   Top = 0
   Left = 0
   frm_eliminar.Visible = False
   Me.txt_embarque = var_numero_embarque
   Me.txt_jaula = var_numero_jaula
   If IsNumeric(Me.txt_archivo) Then
      Call ejecuta
   End If
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   Me.lbl_cantidad.Visible = False
   Me.txt_cantidad.Visible = False
   cmd_pasar_movimiento.Visible = False
   Me.frm_busqueda.Visible = False
   rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
   VAR_ESTATUS = IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus)
   rs.Close
   If Trim(VAR_ESTATUS) <> "" Then
      Me.txt_archivo.Enabled = False
      Me.txt_codigo.Enabled = False
      MsgBox "El embarque ya no puede ser modificado", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub lv_salidas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_salidas, ColumnHeader)
End Sub

Private Sub lv_salidas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      Me.txt_cantidad_eliminar = ""
      Me.frm_eliminar.Visible = True
      Me.txt_cantidad_eliminar.SetFocus
   End If
End Sub

Private Sub Timer1_Timer()
On Error GoTo SALIR:
   x = 0
   If x = 0 Then
   textin = ""
   textin = puerto.Input
   If textin <> "" Then
      var_j = Len(textin)
      texto = ""
      While var_j > 0
            If IsNumeric(Mid(textin, var_j, 1)) Or Mid(textin, var_j, 1) = "." Then
              texto = texto + Mid(textin, var_j, 1)
            End If
            
            var_j = var_j - 1
      Wend
      If IsNumeric(texto) Then
         texto = CDbl(texto)
      End If
      Me.lbl_bascula = texto
   End If
   Else
      puerto.Output = 12313
   End If
   Exit Sub
SALIR:
   Me.lbl_bascula = "ERROR"
End Sub

Private Sub txt_archivo_KeyDown(KeyCode As Integer, Shift As Integer)
   If var_bandera_asignacion = 0 Then
      If KeyCode = 116 Then
         var_embarque_global = CDbl(Me.txt_embarque)
         frmoracle_seleccion_pedido.Show 1
         Me.txt_archivo = var_pedido_global
      End If
   End If
End Sub

Private Sub txt_archivo_KeyPress(KeyAscii As Integer)
   If var_bandera_asignacion = 0 Then
      If KeyAscii = 13 Then
         If IsNumeric(Me.txt_archivo) Then
            Call ejecuta
            If IsNumeric(Me.txt_archivo) Then
               strconsulta = "select nvl(estatus_pedido,0) from xxvia_tb_salidas_cajas where source_header_number = ? "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.txt_archivo))
                    .Parameters.Append parametro
               End With
               Set rsaux8 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               If Not rsaux8.EOF Then
                  If rsaux8(0).Value = 0 Then
                  Else
                     Me.txt_codigo.Enabled = False
                     MsgBox "El pedido ya fue cerrado", vbOKOnly, "ATENCION"
                  End If
               End If
            End If
         Else
            MsgBox "Orden de surtido incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         KeyAscii = 0
      End If
   Else
      If KeyAscii = 13 Then
         If IsNumeric(Me.txt_archivo) Then
            Call ejecuta
            If IsNumeric(Me.txt_archivo) Then
               strconsulta = "select nvl(estatus_pedido,0) from xxvia_tb_salidas_cajas where source_header_number = ? "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.txt_archivo))
                    .Parameters.Append parametro
               End With
               Set rsaux8 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               If Not rsaux8.EOF Then
                  If rsaux8(0).Value = 0 Then
                  Else
                     Me.txt_codigo.Enabled = False
                     MsgBox "El pedido ya fue cerrado", vbOKOnly, "ATENCION"
                  End If
               End If
            End If
         Else
            MsgBox "Orden de surtido incorrecta", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub txt_busqueda_caja_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_busqueda_caja) Then
         
         'rsaux8.Open "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND INTE_PAQ_CAJA = " + Me.txt_busqueda_caja, cnnoracle_4, adOpenDynamic, adLockOptimistic
         strconsulta = "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = ? AND INTE_PAQ_CAJA = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_busqueda_caja))
              .Parameters.Append parametro
         End With
         Set rsaux8 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         
         
         If Not rsaux8.EOF Then
'-----
            var_posible_maquina = 0
            If var_bandera_asignacion = 0 Then
               rsaux9.Open "select * from TB_ORACLE_PEDIDOS_MAQUINAS where pedido = " + CStr(rsaux8!source_header_number), cnn, adOpenDynamic, adLockOptimistic
               If rsaux9.EOF Then
                  var_posible_maquina = 1
               Else
                  If IIf(IsNull(rsaux9!maquina), "", rsaux9!maquina) = fun_NombrePc Then
                     var_posible_maquina = 1
                  Else
                     var_posible_maquina = 0
                  End If
               End If
               rsaux9.Close
            Else
               var_posible_maquina = 1
            End If
            If var_posible_maquina = 1 Then
               Me.txt_caja_pedido = IIf(IsNull(rsaux8!caja_pedido), 0, rsaux8!caja_pedido)
               var_nombre_caja = IIf(IsNull(rsaux8!tipo_caja), "", rsaux8!tipo_caja)
               Me.txt_nombre_caja = var_nombre_caja
               rsaux7.Open "select * from tb_oracle_empaques where empaque = '" + Me.txt_nombre_caja + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux7.EOF Then
                  Me.lbl_maximo = Format(IIf(IsNull(rsaux7!PESO), 0, rsaux7!PESO), "###,###,##0.000")
               Else
                  Me.lbl_maximo = "0.000"
               End If
               rsaux7.Close
               
               rsaux5.Open "SELECT SUM(FLOA_SAL_CANTIDAD_LEIDA * NVL(PESO,0)) AS PESO FROM XXVIA_tB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND INTE_PAQ_CAJA = " + Me.txt_busqueda_caja, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux5.EOF Then
                  Me.lbl_peso = Format(IIf(IsNull(rsaux5(0).Value), 0, rsaux5(0).Value), "###,###,##0.000")
               Else
                  Me.lbl_peso = "0.000"
               End If
               rsaux5.Close
            
            
            
               VAR_ESTATUS = IIf(IsNull(rsaux8!char_paq_estatus), "", rsaux8!char_paq_estatus)
               If VAR_ESTATUS <> "" Then
                  Me.txt_codigo.Enabled = False
                  Me.lbl_impresa.Visible = True
               Else
                  Me.txt_codigo.Enabled = True
                  Me.lbl_impresa.Visible = False
               End If
               var_orden = rsaux8!source_header_number
               Me.txt_archivo = var_orden
               Me.txt_caja = Me.txt_busqueda_caja
               'AQUI EMPIEZA EL CAMBIO variables bind
               'var_cadena = "SELECT HCAS.CUST_ACCOUNT_ID, HCAS.CUST_ACCT_SITE_ID as customer_id, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, oha.attribute8, oha.attribute9, A.source_header_type_name  from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND to_number(source_header_number) BETWEEN " + CStr(var_orden) + " AND " + CStr(var_orden)
               'var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID  AND released_status = 'Y'"
               'rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_cadena = "SELECT HCAS.CUST_ACCOUNT_ID, HCAS.CUST_ACCT_SITE_ID as customer_id, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, oha.attribute8, oha.attribute9, A.source_header_type_name  from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND to_number(source_header_number) = ?"
               var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID  AND released_status = 'Y'"
               
               strconsulta = var_cadena
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(CStr(var_orden)))
                    .Parameters.Append parametro
                End With
                Set rs = comandoORA.execute
                Set comandoORA = Nothing
                Set parametro = Nothing
               
               If Not rs.EOF Then
                  If rsaux.State = 1 Then
                     rsaux.Close
                  End If
                  rsaux.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  VAR_AGENTE_str = IIf(IsNull(rsaux!collector_id), "", rsaux!collector_id)
                  var_nombre_agente_str = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                  rsaux.Close
                  var_primera_vez = 0
                  Me.txt_agente = var_nombre_agente_str
                  Me.txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                  If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rs!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                     If var_pedido_tienda = 0 Then
                        Me.txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                     Else
                        Me.txt_cliente = IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9)
                     End If
                  End If
                  Me.txt_origen = IIf(IsNull(rs!subinventory), "", rs!subinventory)
                  Me.lv_salidas.ListItems.Clear
                  var_cantidad_enviada = 0
                  While Not rs.EOF
                        Set list_item = lv_salidas.ListItems.Add(, , rs!SEGMENT1)
                        list_item.SubItems(1) = IIf(IsNull(rs!Description), "", rs!Description)
                        list_item.SubItems(2) = Format(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity), "###,###,##0.00")
                        var_cantidad_enviada = var_cantidad_enviada + IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)
                        list_item.SubItems(3) = 0
                        list_item.SubItems(4) = 0
                        list_item.SubItems(5) = Format(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity), "###,###,##0.00")
                        list_item.SubItems(6) = IIf(IsNull(rs!inventory_item_id), 0, rs!inventory_item_id)
                        list_item.SubItems(7) = IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)
                        list_item.SubItems(8) = IIf(IsNull(rs!SOURCE_LINE_NUMBER), 0, rs!SOURCE_LINE_NUMBER)
                        list_item.SubItems(9) = IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)
                        list_item.SubItems(10) = IIf(IsNull(rs!CUSTOMER_ID), 0, rs!CUSTOMER_ID)
                        rs.MoveNext
                  Wend
                  Me.lbl_enviados = Format(var_cantidad_enviada, "###,###,##0.00")
                  Me.lbl_recibidos = Format(0, "###,###,##0.00")
                  Me.lbl_cantidad_caja = Format(0, "###,###,##0.00")
                  Me.txt_archivo.Enabled = False
                  var_cantidad_recibida = 0
                  'rsaux2.Open "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND  source_header_number = " + CStr(CDbl(var_orden)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  strconsulta = "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = ? AND  source_header_number = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_orden))
                       .Parameters.Append parametro
                  End With
                  Set rsaux2 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  
                  If Not rsaux2.EOF Then
                     Me.txt_entrega = Replace(IIf(IsNull(rsaux2!ENTREGA), "", rsaux2!ENTREGA), "'", " ")
                     While Not rsaux2.EOF
                           var_codigo = rsaux2!SEGMENT1
                           For var_j = 1 To Me.lv_salidas.ListItems.Count
                               Me.lv_salidas.ListItems.Item(var_j).Selected = True
                               If Me.lv_salidas.selectedItem.SubItems(7) = rsaux2!delivery_detail_id Then
                                  Me.lv_salidas.selectedItem.SubItems(3) = CDbl(Me.lv_salidas.selectedItem.SubItems(3)) + Format(rsaux2!FLOA_SAL_CANTIDAD_LEIDA, "###,###,##0.00")
                                  Me.lv_salidas.selectedItem.SubItems(5) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(2)) - CDbl(Me.lv_salidas.selectedItem.SubItems(3)), "###,###,##0.00")
                               End If
                           Next var_j
                           var_cantidad_recibida = var_cantidad_recibida + rsaux2!FLOA_SAL_CANTIDAD_LEIDA
                           rsaux2.MoveNext
                     Wend
                  End If
                  rsaux2.Close
                  rsaux2.Open "select * from tb_oracle_pedidos_asignados_embarques where pedido = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     Me.txt_orden_lectura = IIf(IsNull(rsaux2!orden_pedido), "", rsaux2!orden_pedido)
                  Else
                     Me.txt_orden_lectura = ""
                  End If
                  rsaux2.Close
                  
                  Me.lbl_recibidos = Format(var_cantidad_recibida, "###,###,##0.00")
               End If
            
'------
               For var_j = 1 To lv_salidas.ListItems.Count
                   lv_salidas.ListItems.Item(var_j).Selected = True
                   lv_salidas.selectedItem.SubItems(4) = "0.00"
               Next var_j
               Me.lbl_cantidad_caja = "0.00"
               var_cantidad_recibida = "0.00"
               While Not rsaux8.EOF
                     var_codigo = rsaux8!SEGMENT1
                     For var_j = 1 To Me.lv_salidas.ListItems.Count
                         Me.lv_salidas.ListItems.Item(var_j).Selected = True
                         If CDbl(Me.lv_salidas.selectedItem.SubItems(7)) = CDbl(rsaux8!delivery_detail_id) Then
                            Me.lv_salidas.selectedItem.SubItems(4) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(4)) + rsaux8!FLOA_SAL_CANTIDAD_LEIDA, "###,###,##0.00")
                         End If
                     Next var_j
                     var_cantidad_recibida = var_cantidad_recibida + rsaux8!FLOA_SAL_CANTIDAD_LEIDA
                     rsaux8.MoveNext
               Wend
               Me.lbl_cantidad_caja = Format(var_cantidad_recibida, "###,###,##0.00")
               If Me.txt_codigo.Enabled = True Then
                  Me.txt_codigo.SetFocus
               End If
               rs.Close
               Me.frm_busqueda.Visible = False
            Else
               Me.txt_codigo.Enabled = False
               MsgBox "La caja fue hecha en otra máquina", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "La caja no existe", vbOKOnly, "ATENCION"
         End If
         rsaux8.Close
      Else
         MsgBox "Número de caja incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_busqueda_embarque_KeyPress(KeyAscii As Integer)
End Sub

Private Sub txt_busqueda_embarque_LostFocus()
   Me.frm_busqueda.Visible = False
End Sub

Private Sub txt_busqueda_caja_LostFocus()
   Me.frm_busqueda.Visible = False
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   Dim ReturnFlag As String
   Dim clnt As New SoapClient30
   If KeyAscii = 13 Then
      
      If IsNumeric(Me.txt_cantidad_eliminar) Then
         If CDbl(Me.lv_salidas.selectedItem.SubItems(4)) - CDbl(Me.txt_cantidad_eliminar) >= 0 Then
            If IsNumeric(Me.txt_caja) Then
               strconsulta = "select nvl(estatus_pedido,0) from xxvia_tb_salidas_cajas where source_header_number = ? "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.txt_archivo))
                    .Parameters.Append parametro
               End With
               Set rsaux8 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               If Not rsaux8.EOF Then
                  If rsaux8(0).Value = 0 Then
                     Me.lv_salidas.selectedItem.SubItems(3) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(3)) - CDbl(Me.txt_cantidad_eliminar), "###,###,##0.00")
                     Me.lv_salidas.selectedItem.SubItems(4) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(4)) - CDbl(Me.txt_cantidad_eliminar), "###,###,##0.00")
                     Me.lv_salidas.selectedItem.SubItems(5) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(5)) + CDbl(Me.txt_cantidad_eliminar), "###,###,##0.00")
                     Me.lbl_recibidos = Format(CDbl(Me.lbl_recibidos) - CDbl(Me.txt_cantidad_eliminar), "###,###,##0.00")
                     Me.lbl_cantidad_caja = Format(CDbl(Me.lbl_cantidad_caja) - CDbl(Me.txt_cantidad_eliminar), "###,###,##0.00")
                     rsaux.Open "update XXVIA_TB_SALIDAS_CAJAS set FLOA_SAL_CANTIDAD_LEIDA = FLOA_SAL_CANTIDAD_LEIDA - " + Me.txt_cantidad_eliminar + " where inte_emb_embarque = " + Me.txt_embarque + " and SOURCE_HEADER_NUMBER = " + CStr(CDbl(Me.txt_archivo)) + " and DELIVERY_DETAIL_ID = '" + Me.lv_salidas.selectedItem.SubItems(7) + "' and inte_paq_caja = " + Me.txt_caja, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     rsaux5.Open "update TB_DETALLE_EQUIPOS_ORDEN_SURTIDO set FLOA_ORS_CANTIDAD_SURTIDA = isnull(FLOA_ORS_CANTIDAD_SURTIDA,0) - " + CStr(Me.txt_cantidad_eliminar) + " where INTE_ORS_ORDEN_SURTIDO = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
                     strconsulta = "select * from xxvia_vw_categorias_item_b where codigo = ? and organization_id = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.lv_salidas.selectedItem)
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                          .Parameters.Append parametro
                     End With
                     Set rsaux5 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                  
                     If Not rsaux5.EOF Then
                        var_linea = IIf(IsNull(rsaux5!Linea), "", rsaux5!Linea)
                     Else
                        var_linea = ""
                     End If
                     If var_linea <> "CATALOGOS" Then
                     
                        Call cantidad_leida_por_persona(CDbl(txt_cantidad_eliminar), "-")
                     End If
                     rsaux5.Close
                     rsaux.Open "SELECT TO_CHAR(SYSDATE,'dd/mm/yyyy hh24:mi:ss') AS FECHA FROM DUAL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     VAR_FECHA_HORA = rsaux(0).Value
                     rsaux.Close
                
                     rsaux.Open "INSERT INTO XXVIA_TB_BITACORA_LECTURA (PEDIDO, CAJA, CODIGO, USUARIO, CANTIDAD, FECHA_HORA,MAQUINA, DVR, PUERTO) VALUES (" + Me.txt_archivo + ", " + Me.txt_caja + ",'" + Me.lv_salidas.selectedItem + "','" + var_clave_usuario_global + "',-" + CStr(Me.txt_cantidad_eliminar) + ",TO_DATE('" + VAR_FECHA_HORA + "','dd/mm/yyyy hh24:mi:ss'),'" + fun_NombrePc + "','" + CStr(var_dvr_texto) + "','" + CStr(var_puerto_texto) + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
    
                     rsaux10.Open "select * from tb_video", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux10.EOF Then
                        V = IIf(IsNull(rsaux10(0).Value), 0, rsaux10(0).Value)
                     Else
                        V = 0
                     End If
                     rsaux10.Close
                     If V = 1 Then
                        If var_modo_texto_ip = 1 Then
                           On Error GoTo SALIR:
                           Set clnt = Nothing
                           clnt.MSSoapInit var_webservice_texto
                           var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MAQUINA: " + fun_NombrePc + ", USUARIO: " + var_nombre_usuario + Chr(13) + " FECHA Y HORA: " + VAR_FECHA_HORA + Chr(13) + Me.txt_embarque + "-" + CStr(CDbl(Me.txt_archivo)) + "-" + Me.txt_caja + "-" + Me.lv_salidas.selectedItem + "   " + Me.lv_salidas.selectedItem.SubItems(1) + " CANTIDAD ELIMINAR: " + CStr(CDbl(Me.txt_cantidad_eliminar)) + Chr(13))
                           Set clnt = Nothing
                         End If
                      End If
                     If Me.txt_codigo.Enabled = True Then
                        Me.txt_codigo.SetFocus
                     End If
                     rsaux5.Open "SELECT SUM(FLOA_SAL_CANTIDAD_LEIDA * NVL(PESO,0)) AS PESO FROM XXVIA_tB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND INTE_PAQ_CAJA = " + Me.txt_caja, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux5.EOF Then
                        Me.lbl_peso = Format(IIf(IsNull(rsaux5(0).Value), 0, rsaux5(0).Value), "###,###,##0.000")
                     Else
                        Me.lbl_peso = "0.000"
                     End If
                     rsaux5.Close
                  Else
                     Me.txt_codigo.Enabled = False
                     MsgBox "El pedido ya fue cerrado", vbOKOnly, "ATENCION"
                  End If
               
               Else
                  Me.txt_codigo.Enabled = False
                  MsgBox "El pedido ya fue cerrado", vbOKOnly, "ATENCION"
               End If
               rsaux8.Close
            Else
               MsgBox "No se a seleccionado una caja", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Cantidad a eliminar incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Cantidad a eliminar incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_eliminar.Visible = False
   End If
   Exit Sub
SALIR:
   'MsgBox Err.Description
   Resume Next
   
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   Me.frm_eliminar.Visible = False
End Sub

Private Sub txt_codigo_GotFocus()
   'Me.txt_codigo = ""
   'var_encontro = 0
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      var_caja_motor = ""
      var_si_tela = 0
      'var_codigo_aux = ""
      'var_contador_tela = 0
      'For var_j = 1 To Len(Me.txt_codigo)
      '    If Mid(Me.txt_codigo, var_j, 1) = "'" Then
      '       var_codigo_aux = var_codigo_aux + "-"
      '       var_contador_tela = var_contador_tela + 1
      '    Else
      '       var_codigo_aux = var_codigo_aux + Mid(Me.txt_codigo, var_j, 1)
      '    End If
      'Next var_j
      'If var_contador_tela = 4 Then
      '   var_si_tela = 1
      'End If
      'Me.txt_codigo = var_codigo_aux
      Dim var_tela As String
      var_tela = ""
      var_caja_motor = ""
      For var_j = 1 To Len(Me.txt_codigo)
          If Mid(Me.txt_codigo, var_j, 1) = "-" Then
             var_tela = var_tela + Mid(Me.txt_codigo, var_j, 1)
          End If
      Next var_j
      If Mid(Me.txt_codigo, 1, 2) = "CA" Or var_tela = "---" Then
         rs.Open "SELECT * FROM XXVIA_TB_CAJAS_PROD WHERE vcha_caj_caja_id = '" + UCase(Me.txt_codigo) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_almacen_destino_caja = IIf(IsNull(rs!VCHA_CAJ_DESTINO), "", rs!VCHA_CAJ_DESTINO)
            If IIf(IsNull(rs!vcha_caj_staus), "", rs!vcha_caj_staus) <> "S" Then
                
                var_caja_motor = IIf(IsNull(rs!vcha_caj_caja_id), "", rs!vcha_caj_caja_id)
                  
                If var_almacen_motor_logistico = "" Then
                   If var_almacen_destino_caja = "" Then
                      var_almacen_destino_caja = var_almacen_motor_logistico
                   End If
                End If
                If var_almacen_motor_logistico <> "" Then
                   rsaux1.Open "SELECT * FROM TB_ORACLE_UBICACIONES_MOTOR_LOGISTICO WHERE CLAVE = '" + var_almacen_motor_logistico + "' AND CODIGO ='" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                   If rsaux1.EOF Then
                      var_almacen_destino_caja = var_almacen_motor_logistico
                   End If
                   rsaux1.Close
                End If
                'If var_almacen_destino_caja = "" And var_almacen_motor_logistico <> "" Then
                '   var_almacen_destino_caja = var_almacen_motor_logistico
                'End If
                
                If var_almacen_destino_caja = var_almacen_motor_logistico Then
                   If rsaux8.State = 1 Then
                      rsaux8.Close
                   End If
                   rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + rs!VCHA_ART_ARTICULO_ID + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                   If Not rsaux8.EOF Then
                      var_peso = IIf(IsNull(rsaux8!UNIT_WEIGHT), 0, rsaux8!UNIT_WEIGHT)
                      var_salida_masiva = IIf(IsNull(rsaux8!attribute10), "N", rsaux8!attribute10)
                      var_cantidad_leida = rs!numb_caj_cantidad
                      For var_j = 1 To Me.lv_salidas.ListItems.Count
                          lv_salidas.ListItems.Item(var_j).Selected = True
                          If rs!VCHA_ART_ARTICULO_ID = lv_salidas.selectedItem And CDbl(Me.lv_salidas.selectedItem.SubItems(5)) > 0 Then
                             var_encontro = var_j
                          End If
                      Next var_j
                      If var_encontro > 0 Then
                         Me.lv_salidas.ListItems.Item(var_encontro).Selected = True
                         If CDbl(Me.lv_salidas.selectedItem.SubItems(2)) >= CDbl(Me.lv_salidas.selectedItem.SubItems(3)) + var_cantidad_leida Then
                            Me.txt_codigo = rs!VCHA_ART_ARTICULO_ID
                            Me.txt_foco.Enabled = True
                            Me.txt_foco.SetFocus
                         Else
                            Call cmd_mensaje_2_Click
                            txt_codigo = ""
                            frmmensaje.lbl_articulo = Me.lv_salidas.selectedItem.SubItems(1)
                            frmmensaje.lbl_mensaje = "La cantidad supera a la posible a surtir"
                            frmmensaje.Show 1
                         End If
                      Else
                         Call cmd_mensaje_2_Click
                         txt_codigo = ""
                         frmmensaje.lbl_articulo = ""
                         frmmensaje.lbl_mensaje = "El artículo no se encuentra en la orden de surtido"
                         frmmensaje.Show 1
                      End If
                   Else
                      Call cmd_mensaje_2_Click
                      txt_codigo = ""
                      frmmensaje.lbl_articulo = ""
                      frmmensaje.lbl_mensaje = "El artículo no se encuentra en la orden de surtido"
                      frmmensaje.Show 1
                   End If
                   rsaux8.Close
                Else
                   Call cmd_mensaje_2_Click
                   txt_codigo = ""
                   frmmensaje.lbl_articulo = ""
                   frmmensaje.lbl_mensaje = "El bulto  no pertenece al destino indicado en la orden de surtido"
                   frmmensaje.Show 1
                End If
            Else
                Call cmd_mensaje_2_Click
                txt_codigo = ""
                frmmensaje.lbl_articulo = ""
                frmmensaje.lbl_mensaje = "El bulto ya fue enviado en el lote " + IIf(IsNull(rs!pedido_almacen), "", rs!pedido_almacen)
                frmmensaje.Show 1
            End If
            
            
            
         Else
            Call cmd_mensaje_2_Click
            txt_codigo = ""
            frmmensaje.lbl_articulo = ""
            frmmensaje.lbl_mensaje = "La caja no existe"
            frmmensaje.Show 1
         End If
         rs.Close
      Else
         var_encontro = 0
         If rsaux8.State = 1 Then
            rsaux8.Close
         End If
         var_localizador_subinventario = " "
         If rsaux9.State = 1 Then
            rsaux9.Close
         End If
         strconsulta = "select * from xxvia_system_items_b where segment1 = ? and organization_id = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
              .Parameters.Append parametro
         End With
         Set rsaux9 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         'rsaux9.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If rsaux9.EOF Then
            'rsaux8.Open "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, nvl(a.description,'') as localizador, b.UNIT_WEIGHT FROM mtl_cross_references_b A, xxvia_system_items_b B WHERE A.inventory_item_id = B.inventory_item_id AND B.organization_id = " + var_unidad_organizacional + " AND CROSS_REFERENCE = '" + Me.txt_codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, NVL(a.description,'') AS localizador, B.UNIT_WEIGHT FROM (select INVENTORY_ITEM_ID, description, cross_reference from mtl_cross_references_b) A, (select inventory_item_id, DESCRIPTION, organization_id, segment1, UNIT_WEIGHT from xxvia_system_items_b) B Where a.inventory_item_id = B.inventory_item_id AND B.organization_id = " + var_unidad_organizacional + " AND CROSS_REFERENCE       = '" + Me.txt_codigo + "'"
            x = 0
            If x = 0 Then
               strconsulta = "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, NVL(a.description,'') AS localizador, B.UNIT_WEIGHT, a.cantidad FROM (select INVENTORY_ITEM_ID, description, cross_reference, nvl(attribute1,1) as cantidad from mtl_cross_references_b) A, (select inventory_item_id, DESCRIPTION, organization_id, segment1, UNIT_WEIGHT from xxvia_system_items_b) B Where a.inventory_item_id = B.inventory_item_id AND B.organization_id = ? AND CROSS_REFERENCE = ?"
              'strConsulta = "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, NVL(a.description,'') AS localizador, B.UNIT_WEIGHT FROM (select INVENTORY_ITEM_ID, description, cross_reference                                  from mtl_cross_references_b) A, (select inventory_item_id, DESCRIPTION, organization_id, segment1, UNIT_WEIGHT from xxvia_system_items_b) B Where a.inventory_item_id = B.inventory_item_id AND B.organization_id = ? AND CROSS_REFERENCE = ? "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                    .Parameters.Append parametro
               End With
               Set rsaux8 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
            End If
            'rsaux8.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cantidad_leida = 1
            If Not rsaux8.EOF Then
               var_peso = IIf(IsNull(rsaux8!UNIT_WEIGHT), 0, rsaux8!UNIT_WEIGHT)
               var_cantidad_leida = IIf(IsNull(rsaux8!cantidad), 1, rsaux8!cantidad)
               If IIf(IsNull(rsaux8!localizador), "", rsaux8!localizador) <> "" Then
                  var_localizador_subinventario = txt_almacen + IIf(IsNull(rsaux8!localizador), "", rsaux8!localizador)
                  If var_localizador_subinventario <> "" Then
                     Me.txt_codigo = rsaux8!SEGMENT1
                  Else
                     Me.txt_codigo = ""
                  End If
               Else
                  Me.txt_codigo = rsaux8!SEGMENT1
                  'Me.txt_codigo = ""
               End If
            Else
               Me.txt_codigo = ""
            End If
            rsaux8.Close
         Else
            var_cantidad_leida = 1
         End If
         rsaux9.Close
         If Me.txt_codigo <> "" Then
            x = 0
            If x = 0 Then
               strconsulta = "select * from xxvia_system_items_b where segment1 = ? and organization_id = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                    .Parameters.Append parametro
               End With
               Set rsaux8 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
            End If
            'rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
               var_peso = IIf(IsNull(rsaux8!UNIT_WEIGHT), 0, rsaux8!UNIT_WEIGHT)
               'var_salida_masiva = "Y"
               If var_cantidad_leida > 1 Then
                  var_salida_masiva = "N"
               Else
                  var_salida_masiva = IIf(IsNull(rsaux8!attribute10), "N", rsaux8!attribute10)
               End If
               
               If var_salida_masiva = "Y" Then
                  var_codigo_global = Me.txt_codigo
                  frmoracle_cantidad.Show 1
                  var_cantidad_leida = var_cantidad_global
                  Me.txt_codigo = var_codigo_global
               Else
                  'var_cantidad_leida = 1
                  var_cantidad_leida = var_cantidad_leida
               End If
               VAR_PIEZAS = 0
               If var_almacen_motor_logistico <> "" Then
                  If Me.txt_codigo = "00035161" Then
                     VAR_PIEZAS = 1
                  End If
               Else
                  VAR_PIEZAS = 0
               End If
               If VAR_PIEZAS = 0 Then
                  For var_j = 1 To Me.lv_salidas.ListItems.Count
                      lv_salidas.ListItems.Item(var_j).Selected = True
                      If Me.txt_codigo = lv_salidas.selectedItem And CDbl(Me.lv_salidas.selectedItem.SubItems(5)) > 0 Then
                         var_encontro = var_j
                      End If
                  Next var_j
                  If var_encontro > 0 Then
                     Me.lv_salidas.ListItems.Item(var_encontro).Selected = True
                     If CDbl(Me.lv_salidas.selectedItem.SubItems(2)) >= CDbl(Me.lv_salidas.selectedItem.SubItems(3)) + var_cantidad_leida Then
                        Me.txt_foco.Enabled = True
                        Me.txt_foco.SetFocus
                     Else
                        Call cmd_mensaje_2_Click
                        txt_codigo = ""
                        frmmensaje.lbl_articulo = Me.lv_salidas.selectedItem.SubItems(1)
                        frmmensaje.lbl_mensaje = "La cantidad supera a la posible a surtir"
                        frmmensaje.Show 1
                     End If
                  Else
                     Call cmd_mensaje_2_Click
                     txt_codigo = ""
                     frmmensaje.lbl_articulo = ""
                     frmmensaje.lbl_mensaje = "El artículo no se encuentra en la orden de surtido"
                     frmmensaje.Show 1
                  End If
               Else
                  Call cmd_mensaje_2_Click
                  txt_codigo = ""
                  frmmensaje.lbl_articulo = ""
                  frmmensaje.lbl_mensaje = "El artículo no se puede leer pieza a pieza"
                  frmmensaje.Show 1
               End If
            Else
               Call cmd_mensaje_2_Click
               txt_codigo = ""
               frmmensaje.lbl_articulo = ""
               frmmensaje.lbl_mensaje = "El artículo no se encuentra en la orden de surtido"
               frmmensaje.Show 1
            End If
            rsaux8.Close
         Else
            Call cmd_mensaje_2_Click
            txt_codigo = ""
            frmmensaje.lbl_articulo = ""
            frmmensaje.lbl_mensaje = "El artículo no existe"
            frmmensaje.Show 1
         End If
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   Dim InputData As String
   Dim ErrorFlag
   Dim StartTime
   Dim RequestString As String
   Dim i As Single
   Dim ReturnFlag As String
   Dim clnt As New SoapClient30
   If Trim(Me.txt_codigo) <> "" Then
      If var_encontro > 0 Then
         rs.Open "SELECT TO_CHAR(SYSDATE,'dd/mm/yyyy hh24:mi:ss') AS FECHA FROM DUAL", cnnoracle_4, adOpenDynamic, adLockOptimistic
         VAR_FECHA_HORA = rs(0).Value
         rs.Close
         If CDbl(Me.lv_salidas.selectedItem.SubItems(2) >= CDbl(Me.lv_salidas.selectedItem.SubItems(3)) + var_cantidad_leida) Then
            If rsaux1.State = 1 Then
               rsaux1.Close
            End If
            var_posibe_maquina = 1
            If var_bandera_asignacion = 0 Then
               rs.Open "SELECT * FROM tb_oracle_pedidos_maquinas where pedido = " + CStr(CDbl(Me.txt_archivo)), cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  If rs!maquina = fun_NombrePc Then
                     var_posibe_maquina = 1
                  Else
                     var_posibe_maquina = 0
                  End If
               Else
                  var_posibe_maquina = 1
               End If
               rs.Close
            End If
            If var_posibe_maquina = 1 Then
               rsaux1.Open "SELECT * FROM TB_ORACLE_PEDIDOS_MAQUINAS WHERE PEDIDO = " + CStr(CDbl(Me.txt_archivo)), cnn, adOpenDynamic, adLockOptimistic
               If rsaux1.EOF Then
                  rsaux2.Open "INSERT INTO TB_ORACLE_PEDIDOS_MAQUINAS (MAQUINA, PEDIDO, USUARIO) VALUES ('" + fun_NombrePc + "'," + Me.txt_archivo + ",'" + var_clave_usuario_global + "')", cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux1.Close
               rsaux1.Open "SELECT * FROM TB_ORACLE_EMBARQUES_ORDENES WHERE source_header_number = " + CStr(CDbl(Me.txt_archivo)), cnn, adOpenDynamic, adLockOptimistic
               If rsaux1.EOF Then
                  rs.Open "select * from tb_oracle_embarques_ordenes where  INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND source_header_number = " + CStr(CDbl(Me.txt_archivo)), cnn, adOpenDynamic, adLockOptimistic
                  If rs.EOF Then
                     rsaux.Open "INSERT INTO TB_ORACLE_EMBARQUES_ORDENES (INTE_EMB_EMBARQUE, source_header_number) VALUES (" + Me.txt_embarque + "," + CStr(CDbl(Me.txt_archivo)) + ")", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rs.Close
                  If var_primera_vez = 1 Then
                     cnn.BeginTrans
                     rsaux11.Open "select max(inte_tvf_consecutivo) from tb_temp_valuacion_facturacion", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux11.EOF Then
                        var_consecutivo = IIf(IsNull(rsaux11(0).Value), 0, rsaux11(0).Value)
                     Else
                        var_consecutivo = 0
                     End If
                     var_consecutivo = var_consecutivo + 1
                     rsaux11.Close
                     rs.Open "Insert into tb_temp_valuacion_facturacion (INTE_TVF_CONSECUTIVO, vcha_aud_usuario, vcha_aud_maquina) values (" + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic
                     
                     rs.Open "select max(inte_paq_caja) from XXVIA_TB_SALIDAS_CAJAS where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     x = 0
                     If x = 1 Then
                        strconsulta = "select max(inte_paq_caja) from XXVIA_TB_SALIDAS_CAJAS where inte_emb_embarque = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
                             .Parameters.Append parametro
                        End With
                        Set rs = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                     End If
                     If Not rs.EOF Then
                        Me.txt_caja = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                     Else
                        Me.txt_caja = 1
                     End If
                     rs.Close
                  
                     rs.Open "select max(caja_pedido) from XXVIA_TB_SALIDAS_CAJAS where inte_emb_embarque = " + Me.txt_embarque + " and source_header_number = " + CStr(CDbl(Me.txt_archivo)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     x = 0
                     If x = 1 Then
                        strconsulta = "select max(inte_paq_caja) from XXVIA_TB_SALIDAS_CAJAS where inte_emb_embarque = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
                             .Parameters.Append parametro
                        End With
                        Set rs = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                     End If
                     If Not rs.EOF Then
                        Me.txt_caja_pedido = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                     Else
                        Me.txt_caja_pedido = 1
                     End If
                     rs.Close
                     
                     var_numero_caja = Me.txt_caja
                     var_referencia_caja = ""
                     var_contador = 0
                     If Len(Trim(Str(var_numero_caja))) = 1 Then
                        var_referencia_caja = "00" + Trim(Str(var_numero_caja))
                     End If
                     If Len(Trim(Str(var_numero_caja))) = 2 Then
                        var_referencia_caja = "0" + Trim(Str(var_numero_caja))
                     End If
                     If Len(Trim(Str(var_numero_caja))) = 3 Then
                        var_referencia_caja = Trim(Str(var_numero_caja))
                     End If
                     If Len(Trim(Str(txt_embarque))) = 1 Then
                        var_referencia_embarque = "00000" + Trim(Str(txt_embarque))
                     End If
                     If Len(Trim(Str(txt_embarque))) = 2 Then
                        var_referencia_embarque = "0000" + Trim(Str(txt_embarque))
                     End If
                     If Len(Trim(Str(txt_embarque))) = 3 Then
                        var_referencia_embarque = "000" + Trim(Str(txt_embarque))
                     End If
                     If Len(Trim(Str(txt_embarque))) = 4 Then
                        var_referencia_embarque = "00" + Trim(Str(txt_embarque))
                     End If
                     If Len(Trim(Str(txt_embarque))) = 5 Then
                        var_referencia_embarque = "0" + Trim(Str(txt_embarque))
                     End If
                     If Len(Trim(Str(txt_embarque))) = 6 Then
                        var_referencia_embarque = Trim(Str(txt_embarque))
                     End If
                     On Error GoTo SALIR:
                     rsaux12.Open "insert into TB_ORACLE_CAJAS_UNICAS_EMBARQUES (caja, usuario, maquina) values ('C" + var_referencia_embarque + var_referencia_caja + "','" + var_clave_usuario_global + "','" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic
                     
                     
                     
                     var_primera_vez = 0
                     
                     
                     cnn.CommitTrans
                  End If
                  'rs.Open "     SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND source_header_number = " + CStr(CDbl(Me.txt_archivo)) + " AND SEGMENT1 = '" + Me.txt_codigo + "' and inte_paq_caja = " + Me.txt_caja + " AND DELIVERY_DETAIL_ID = " + Me.lv_salidas.selectedItem.SubItems(7), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  strconsulta = "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE =     ?                   AND source_header_number = ? AND SEGMENT1 = ? and inte_paq_caja = ? AND DELIVERY_DETAIL_ID = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_archivo))
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CDbl(Me.txt_codigo))
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_caja))
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_salidas.selectedItem.SubItems(7)))
                       .Parameters.Append parametro
                  End With
                  Set rs = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  
                  If rs.EOF Then
                     var_cadena = "INSERT INTO XXVIA_TB_SALIDAS_CAJAS (INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER, SEGMENT1, FLOA_SAL_CANTIDAD_LEIDA, INVENTORY_ITEM_ID, DELIVERY_DETAIL_ID, SOURCE_LINE_NUMBER, DELIVERY_ID, INTE_PAQ_CAJA, CUSTOMER_ID, SUBINVENTORY, NAME, COLLECTOR_ID, ITEM_DESCRIPTION, CUSTOMER_NAME, TIPO_cAJA, CAJA_PEDIDO,PESO, ENTREGA)"
                     var_cadena = var_cadena + " values (" + Me.txt_embarque + "," + CStr(CDbl(Me.txt_archivo)) + ",'" + Me.txt_codigo + "'," + CStr(var_cantidad_leida) + "," + lv_salidas.selectedItem.SubItems(6) + "," + Me.lv_salidas.selectedItem.SubItems(7) + "," + Me.lv_salidas.selectedItem.SubItems(8) + "," + Me.lv_salidas.selectedItem.SubItems(9) + "," + Me.txt_caja + "," + Me.lv_salidas.selectedItem.SubItems(10) + ",'" + Me.txt_origen + "', '" + Me.txt_agente + "','" + lv_salidas.selectedItem.SubItems(11) + "','" + lv_salidas.selectedItem.SubItems(1) + "','" + Me.txt_cliente + "','" + var_nombre_caja + "'," + Me.txt_caja_pedido + "," + CStr(var_peso) + ",'" + Replace(Me.txt_entrega, "'", " ") + "') "
                     rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     rsaux.Open "INSERT INTO XXVIA_TB_BITACORA_LECTURA (PEDIDO, CAJA, CODIGO, USUARIO, CANTIDAD, FECHA_HORA, MAQUINA, DVR, PUERTO) VALUES (" + Me.txt_archivo + ", " + Me.txt_caja + ",'" + Me.txt_codigo + "','" + var_clave_usuario_global + "'," + CStr(var_cantidad_leida) + ",TO_DATE('" + VAR_FECHA_HORA + "','dd/mm/yyyy hh24:mi:ss'),'" + fun_NombrePc + "','" + CStr(var_dvr_texto) + "','" + CStr(var_puerto_texto) + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
'-- actualiza caja

                           strconsulta = "update XXVIA_TB_CAJAS_PROD  set vcha_caj_staus = 'S', PEDIDO_ALMACEN =?, USUARIO_almacen = ?, MAQUINA_ALMACEN = ?  where vcha_caj_caja_id = ?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Me.txt_archivo)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_clave_usuario_global)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, fun_NombrePc)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_caja_motor)
                                .Parameters.Append parametro
                           
                           End With
                           Set rsaux = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing




'--
                  
                  
                  
                  Else
                     'rsaux.Open "update XXVIA_TB_SALIDAS_CAJAS set FLOA_SAL_CANTIDAD_LEIDA = FLOA_SAL_CANTIDAD_LEIDA + " + CStr(var_cantidad_leida) + ", PESO = " + CStr(var_peso) + ", ENTREGA = '" + Replace(Me.txt_entrega, "'", " ") + "' where inte_emb_embarque = " + Me.txt_embarque + " and SOURCE_HEADER_NUMBER = " + CStr(CDbl(Me.txt_archivo)) + " and segment1 = '" + Me.txt_codigo + "' and inte_paq_caja = " + Me.txt_caja + " AND DELIVERY_dETAIL_ID = " + Me.lv_salidas.selectedItem.SubItems(7), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     strconsulta = "update XXVIA_TB_SALIDAS_CAJAS set FLOA_SAL_CANTIDAD_LEIDA = FLOA_SAL_CANTIDAD_LEIDA + ?, PESO = ?, ENTREGA = ? where inte_emb_embarque = ? and SOURCE_HEADER_NUMBER = ? and segment1 = ? and inte_paq_caja = ? AND DELIVERY_dETAIL_ID = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_cantidad_leida)
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_peso)
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Replace(Me.txt_entrega, "'", " "))
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_archivo))
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 300, Me.txt_codigo)
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 300, CDbl(Me.txt_caja))
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.lv_salidas.selectedItem.SubItems(7)))
                          .Parameters.Append parametro
                     End With
                     Set rsaux = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     
                     
'-- actualiza caja

                           strconsulta = "update XXVIA_TB_CAJAS_PROD  set vcha_caj_staus = 'S', PEDIDO_ALMACEN =?, USUARIO_almacen = ?, MAQUINA_ALMACEN = ?  where vcha_caj_caja_id = ?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Me.txt_archivo)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_clave_usuario_global)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, fun_NombrePc)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_caja_motor)
                                .Parameters.Append parametro
                           
                           End With
                           Set rsaux = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing

'--
                     
                     
                     rsaux.Open "INSERT INTO XXVIA_TB_BITACORA_LECTURA (PEDIDO, CAJA, CODIGO, USUARIO, CANTIDAD, FECHA_HORA, MAQUINA, DVR, PUERTO) VALUES (" + Me.txt_archivo + ", " + Me.txt_caja + ",'" + Me.txt_codigo + "','" + var_clave_usuario_global + "'," + CStr(var_cantidad_leida) + ",TO_DATE('" + VAR_FECHA_HORA + "','dd/mm/yyyy hh24:mi:ss'),'" + fun_NombrePc + "','" + CStr(var_dvr_texto) + "','" + CStr(var_puerto_texto) + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  
                  rsaux5.Open "update TB_DETALLE_EQUIPOS_ORDEN_SURTIDO set FLOA_ORS_CANTIDAD_SURTIDA = isnull(FLOA_ORS_CANTIDAD_SURTIDA,0) + " + CStr(var_cantidad_leida) + " where INTE_ORS_ORDEN_SURTIDO = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
                  strconsulta = "select * from xxvia_vw_categorias_item_b where codigo = ? and organization_id = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.lv_salidas.selectedItem)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                       .Parameters.Append parametro
                  End With
                  Set rsaux5 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  
                  If Not rsaux5.EOF Then
                     var_linea = IIf(IsNull(rsaux5!Linea), "", rsaux5!Linea)
                  Else
                     var_linea = ""
                  End If
                  If var_linea <> "CATALOGOS" Then
                     Call cantidad_leida_por_persona(var_cantidad_leida, "+")
                  End If
                  rsaux5.Close
                  Me.lv_salidas.selectedItem.SubItems(3) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(3)) + var_cantidad_leida, "###,###,##0.00")
                  Me.lv_salidas.selectedItem.SubItems(4) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(4)) + var_cantidad_leida, "###,###,##0.00")
                  Me.lv_salidas.selectedItem.SubItems(5) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(2)) - (CDbl(Me.lv_salidas.selectedItem.SubItems(3))), "###,###,##0.00")
                  Me.lbl_recibidos = Format(CDbl(Me.lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                  Me.lbl_cantidad_caja = Format(CDbl(lbl_cantidad_caja) + var_cantidad_leida, "###,###,##0.00")
                  Me.txt_codigo.SetFocus
                  rs.Close
                  rsaux5.Open "SELECT SUM(FLOA_SAL_CANTIDAD_LEIDA * NVL(PESO,0)) AS PESO FROM XXVIA_tB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND INTE_PAQ_CAJA = " + Me.txt_caja, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux5.EOF Then
                     Me.lbl_peso = Format(IIf(IsNull(rsaux5(0).Value), 0, rsaux5(0).Value), "###,###,##0.000")
                  Else
                     Me.lbl_peso = "0.000"
                  End If
                  rsaux5.Close
                  
                  Call cmd_mensaje_4_Click
                  var_renglon = lv_salidas.selectedItem.Index
                  Call ilumina_grid
                  rsaux10.Open "select * from tb_video", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux10.EOF Then
                     V = IIf(IsNull(rsaux10(0).Value), 0, rsaux10(0).Value)
                  Else
                     V = 0
                  End If
                  rsaux10.Close
                  If V = 1 Then
                     If var_modo_texto_ip = 1 Then
                        On Error GoTo SALIR:
                        Set clnt = Nothing
                        clnt.MSSoapInit var_webservice_texto
                        '1
                        var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MAQUINA: " + fun_NombrePc + ", USUARIO: " + var_nombre_usuario + Chr(13) + " FECHA Y HORA: " + VAR_FECHA_HORA + Chr(13) + Me.txt_embarque + "-" + CStr(CDbl(Me.txt_archivo)) + "-" + Me.txt_caja + "-" + Me.txt_codigo + "   " + Me.lv_salidas.selectedItem.SubItems(1) + " CANTIDAD: " + CStr(var_cantidad_leida) + Chr(13))
                        Set clnt = Nothing
                     Else
                        On Error GoTo SALIR:
                        'If mscomm1.PortOpen = True Then
                           'mscomm1.PortOpen = False
                        'End If
                        'mscomm1.CommPort = 1
                        'mscomm1.settings = var_baudios
                        'mscomm1.PortOpen = True
                        'mscomm1.Output = "@B@ " + Chr(13)
                        'mscomm1.Output = Me.txt_embarque + "-" + Me.txt_caja + "-" + Me.txt_codigo + "   " + Me.lv_salidas.selectedItem.SubItems(1) + " CANTIDAD: " + CStr(var_cantidad_leida) + "^]EOL" + Chr(13)
                        'mscomm1.Output = "@E@ "
                        'mscomm1.OutBufferCount = 0
                        'mscomm1.PortOpen = False
                     End If
                  End If
               Else
                  If rsaux1!inte_Emb_Embarque = CDbl(Me.txt_embarque) Then
                     If var_primera_vez = 1 Then
                        cnn.BeginTrans
                        rsaux11.Open "select max(inte_tvf_consecutivo) from tb_temp_valuacion_facturacion", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux11.EOF Then
                           var_consecutivo = IIf(IsNull(rsaux11(0).Value), 0, rsaux11(0).Value)
                        Else
                           var_consecutivo = 0
                        End If
                        var_consecutivo = var_consecutivo + 1
                        rsaux11.Close
                        rs.Open "Insert into tb_temp_valuacion_facturacion (INTE_TVF_CONSECUTIVO, vcha_aud_usuario, vcha_aud_maquina) values (" + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic
                        
                        'rs.Open "select max(inte_paq_caja) from XXVIA_TB_SALIDAS_CAJAS where inte_emb_embarque = " + Me.txt_embarque + " and source_header_number = " + CStr(CDbl(Me.txt_archivo)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rs.Open "select max(inte_paq_caja) from XXVIA_TB_SALIDAS_CAJAS where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           Me.txt_caja = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                        Else
                           Me.txt_caja = 1
                        End If
                        rs.Close
                        
                        rs.Open "select max(caja_pedido) from XXVIA_TB_SALIDAS_CAJAS where inte_emb_embarque = " + Me.txt_embarque + " and source_header_number = " + CStr(CDbl(Me.txt_archivo)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           Me.txt_caja_pedido = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                        Else
                           Me.txt_caja_pedido = 1
                        End If
                        rs.Close
                        
                        
                        var_numero_caja = Me.txt_caja
                        var_referencia_caja = ""
                        var_contador = 0
                        If Len(Trim(Str(var_numero_caja))) = 1 Then
                           var_referencia_caja = "00" + Trim(Str(var_numero_caja))
                        End If
                        If Len(Trim(Str(var_numero_caja))) = 2 Then
                           var_referencia_caja = "0" + Trim(Str(var_numero_caja))
                        End If
                        If Len(Trim(Str(var_numero_caja))) = 3 Then
                           var_referencia_caja = Trim(Str(var_numero_caja))
                        End If
                        If Len(Trim(Str(txt_embarque))) = 1 Then
                           var_referencia_embarque = "00000" + Trim(Str(txt_embarque))
                        End If
                        If Len(Trim(Str(txt_embarque))) = 2 Then
                           var_referencia_embarque = "0000" + Trim(Str(txt_embarque))
                        End If
                        If Len(Trim(Str(txt_embarque))) = 3 Then
                           var_referencia_embarque = "000" + Trim(Str(txt_embarque))
                        End If
                        If Len(Trim(Str(txt_embarque))) = 4 Then
                           var_referencia_embarque = "00" + Trim(Str(txt_embarque))
                        End If
                        If Len(Trim(Str(txt_embarque))) = 5 Then
                           var_referencia_embarque = "0" + Trim(Str(txt_embarque))
                        End If
                        If Len(Trim(Str(txt_embarque))) = 6 Then
                           var_referencia_embarque = Trim(Str(txt_embarque))
                        End If
                        On Error GoTo SALIR:
                        rsaux12.Open "insert into TB_ORACLE_CAJAS_UNICAS_EMBARQUES (caja, usuario, maquina) values ('C" + var_referencia_embarque + var_referencia_caja + "','" + var_clave_usuario_global + "','" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic
                        var_primera_vez = 0
                        cnn.CommitTrans
                     End If
                     rsaux5.Open "SELECT SUM(FLOA_SAL_CANTIDAD_LEIDA * NVL(PESO,0)) AS PESO FROM XXVIA_tB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND INTE_PAQ_CAJA = " + Me.txt_caja, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_peso_general = IIf(IsNull(rsaux5(0).Value), 0, rsaux5(0).Value)
                     rsaux5.Close
                     If CDbl(Me.lbl_maximo) = 0 Then
                        var_si = 1
                     Else
                        If var_peso_general + var_peso > CDbl(Me.lbl_maximo) Then
                           Call cmd_mensaje_2_Click
                           txt_codigo = ""
                           frmmensaje.lbl_mensaje = "El peso supera al maximo permitido"
                           frmmensaje.Show 1
                           Call cmd_imprimir_Click
                           var_si = 0
                        Else
                           var_si = 1
                        End If
                     End If
                     If var_si = 1 Then
                        'rs.Open "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND source_header_number = " + CStr(CDbl(Me.txt_archivo)) + " AND SEGMENT1 = '" + Me.txt_codigo + "' and inte_paq_caja = " + Me.txt_caja + " AND DELIVERY_DETAIL_ID = " + Me.lv_salidas.selectedItem.SubItems(7), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        strconsulta = "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = ? AND source_header_number = ? AND SEGMENT1 = ? and inte_paq_caja = ? AND DELIVERY_DETAIL_ID = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_archivo))
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(Me.txt_codigo))
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_caja))
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_salidas.selectedItem.SubItems(7)))
                             .Parameters.Append parametro
                        End With
                        Set rs = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        
                        If rs.EOF Then
                           var_cadena = "INSERT INTO XXVIA_TB_SALIDAS_CAJAS (INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER, SEGMENT1, FLOA_SAL_CANTIDAD_LEIDA, INVENTORY_ITEM_ID, DELIVERY_DETAIL_ID, SOURCE_LINE_NUMBER, DELIVERY_ID, INTE_PAQ_CAJA, CUSTOMER_ID, SUBINVENTORY, NAME, COLLECTOR_ID, ITEM_DESCRIPTION, CUSTOMER_NAME, TIPO_cAJA, CAJA_PEDIDO, PESO, ENTREGA)"
                           var_cadena = var_cadena + " values (" + Me.txt_embarque + "," + CStr(CDbl(Me.txt_archivo)) + ",'" + Me.txt_codigo + "'," + CStr(var_cantidad_leida) + "," + lv_salidas.selectedItem.SubItems(6) + "," + Me.lv_salidas.selectedItem.SubItems(7) + ",'" + Me.lv_salidas.selectedItem.SubItems(8) + "'," + Me.lv_salidas.selectedItem.SubItems(9) + "," + Me.txt_caja + "," + Me.lv_salidas.selectedItem.SubItems(10) + ",'" + Me.txt_origen + "', '" + Me.txt_agente + "','" + Me.lv_salidas.selectedItem.SubItems(11) + "','" + Me.lv_salidas.selectedItem.SubItems(1) + "','" + Me.txt_cliente + "','" + var_nombre_caja + "'," + Me.txt_caja_pedido + "," + CStr(var_peso) + ",'" + Replace(Me.txt_entrega, "'", " ") + "') "
                           rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           rsaux.Open "INSERT INTO XXVIA_TB_BITACORA_LECTURA (PEDIDO, CAJA, CODIGO, USUARIO, CANTIDAD, FECHA_HORA, MAQUINA, DVR, PUERTO) VALUES (" + Me.txt_archivo + ", " + Me.txt_caja + ",'" + Me.txt_codigo + "','" + var_clave_usuario_global + "'," + CStr(var_cantidad_leida) + ",TO_DATE('" + VAR_FECHA_HORA + "','dd/mm/yyyy hh24:mi:ss'),'" + fun_NombrePc + "'," + var_dvr_texto + "," + var_puerto_texto + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
'-- actualiza caja

                           strconsulta = "update XXVIA_TB_CAJAS_PROD  set vcha_caj_staus = 'S', PEDIDO_ALMACEN =?, USUARIO_almacen = ?, MAQUINA_ALMACEN = ?  where vcha_caj_caja_id = ?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Me.txt_archivo)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_clave_usuario_global)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, fun_NombrePc)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_caja_motor)
                                .Parameters.Append parametro
                           
                           End With
                           Set rsaux = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing



'--
                           
                           
                           
                        Else
                             'rsaux.Open "update XXVIA_TB_SALIDAS_CAJAS set FLOA_SAL_CANTIDAD_LEIDA = FLOA_SAL_CANTIDAD_LEIDA + " + CStr(var_cantidad_leida) + ", PESO = " + CStr(var_peso) + ", ENTREGA = '" + Replace(Me.txt_entrega, "'", " ") + "' where inte_emb_embarque = " + Me.txt_embarque + " and SOURCE_HEADER_NUMBER = " + CStr(CDbl(Me.txt_archivo)) + " and segment1 = '" + Me.txt_codigo + "' and inte_paq_caja = " + Me.txt_caja + " AND DELIVERY_dETAIL_ID = " + Me.lv_salidas.selectedItem.SubItems(7), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           strconsulta = "update XXVIA_TB_SALIDAS_CAJAS set FLOA_SAL_CANTIDAD_LEIDA = FLOA_SAL_CANTIDAD_LEIDA + ?, PESO = ?, ENTREGA = ? where inte_emb_embarque = ? and SOURCE_HEADER_NUMBER = ? and segment1 = ? and inte_paq_caja = ? AND DELIVERY_dETAIL_ID = ?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_cantidad_leida)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_peso)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Replace(Me.txt_entrega, "'", " "))
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_archivo))
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 300, Me.txt_codigo)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 300, CDbl(Me.txt_caja))
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.lv_salidas.selectedItem.SubItems(7)))
                                .Parameters.Append parametro
                           End With
                           Set rsaux = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           
'-- actualiza caja

                           strconsulta = "update XXVIA_TB_CAJAS_PROD  set vcha_caj_staus = 'S', PEDIDO_ALMACEN =?, USUARIO_almacen = ?, MAQUINA_ALMACEN = ?  where vcha_caj_caja_id = ?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Me.txt_archivo)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_clave_usuario_global)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, fun_NombrePc)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_caja_motor)
                                .Parameters.Append parametro
                           
                           End With
                           Set rsaux = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing



'--
                           
                           
                           
                           
                           rsaux.Open "INSERT INTO XXVIA_TB_BITACORA_LECTURA (PEDIDO, CAJA, CODIGO, USUARIO, CANTIDAD, FECHA_HORA, MAQUINA, DVR, PUERTO) VALUES (" + Me.txt_archivo + ", " + Me.txt_caja + ",'" + Me.txt_codigo + "','" + var_clave_usuario_global + "'," + CStr(var_cantidad_leida) + ",TO_DATE('" + VAR_FECHA_HORA + "','dd/mm/yyyy hh24:mi:ss'),'" + fun_NombrePc + "'," + var_dvr_texto + "," + var_puerto_texto + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        rs.Close
                        rsaux5.Open "update TB_DETALLE_EQUIPOS_ORDEN_SURTIDO set FLOA_ORS_CANTIDAD_SURTIDA = isnull(FLOA_ORS_CANTIDAD_SURTIDA,0) + " + CStr(var_cantidad_leida) + " where INTE_ORS_ORDEN_SURTIDO = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
                        rsaux5.Open "update TB_DETALLE_EQUIPOS_ORDEN_SURTIDO set FLOA_ORS_CANTIDAD_SURTIDA = isnull(FLOA_ORS_CANTIDAD_SURTIDA,0) + " + CStr(var_cantidad_leida) + " where INTE_ORS_ORDEN_SURTIDO = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
                        strconsulta = "select * from xxvia_vw_categorias_item_b where codigo = ? and organization_id = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.lv_salidas.selectedItem)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                             .Parameters.Append parametro
                        End With
                        Set rsaux5 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        
                        If Not rsaux5.EOF Then
                           var_linea = IIf(IsNull(rsaux5!Linea), "", rsaux5!Linea)
                        Else
                           var_linea = ""
                        End If
                        If var_linea <> "CATALOGOS" Then
                            Call cantidad_leida_por_persona(var_cantidad_leida, "+")
                        End If
                        rsaux5.Close
                        Me.lv_salidas.selectedItem.SubItems(3) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(3)) + var_cantidad_leida, "###,###,##0.00")
                        Me.lv_salidas.selectedItem.SubItems(4) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(4)) + var_cantidad_leida, "###,###,##0.00")
                        Me.lv_salidas.selectedItem.SubItems(5) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(2)) - (CDbl(Me.lv_salidas.selectedItem.SubItems(3))), "###,###,##0.00")
                        Me.lbl_recibidos = Format(CDbl(Me.lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                        Me.lbl_cantidad_caja = Format(CDbl(lbl_cantidad_caja) + var_cantidad_leida, "###,###,##0.00")
                        rsaux5.Open "SELECT SUM(FLOA_SAL_CANTIDAD_LEIDA * NVL(PESO,0)) AS PESO FROM XXVIA_tB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND INTE_PAQ_CAJA = " + Me.txt_caja, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux5.EOF Then
                           Me.lbl_peso = Format(IIf(IsNull(rsaux5(0).Value), 0, rsaux5(0).Value), "###,###,##0.000")
                        Else
                           Me.lbl_peso = "0.000"
                        End If
                        rsaux5.Close
                        Call cmd_mensaje_4_Click
                        If Me.txt_codigo.Enabled = True Then
                           Me.txt_codigo.SetFocus
                        End If
                        var_renglon = lv_salidas.selectedItem.Index
                        Call ilumina_grid
                  
                        rsaux10.Open "select * from tb_video", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux10.EOF Then
                           V = IIf(IsNull(rsaux10(0).Value), 0, rsaux10(0).Value)
                        Else
                           V = 0
                        End If
                        rsaux10.Close
                        If V = 1 Then
                           If var_modo_texto_ip = 1 Then
                              On Error GoTo SALIR:
                              Set clnt = Nothing
                              clnt.MSSoapInit var_webservice_texto
                              var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MAQUINA: " + fun_NombrePc + ", USUARIO: " + var_nombre_usuario + Chr(13) + " FECHA Y HORA: " + VAR_FECHA_HORA + Chr(13) + Me.txt_embarque + "-" + CStr(CDbl(Me.txt_archivo)) + "-" + Me.txt_caja + "-" + Me.txt_codigo + "   " + Me.lv_salidas.selectedItem.SubItems(1) + " CANTIDAD: " + CStr(var_cantidad_leida) + Chr(13))
                              Set clnt = Nothing
                           Else
                           
                              On Error GoTo SALIR:
                              'If mscomm1.PortOpen = True Then
                                 'mscomm1.PortOpen = False
                              'End If
                              'mscomm1.CommPort = 1
                              'mscomm1.settings = var_baudios
                              'mscomm1.PortOpen = True
                              'mscomm1.Output = "@B@ " + Chr(13) + Chr(10)
                              'mscomm1.Output = Me.txt_embarque + "-" + Me.txt_caja + "-" + Me.txt_codigo + "   " + Me.lv_salidas.selectedItem.SubItems(1) + " CANTIDAD: " + CStr(var_cantidad_leida) + "^]EOL" + Chr(13) + Chr(10)
                              'mscomm1.Output = "@E@ "
                              'mscomm1.OutBufferCount = 0
                              'mscomm1.PortOpen = False
                           End If
                        End If
                        
                     End If
                  Else
                     Call cmd_mensaje_2_Click
                     txt_codigo = ""
                     rsaux1.Open "SELECT dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_EMBARQUES.INTE_JAU_JAULA_ID, dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID, dbo.TB_USUARIOS.VCHA_USU_NOMBRE, dbo.TB_ENCABEZADO_EMBARQUES.VCHA_AUD_MAQUINA, dbo.Tb_usuarios.VCHA_USU_APELLIDOS FROM dbo.TB_ENCABEZADO_EMBARQUES INNER JOIN dbo.TB_USUARIOS ON dbo.TB_ENCABEZADO_EMBARQUES.VCHA_AUD_USUARIO = dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID Where (dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE = " + CStr(rsaux!inte_Emb_Embarque) + ")", cnn, adOpenDynamic, adLockOptimistic
                     frmmensaje.lbl_articulo = "La orden de surtido se encuentra en el embarque " + CStr(rsaux1!inte_Emb_Embarque)
                     frmmensaje.lbl_mensaje = " en la máquina " + IIf(IsNull(rsaux1!vcha_aud_maquina), "", rsaux1!vcha_aud_maquina) + " con el usuario " + IIf(IsNull(rsaux1!vcha_usu_nombre), "", rsaux1!vcha_usu_nombre) + " " + IIf(IsNull(rsaux1!vcha_usu_apellidos), "", rsaux1!vcha_usu_apellidos)
                     rsaux1.Close
                     Me.txt_codigo.Enabled = False
                     frmmensaje.Show 1
                  End If
               End If
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
            Else
               MsgBox "El pedido ya esta siendo utilizado en la máquina", vbOKOnly, "ATENCION"
            End If
         Else
            Call cmd_mensaje_2_Click
            txt_codigo = ""
            frmmensaje.lbl_articulo = "La cantidad supera a la posible a surtir"
            frmmensaje.lbl_mensaje = ""
            Me.txt_codigo.Enabled = False
            frmmensaje.Show 1
         End If
         Me.txt_codigo = ""
         var_encontro = 0
      End If
   End If
   Exit Sub
SALIR:
   'MsgBox Err.Number
   If Err.Number = -2147217873 Then
      cnn.RollbackTrans
      Call cmd_mensaje_2_Click
      txt_codigo = ""
      frmmensaje.lbl_mensaje = "Error al leer el artículo, intentelo nuevamente"
      frmmensaje.Show 1
   Else
      Resume Next
   End If
End Sub

Private Sub txt_foco_LostFocus()
   Me.txt_foco.Enabled = False
End Sub

