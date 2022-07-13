VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmsalidas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salidas"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   Icon            =   "frmsalidas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11640
   Begin VB.CommandButton cmd_mensaje_4 
      Caption         =   "mensaje 4"
      Height          =   195
      Left            =   5670
      TabIndex        =   92
      Top             =   645
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton cmd_mensaje_3 
      Caption         =   "mensaje 3"
      Height          =   195
      Left            =   1785
      TabIndex        =   90
      Top             =   645
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton cmd_mensaje_2 
      Caption         =   "mensaje 2"
      Height          =   195
      Left            =   3180
      TabIndex        =   88
      Top             =   645
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton cmd_mensaje_1 
      Caption         =   "mensaje 1"
      Height          =   195
      Left            =   4050
      TabIndex        =   86
      Top             =   675
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton cmd_restructuracion 
      Caption         =   "Command2"
      Height          =   315
      Left            =   7755
      TabIndex        =   84
      Top             =   630
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Frame frm_sellos 
      Height          =   2340
      Left            =   525
      TabIndex        =   62
      Top             =   810
      Width           =   3045
      Begin VB.Frame Frame4 
         Height          =   75
         Left            =   30
         TabIndex        =   72
         Top             =   645
         Width           =   2970
      End
      Begin VB.CommandButton cmd_cerrar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   705
         Picture         =   "frmsalidas.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Cerrar Alt + C"
         Top             =   330
         Width           =   330
      End
      Begin VB.TextBox txt_sello 
         Height          =   315
         Left            =   585
         TabIndex        =   69
         Top             =   795
         Width           =   2385
      End
      Begin VB.CommandButton cmd_aceptar_sello 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   45
         Picture         =   "frmsalidas.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   330
         Width           =   330
      End
      Begin VB.CommandButton cmd_cancelar_sello 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmsalidas.frx":0B16
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   330
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_sellos 
         Height          =   1200
         Left            =   30
         TabIndex        =   66
         Top             =   1110
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   2117
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Número de Sello"
            Object.Width           =   5115
         EndProperty
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sello:"
         Height          =   195
         Left            =   90
         TabIndex        =   70
         Top             =   840
         Width           =   390
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Sellos"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   7
         Left            =   30
         TabIndex        =   64
         Top             =   120
         Width           =   2970
      End
   End
   Begin VB.Frame frm_resumen 
      Height          =   3510
      Left            =   3825
      TabIndex        =   79
      Top             =   810
      Width           =   5445
      Begin VB.TextBox txt_cantidad_total_linea 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   3765
         TabIndex        =   80
         Top             =   3060
         Width           =   1605
      End
      Begin MSComctlLib.ListView lv_resumen 
         Height          =   2910
         Left            =   30
         TabIndex        =   81
         Top             =   120
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   5133
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   3270
         TabIndex        =   82
         Top             =   3135
         Width           =   405
      End
   End
   Begin VB.CommandButton cmd_restructurar 
      Height          =   330
      Left            =   7200
      TabIndex        =   77
      Top             =   630
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame frm_movimientos 
      Height          =   1635
      Left            =   420
      TabIndex        =   73
      Top             =   975
      Width           =   6315
      Begin MSComctlLib.ListView lv_movimientos 
         Height          =   1200
         Left            =   45
         TabIndex        =   74
         Top             =   360
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   2117
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Movimiento"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Número"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "O.S"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Movimientos contenidos en el embarque"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   8
         Left            =   45
         TabIndex        =   75
         Top             =   120
         Width           =   6225
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1035
      Picture         =   "frmsalidas.frx":0C60
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Cancelar Movimiento"
      Top             =   630
      Width           =   330
   End
   Begin VB.CommandButton cmd_cerrar_embarque 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   705
      Picture         =   "frmsalidas.frx":0D62
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Cerrar Embarque"
      Top             =   630
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmsalidas.frx":0E64
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   630
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   375
      Picture         =   "frmsalidas.frx":0F66
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Buscar Movimiento"
      Top             =   630
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      Picture         =   "frmsalidas.frx":1068
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Cerrar Movimiento"
      Top             =   630
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11145
      Picture         =   "frmsalidas.frx":116A
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Salir"
      Top             =   630
      Width           =   330
   End
   Begin VB.Frame frm_cajas_faltantes 
      Height          =   2235
      Left            =   5070
      TabIndex        =   53
      Top             =   4890
      Width           =   3180
      Begin MSComctlLib.ListView lv_cajas_faltantes 
         Height          =   1830
         Left            =   45
         TabIndex        =   55
         Top             =   360
         Width           =   3090
         _ExtentX        =   5450
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
            Text            =   "     Orden Surtido"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Caja           "
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Cajas Faltantes"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   6
         Left            =   30
         TabIndex        =   54
         Top             =   120
         Width           =   3105
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   1245
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
      Left            =   630
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.TextBox txt_clave_movimiento 
      Height          =   285
      Left            =   2175
      TabIndex        =   49
      Top             =   720
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   1170
      TabIndex        =   10
      Top             =   1800
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   210
         TabIndex        =   11
         Top             =   495
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   2
         Left            =   30
         TabIndex        =   12
         Top             =   120
         Width           =   3075
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   30
      TabIndex        =   33
      Top             =   510
      Width           =   11505
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   12345
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1785
      Width           =   2100
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   45
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
            Picture         =   "frmsalidas.frx":17A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas.frx":207E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas.frx":2958
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas.frx":2EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas.frx":37D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas.frx":40AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas.frx":4984
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas.frx":4A96
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas.frx":4BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas.frx":4CBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas.frx":4DCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas.frx":4EDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas.frx":4FF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas.frx":5192
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas.frx":5FE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas.frx":61BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas.frx":62CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd_resumen 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1365
      Picture         =   "frmsalidas.frx":63DE
      Style           =   1  'Graphical
      TabIndex        =   83
      ToolTipText     =   "Resumen de Remisión"
      Top             =   630
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   15
      TabIndex        =   34
      Top             =   870
      Width           =   11505
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Index           =   0
      Left            =   60
      TabIndex        =   20
      Top             =   915
      Width           =   6975
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
         Left            =   5670
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   375
         Width           =   1140
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
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   375
         Width           =   1620
      End
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
         Left            =   750
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   375
         Width           =   1590
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Jaula:"
         Height          =   195
         Left            =   5160
         TabIndex        =   57
         Top             =   525
         Width           =   420
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   2385
         TabIndex        =   52
         Top             =   525
         Width           =   765
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   90
         TabIndex        =   51
         Top             =   525
         Width           =   600
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   1
         Left            =   30
         TabIndex        =   22
         Top             =   120
         Width           =   6900
      End
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Index           =   3
      Left            =   7110
      TabIndex        =   17
      Top             =   915
      Width           =   2220
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad a Surtir"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   4
         Left            =   30
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   420
         Width           =   1845
      End
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Index           =   4
      Left            =   9420
      TabIndex        =   14
      Top             =   915
      Width           =   2115
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad Surtida"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   5
         Left            =   30
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   420
         Width           =   1770
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2205
      Index           =   1
      Left            =   90
      TabIndex        =   0
      Top             =   1800
      Width           =   11460
      Begin VB.TextBox txt_descuento2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8910
         TabIndex        =   46
         Top             =   1800
         Width           =   1155
      End
      Begin VB.TextBox txt_descuento1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7080
         TabIndex        =   45
         Top             =   1800
         Width           =   1170
      End
      Begin VB.TextBox txt_titular 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7080
         TabIndex        =   42
         Top             =   810
         Width           =   4230
      End
      Begin VB.TextBox txt_ruta 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1395
         TabIndex        =   40
         Top             =   1470
         Width           =   4230
      End
      Begin VB.TextBox txt_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7080
         TabIndex        =   38
         Top             =   1140
         Width           =   4230
      End
      Begin VB.TextBox txt_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1395
         TabIndex        =   36
         Top             =   810
         Width           =   4230
      End
      Begin VB.TextBox txt_establecimiento 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1395
         TabIndex        =   4
         Top             =   1140
         Width           =   4230
      End
      Begin VB.TextBox txt_archivo 
         Height          =   315
         Left            =   7080
         TabIndex        =   3
         Top             =   1470
         Width           =   1170
      End
      Begin VB.TextBox txt_pedido 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1395
         TabIndex        =   2
         Top             =   1800
         Width           =   1170
      End
      Begin VB.TextBox txt_origen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1395
         TabIndex        =   1
         Top             =   480
         Width           =   4230
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   10110
         TabIndex        =   48
         Top             =   1860
         Width           =   120
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   8310
         TabIndex        =   47
         Top             =   1860
         Width           =   120
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descuentos:"
         Height          =   195
         Left            =   6075
         TabIndex        =   44
         Top             =   1860
         Width           =   900
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Left            =   6075
         TabIndex        =   43
         Top             =   870
         Width           =   480
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Left            =   195
         TabIndex        =   41
         Top             =   1530
         Width           =   390
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   6075
         TabIndex        =   39
         Top             =   1200
         Width           =   525
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   195
         TabIndex        =   37
         Top             =   870
         Width           =   555
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   30
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   510
         Width           =   660
      End
      Begin VB.Label lbl_origen 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Left            =   195
         TabIndex        =   7
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label lbl_archivo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "O. de Surtido:"
         Height          =   195
         Left            =   6075
         TabIndex        =   6
         Top             =   1530
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   1860
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3330
      Left            =   75
      TabIndex        =   23
      Top             =   3945
      Width           =   11475
      Begin VB.CommandButton cmd_pasar_movimiento 
         Height          =   330
         Left            =   8880
         Picture         =   "frmsalidas.frx":64E0
         Style           =   1  'Graphical
         TabIndex        =   78
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
         TabIndex        =   28
         Top             =   465
         Width           =   3390
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   4440
         TabIndex        =   25
         Top             =   1575
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   75
            TabIndex        =   26
            Top             =   375
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   "Cantidad a eliminar"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   27
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
         TabIndex        =   24
         Top             =   495
         Width           =   1890
      End
      Begin MSComctlLib.ListView lv_salidas 
         Height          =   2250
         Left            =   15
         TabIndex        =   29
         Top             =   1035
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   3969
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
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7585
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
            Text            =   "Surtidos    "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Paquete    "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Movimiento "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Faltan    "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Costo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "precio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Promocion 1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Promocion 2"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "tipo"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   5115
         TabIndex        =   30
         Top             =   615
         Width           =   675
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
         Left            =   5325
         TabIndex        =   76
         Top             =   420
         Width           =   6045
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   615
         Width           =   1395
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   0
         Left            =   30
         TabIndex        =   31
         Top             =   120
         Width           =   11400
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp4 
      Height          =   75
      Left            =   10185
      TabIndex        =   91
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
      Left            =   4710
      TabIndex        =   89
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
   Begin WMPLibCtl.WindowsMediaPlayer wmp2 
      Height          =   135
      Left            =   5940
      TabIndex        =   87
      Top             =   720
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
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   30
      Left            =   8490
      TabIndex        =   85
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
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   105
      TabIndex        =   35
      Top             =   0
      Width           =   11445
   End
End
Attribute VB_Name = "frmsalidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_cantidad_multibondeados As Integer
Dim var_kanban As String
Dim var_fecha_numero As Double
Dim var_nombre_articulo_mensaje As String
Dim var_nombre_tabla As String
Dim var_consecutivo As Double
Dim var_numero_pedido_cliente As Double
Dim var_origen As String
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
Dim var_orden_surtido As Double
Dim var_clave_agente As String
Dim var_clave_establecimiento As String
Dim var_clave_titular As String
Dim var_clave_cliente As String
Dim var_clave_ruta As String
Dim var_plazo As Integer
Dim var_descuento_1 As Double
Dim var_descuento_3 As Double
Dim var_descuento_2 As Double
Dim var_iva As Variant
Dim var_agrupador As String
Dim var_correo_electronico As String
Dim var_autorizo_embarque As Boolean
Dim var_es_caja As Boolean
Dim var_cajas As Boolean
Dim var_almacen_OS As String
Dim var_nota As recordSet
Dim var_movimiento_dependencia As String
Dim var_clave_moneda As String
Dim var_factura_ceros As Integer
Dim var_renglon As Double
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long


Private Sub envio_tb_transito()
    var_cadena = " SELECT dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID, dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID, dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID, dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_SALIDAS.DTIM_SAL_FECHA, dbo.TB_SALIDAS.INTE_SAL_NUMERO, dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID, dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, dbo.TB_SALIDAS.FLOA_SAL_COSTO, dbo.TB_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO,  dbo.TB_SALIDAS.VCHA_SER_SERIE_ID, dbo.TB_SALIDAS.INTE_CAR_NUMERO, dbo.TB_SALIDAS.FLOA_SAL_PROMOCION_1, dbo.TB_SALIDAS.FLOA_SAL_PROMOCION_2, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID , dbo.TB_ENCABEZADO_CARTERA.VCHA_AGE_AGENTE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ENCABEZADO_CARTERA.VCHA_ESB_ESTABLECIMIENTO_ID FROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_SALIDAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND "
    var_cadena = var_cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON  dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_SALIDAS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID WHERE (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque + ") AND (dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')"
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



Function fun_copia_archivo(Origen, Destino)
    Copy_File = CopyFile(Origen, Destino, 1)
End Function





Sub ilumina_grid()
   var_n = lv_salidas.ListItems.Count
   For var_i = 1 To var_n
       If var_i = var_renglon Then
          lv_salidas.ListItems.item(var_i).Bold = True
          lv_salidas.ListItems.item(var_i).ListSubItems(1).Bold = True
          lv_salidas.ListItems.item(var_i).ListSubItems(2).Bold = True
          lv_salidas.ListItems.item(var_i).ListSubItems(3).Bold = True
          lv_salidas.ListItems.item(var_i).ListSubItems(4).Bold = True
          lv_salidas.ListItems.item(var_i).ListSubItems(5).Bold = True
          lv_salidas.ListItems.item(var_i).ListSubItems(6).Bold = True
          lv_salidas.ListItems.item(var_i).ListSubItems(7).Bold = True
          lv_salidas.ListItems.item(var_i).ForeColor = &H8000&
          lv_salidas.ListItems.item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_salidas.ListItems.item(var_i).ListSubItems(2).ForeColor = &H8000&
          lv_salidas.ListItems.item(var_i).ListSubItems(3).ForeColor = &H8000&
          lv_salidas.ListItems.item(var_i).ListSubItems(4).ForeColor = &H8000&
          lv_salidas.ListItems.item(var_i).ListSubItems(5).ForeColor = &H8000&
          lv_salidas.ListItems.item(var_i).ListSubItems(6).ForeColor = &H8000&
          lv_salidas.ListItems.item(var_i).ListSubItems(7).ForeColor = &H8000&
       Else
          If lv_salidas.ListItems.item(var_i).ListSubItems(6) * 1 = 0 Then
             lv_salidas.ListItems.item(var_i).Bold = False
             lv_salidas.ListItems.item(var_i).ListSubItems(1).Bold = False
             lv_salidas.ListItems.item(var_i).ListSubItems(2).Bold = False
             lv_salidas.ListItems.item(var_i).ListSubItems(3).Bold = False
             lv_salidas.ListItems.item(var_i).ListSubItems(4).Bold = False
             lv_salidas.ListItems.item(var_i).ListSubItems(5).Bold = False
             lv_salidas.ListItems.item(var_i).ListSubItems(6).Bold = False
             lv_salidas.ListItems.item(var_i).ListSubItems(7).Bold = False
             lv_salidas.ListItems.item(var_i).ForeColor = &HFF&
             lv_salidas.ListItems.item(var_i).ListSubItems(1).ForeColor = &HFF&
             lv_salidas.ListItems.item(var_i).ListSubItems(2).ForeColor = &HFF&
             lv_salidas.ListItems.item(var_i).ListSubItems(3).ForeColor = &HFF&
             lv_salidas.ListItems.item(var_i).ListSubItems(4).ForeColor = &HFF&
             lv_salidas.ListItems.item(var_i).ListSubItems(5).ForeColor = &HFF&
             lv_salidas.ListItems.item(var_i).ListSubItems(6).ForeColor = &HFF&
             lv_salidas.ListItems.item(var_i).ListSubItems(7).ForeColor = &HFF&
          Else
             If var_i = var_renglon Then
                lv_salidas.ListItems.item(var_i).Bold = True
                lv_salidas.ListItems.item(var_i).ListSubItems(1).Bold = True
                lv_salidas.ListItems.item(var_i).ListSubItems(2).Bold = True
                lv_salidas.ListItems.item(var_i).ListSubItems(3).Bold = True
                lv_salidas.ListItems.item(var_i).ListSubItems(4).Bold = True
                lv_salidas.ListItems.item(var_i).ListSubItems(5).Bold = True
                lv_salidas.ListItems.item(var_i).ListSubItems(6).Bold = True
                lv_salidas.ListItems.item(var_i).ListSubItems(7).Bold = True
                lv_salidas.ListItems.item(var_i).ForeColor = &H8000&
                lv_salidas.ListItems.item(var_i).ListSubItems(1).ForeColor = &H8000&
                lv_salidas.ListItems.item(var_i).ListSubItems(2).ForeColor = &H8000&
                lv_salidas.ListItems.item(var_i).ListSubItems(3).ForeColor = &H8000&
                lv_salidas.ListItems.item(var_i).ListSubItems(4).ForeColor = &H8000&
                lv_salidas.ListItems.item(var_i).ListSubItems(5).ForeColor = &H8000&
                lv_salidas.ListItems.item(var_i).ListSubItems(6).ForeColor = &H8000&
                lv_salidas.ListItems.item(var_i).ListSubItems(7).ForeColor = &H8000&
             Else
                lv_salidas.ListItems.item(var_i).Bold = False
                lv_salidas.ListItems.item(var_i).ListSubItems(1).Bold = False
                lv_salidas.ListItems.item(var_i).ListSubItems(2).Bold = False
                lv_salidas.ListItems.item(var_i).ListSubItems(3).Bold = False
                lv_salidas.ListItems.item(var_i).ListSubItems(4).Bold = False
                lv_salidas.ListItems.item(var_i).ListSubItems(5).Bold = False
                lv_salidas.ListItems.item(var_i).ListSubItems(6).Bold = False
                lv_salidas.ListItems.item(var_i).ListSubItems(7).Bold = False
                lv_salidas.ListItems.item(var_i).ForeColor = &H80000012
                lv_salidas.ListItems.item(var_i).ListSubItems(1).ForeColor = &H80000012
                lv_salidas.ListItems.item(var_i).ListSubItems(2).ForeColor = &H80000012
                lv_salidas.ListItems.item(var_i).ListSubItems(3).ForeColor = &H80000012
                lv_salidas.ListItems.item(var_i).ListSubItems(4).ForeColor = &H80000012
                lv_salidas.ListItems.item(var_i).ListSubItems(5).ForeColor = &H80000012
                lv_salidas.ListItems.item(var_i).ListSubItems(6).ForeColor = &H80000012
                lv_salidas.ListItems.item(var_i).ListSubItems(7).ForeColor = &H80000012
             End If
       End If
   End If
   Next var_i
   If var_renglon > 0 Then
      lv_salidas.ListItems.item(var_renglon).Selected = True
      lv_salidas.selectedItem.EnsureVisible
   End If
   lv_salidas.Refresh
End Sub




Private Sub cmd_aceptar_sello_Click()
   If Trim(txt_sello) <> "" Then
      rs.Open "select * from tb_sellos where vcha_emp_empresa_id ='" + var_empresa + "' and vcha_sel_sello = '" + txt_sello + "'", cnn, adOpenDynamic, adLockOptimistic
      If rs.EOF Then
         rs.Close
         rs.Open "Insert Into tb_sellos (vcha_emp_empresa_id, inte_emb_embarque, vcha_sel_sello)  values ('" + var_empresa + "'," + Str(var_numero_embarque) + ",'" + txt_sello + "')", cnn, adOpenDynamic, adLockOptimistic
         Set list_item = lv_sellos.ListItems.Add(, , txt_sello)
      Else
         rs.Close
         MsgBox "El Sello ya existe", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de sello incorecto", vbOKOnly, "ATENCION"
   End If
   txt_sello = ""
   txt_sello.SetFocus
End Sub

Private Sub cmd_buscar_Click()
   lv_movimientos.ListItems.Clear
   Dim var_si As Integer
   If var_tipo_lectura = 0 Then
      rs.Open "SELECT dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO FROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_MOVIMIENTOS ON dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID where dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque + " and dbo.TB_DETALLE_EMBARQUES.vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
   Else
      rs.Open "SELECT dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN FROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_MOVIMIENTOS ON dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO WHERE (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque + ") AND (dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "') and vcha_aud_maquina = '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   var_si = 1
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_movimientos.ListItems.Add(, , rs!vcha_mov_nombre)
            list_item.SubItems(1) = IIf(IsNull(rs!INTE_SAL_NUMERO), 0, Trim(rs!INTE_SAL_NUMERO))
            If var_tipo_lectura = 1 Then
               list_item.SubItems(2) = IIf(IsNull(rs!inte_emo_numero_origen), 0, Trim(rs!inte_emo_numero_origen))
            End If
            rs.MoveNext
      Wend
   Else
      var_si = 0
      MsgBox "No existen movimiento hechos en esta maquina", vbOKOnly, "ATENCION"
   End If
   rs.Close
   var_n = lv_movimientos.ListItems.Count
   
   If var_n > 4 Then
      lv_movimientos.ColumnHeaders(1).Width = 2950.74
   Else
      lv_movimientos.ColumnHeaders(1).Width = 3199.74
   End If
   If var_si = 1 Then
      frm_movimientos.Visible = True
      lv_movimientos.SetFocus
   End If
End Sub

Private Sub cmd_cancelar_Click()
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_ENTRADAS_I = New TB_ENTRADAS_I
   Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_ENTRADAS_VISTAS_I = New TB_ENTRADAS_VISTAS_I
   Set TB_DET_ORDEN_SURTIDO_M = New TB_DET_ORDEN_SURTIDO_M
   rs.Open "select * from tb_encabezado_embarques where inte_emb_embarque = " + Str(var_numero_embarque), cnn, adOpenDynamic, adLockOptimistic
   var_embarque_agente = rs!VCHA_AGE_AGENTE_ID
   var_embarque_cerrado = Trim(rs!CHAR_EMB_ESTATUS)
   rs.Close
   If var_embarque_cerrado = "F" Then
      MsgBox "El embarque ya fue facturado y es imposible cancelarlo", vbOKOnly, "ATENCION"
   Else
      If var_embarque_cerrado = "I" Then
         If var_numero_folio > 0 Then
            If var_estatus_movimiento = "C" Then
               MsgBox "El Movimiento ya fue cancelado", vbOKOnly, "ATENCION"
            Else
               If var_estatus_movimiento = "I" Then
                  If var_fecha_movimiento <> Date Then
                     var_posible_accion = False
                     frmsupervisor1.Show 1
                     If var_posible_accion = True Then
                        si = MsgBox("¿Desea cancelar el movimiento?", vbYesNo, "ATENCION")
                        If si = 6 Then
                           si = MsgBox("Confirmar la cancelación del movimiento", vbYesNo, "ATENCION")
                           If si = 6 Then
                              Set TB_ENC_MOV_CANCELACION = New TB_ENC_MOV_CANCELACION
                              var_actualizar = False
                              rs.Open "SELECT * FROM TB_ENCABEZADO_MOVIMIENTOS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO =  " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                              var_tipo_Cambio = rs!floa_emo_tipo_cambio
                              rs.Close
                              var_actualizar = TB_ENC_MOV_CANCELACION.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, "C", var_global_supervisor_1, var_global_supervisor_2)
                              rs.Open "SELECT * FROM TB_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_SAL_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                              While Not rs.EOF
                                    var_actualiza = TB_DET_ORDEN_SURTIDO_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_orden_surtido, rs!VCHA_ART_ARTICULO_ID, 0 - rs!floa_Sal_Cantidad, rs!floa_Sal_Cantidad, rs!floa_Sal_precio / var_tipo_Cambio, rs!char_ped_tipo)
                                    rs.MoveNext
                              Wend
                              rs.Close
                              rs.Open "update tb_detalle_cajas set char_paq_estatus = 'I' where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_PAQ_MOVIMIENTO_DESTINO = '" + var_clave_movimiento + "' and INTE_PAQ_NUMERO_DESTINO = " + CStr(var_numero_folio) + " And char_paq_estatus = 'S'", cnn, adOpenDynamic, adLockOptimistic
                              lbl_cancelado = "MOVIMIENTO CANCELADO"
                              Me.cmd_imprimir.Enabled = False
                              Me.cmd_cancelar.Enabled = False
                              MsgBox "El movimiento a sido cancelado", vbOKOnly, "ATENCION"
                              var_estatus_movimiento = "C"
                           End If
                        End If
                     End If
                  Else
                     MsgBox "El movimiento ya no puede ser cancelado ya que no pertence al dia", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "El Movimiento no a sido cerrado aun", vbOKOnly, "ATENCION"
               End If
            End If
         Else
            MsgBox "No se a seleccionado un movimiento", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El embarque no a sido cerrado aun", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub cmd_cancelar_sello_Click()
   frm_sellos.Visible = False
End Sub

Private Sub cmd_cerrar_Click()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   Set TB_ENTRADAS_I = New TB_ENTRADAS_I
   
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
   Dim var_precio_ft As Double
   Dim var_si_multibondeados As Boolean
   Dim var_posible_cerrar_KANBAN As Boolean
   var_si_multibondeados = True
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
   If var_empresa = "16" Then
      var_mes = Month(Date)
      var_año = Year(Date)
      If var_mes = 1 Then
         var_mes = 12
         var_año = var_año - 1
      Else
         var_mes = var_mes - 1
      End If
      rsaux9.Open "select * from TB_IMPORTE_PAGOS_MULTIBONDEADOS where INTE_DMU_AÑO = " + CStr(var_año) + " and INTE_DMU_MES = " + CStr(var_mes), cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux9.EOF Then
         var_si_multibondeados = True
      Else
         var_si_multibondeados = False
      End If
      rsaux9.Close
   Else
      var_si_multibondeados = True
   End If
   If var_si_multibondeados = True Then
      If var_empresa = "31" Then
         var_almacen_traspaso_cantia = "CC_1"
         'Me.txt_folio = 61
         var_numero_traspaso_cantia = CDbl(Me.txt_folio)
         frmpassword_traspasos_cantia.lbl_movimiento = var_clave_movimiento
         frmpassword_traspasos_cantia.Show 1
         If var_acepta_traspaso_global = 0 Then
            var_acepta_traspaso = 0
         Else
            var_acepta_traspaso = 1
         End If
      Else
         var_acepta_traspaso = 1
      End If
      VAR_CADENA_ORDENES_NO_SUBIDAS = ""
      rsaux9.Open "SELECT * FROM TB_EMBARQUES_ORDENES_SURTIDO WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
      While Not rsaux9.EOF
            var_cadena = "SELECT dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN, dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID, dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_DETALLE_EMBARQUES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO WHERE  (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN = " + CStr(rsaux9!INTE_ORS_ORDEN_SURTIDO) + ")"
            'var_cadena = "SELECT     dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID, dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_DETALLE_EMBARQUES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO WHERE (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN = " + rsaux9!inte_ors_orden_surtido + ")"
            'MsgBox var_cadena
            rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            If rsaux10.EOF Then
               rsaux11.Open "SELECT dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ENC_ORDEN_SURTIDO.VCHA_EMP_EMPRESA_ID FROM dbo.TB_ENC_ORDEN_SURTIDO INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENC_ORDEN_SURTIDO.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID Where (dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = " + CStr(rsaux9!INTE_ORS_ORDEN_SURTIDO) + ")", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux11.EOF Then
                  If Trim(VAR_CADENA_ORDENES_NO_SUBIDAS) = "" Then
                     VAR_CADENA_ORDENES_NO_SUBIDAS = CStr(rsaux9!INTE_ORS_ORDEN_SURTIDO) + " del cliente " + IIf(IsNull(rsaux11!VCHA_CLI_NOMBRE), "", rsaux11!VCHA_CLI_NOMBRE)
                  Else
                     VAR_CADENA_ORDENES_NO_SUBIDAS = VAR_CADENA_ORDENES_NO_SUBIDAS + ", " + CStr(rsaux9!INTE_ORS_ORDEN_SURTIDO) + " del cliente " + IIf(IsNull(rsaux11!VCHA_CLI_NOMBRE), "", rsaux11!VCHA_CLI_NOMBRE)
                  End If
               Else
                  If Trim(VAR_CADENA_ORDENES_NO_SUBIDAS) = "" Then
                     VAR_CADENA_ORDENES_NO_SUBIDAS = CStr(rsaux9!INTE_ORS_ORDEN_SURTIDO)
                  Else
                     VAR_CADENA_ORDENES_NO_SUBIDAS = VAR_CADENA_ORDENES_NO_SUBIDAS + ", " + CStr(rsaux9!INTE_ORS_ORDEN_SURTIDO)
                  End If
               End If
               rsaux11.Close
            End If
            rsaux10.Close
            rsaux9.MoveNext
      Wend
      rsaux9.Close
      If Trim(VAR_CADENA_ORDENES_NO_SUBIDAS) <> "" Then
         MsgBox "El embarque no puede ser cerrado ya que no se encuentra cargadas las ordenes de surtido " + VAR_CADENA_ORDENES_NO_SUBIDAS, vbOKOnly, "ATENCION"
      Else
         If var_acepta_traspaso = 1 Then
            rs.Open "select * from tb_encabezado_embarques where inte_emb_embarque = " + Str(var_numero_embarque) + " and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_embarque_cerrado = IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", Trim(rs!CHAR_EMB_ESTATUS))
            End If
            rs.Close
            si = MsgBox("¿Esta seguro que desea cerrar el embarque?", vbYesNo, "ATENCION")
            If si = 6 Then
               si = MsgBox("Confirmar el cerrado del embarque", vbOKCancel, "ATENCION")
               If si = 1 Then
                  If Trim(var_embarque_cerrado) = "" Then
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
                     rsaux3.Open "select * from vw_embarques_cerrar where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_emb_embarque = " + txt_embarque + " and char_emb_estatus = ''", cnn, adOpenDynamic, adLockOptimistic
                     ''' aqui se debe de poner la validacion del kanban
                     var_posible_cerrar_KANBAN = True
                     If var_posible_kanban = 1 Then
                         Set TB_PROC_KANBANS_EN_MOVIMIENTO = New TB_PROC_KANBANS_EN_MOVIMIENTO
                         While Not rsaux3.EOF
                               var_inserta = TB_PROC_KANBANS_EN_MOVIMIENTO.Anadir(rsaux3!VCHA_ALM_ALMACEN_ID, var_clave_movimiento, rsaux3!INTE_SAL_NUMERO, "", "")
                               If var_kanban_exito = "N" Then
                                  var_posible_cerrar_KANBAN = False
                               End If
                               rsaux3.MoveNext
                         Wend
                     Else
                        var_posible_cerrar_KANBAN = True
                     End If
                     ''' finaliza la validacion del kanban
                     If rsaux3.RecordCount > 0 Then
                        rsaux3.MoveFirst
                     End If
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
                     If var_posible_cerrar_KANBAN = True Then
                        If var_posible_tipo_cambio = True Then
                           var_numero_folio_anterior = var_numero_folio
                           If rsaux3.RecordCount > 0 Then
                              rsaux3.MoveFirst
                           End If
                           var_fecha_inicio = CStr(Now)
                           var_clave_movimiento_anterior = var_clave_movimiento
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
                                 If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
                                 Else
                                    If var_tipo_Cambio > 0 Then
                                       If var_fecha_surtido_catalogo <= Date Then
                                          var_si_surtir_catalogo = 1
                                       Else
                                          var_si_surtir_catalogo = 0
                                       End If
                                       If var_tipo_lectura = 0 Then
                                          var_nombre_tabla = "TEMP_" + Trim(var_empresa) + Trim(var_unidad_organizacional) + Trim(var_almacen_origen) + Trim(var_clave_movimiento) + Trim(CStr(var_numero_folio))
                                          rsaux4.Open "select * from dbo.sysobjects where id = object_id(N'[dbo].[" + var_nombre_tabla + "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1", cnn, adOpenDynamic, adLockOptimistic
                                          If rsaux4.EOF Then
                                             rsaux4.Close
                                             Cadena = "CREATE TABLE [dbo].[" + var_nombre_tabla + "] ([VCHA_EMP_EMPRESA_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_UOR_UNIDAD_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_ALM_ALMACEN_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_MOV_MOVIMIENTO_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[INTE_SAL_NUMERO] [int] NULL ,[VCHA_ART_ARTICULO_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[FLOA_SAL_CANTIDAD] [float] NULL ,[FLOA_SAL_COSTO] [float] NULL ,[FLOA_SAL_PRECIO] [float] NULL ,[FLOA_SAL_DESCUENTO] [float] NULL ,[FLOA_SAL_PROMOCION_1] [float] NULL ,[FLOA_SAL_PROMOCION_2] [float] NULL ,[VCHA_REE_FOLIO] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_SAL_REFERENCIA] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[CHAR_PED_TIPO] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_CAT_CATALOGO_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[FLOA_SAL_DESCUENTO_1] [float] NULL ,"
                                             Cadena = Cadena + " [FLOA_SAL_DESCUENTO_2] [float] NULL ,[INTE_SAL_AÑO] [int] NULL , [INTE_SAL_CONSECUTIVO] [int] NULL) ON [PRIMARY]"
                                             rsaux4.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                             rsaux4.Open "INSERT INTO " + var_nombre_tabla + " (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, CHAR_PED_TIPO, INTE_SAL_CONSECUTIVO) select VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, CHAR_PED_TIPO, INTE_SAL_CONSECUTIVO from tb_temporal_salidas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id ='" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                                          Else
                                             rsaux4.Close
                                          End If
                                          rsaux4.Open "delete from  tb_temporal_salidas where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                                          rsaux4.Open "INSERT INTO tb_temporal_salidas (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, CHAR_PED_TIPO, INTE_SAL_CONSECUTIVO) select VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, CHAR_PED_TIPO, INTE_SAL_CONSECUTIVO from " + var_nombre_tabla, cnn, adOpenDynamic, adLockOptimistic
                                          Text1 = "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", " + CStr(var_tipo_Cambio) + ",'" + var_catalogo_1 + "','" + var_catalogo_2 + "','" + var_clave_titular + "','" + var_clave_cliente + "'," + CStr(var_año_catalogo) + "," + CStr(var_mes_catalogo) + "," + CStr(var_si_surtir_catalogo)
                                          If var_clave_movimiento = "FT" Then
                                             rsaux4.Open "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES_TIENDAS '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", " + CStr(var_tipo_Cambio) + ",'" + var_catalogo_1 + "','" + var_catalogo_2 + "','" + var_clave_titular + "','" + var_clave_cliente + "'," + CStr(var_año_catalogo) + "," + CStr(var_mes_catalogo) + "," + CStr(var_si_surtir_catalogo), cnn, adOpenDynamic, adLockOptimistic
                                          Else
                                             rsaux4.Open "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", " + CStr(var_tipo_Cambio) + ",'" + var_catalogo_1 + "','" + var_catalogo_2 + "','" + var_clave_titular + "','" + var_clave_cliente + "'," + CStr(var_año_catalogo) + "," + CStr(var_mes_catalogo) + "," + CStr(var_si_surtir_catalogo), cnn, adOpenDynamic, adLockOptimistic
                                          End If
                                       Else
                                          var_archivo_tabla = Trim(var_empresa) + Trim(var_unidad_organizacional) + Trim(var_almacen_origen) + Trim(var_clave_movimiento) + Trim(CStr(var_numero_folio))
                                          rsaux5.Open "select * from tb_salidas where vcha_sal_archivo = '" + var_archivo_tabla + "' and floa_sal_cantidad > 0", cnnaccess, adOpenDynamic, adLockOptimistic
                                          'MsgBox "DELETE FROM TB_TEMPORAL_SALIDAS WHERE vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id =  '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero  =  " + CStr(var_numero_folio)
                                          rsaux9.Open "DELETE FROM TB_TEMPORAL_SALIDAS WHERE vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id =  '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero  =  " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                                          
                                          cnnaccess.BeginTrans
                                          While Not rsaux5.EOF
                                             rsaux2.Open "INSERT INTO tb_temporal_salidas (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, CHAR_PED_TIPO, INTE_SAL_CONSECUTIVO) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + rsaux5!VCHA_ART_ARTICULO_ID + "', " + CStr(rsaux5!floa_Sal_Cantidad) + ", " + CStr(rsaux5!floa_Sal_costo) + ", " + CStr(rsaux5!floa_Sal_precio) + ", 0, " + CStr(rsaux5!floa_sal_promocion_1) + ", " + CStr(rsaux5!FLOA_SAL_PROMOCION_2) + ", '" + rsaux5!vcha_sal_tipo + "', " + CStr(rsaux5!INTE_SAL_CONSECUTIVO) + ")", cnn, adOpenDynamic, adLockOptimistic
                                             rsaux4.Open "UPDATE TB_detalle_pedidos SET FLOA_ped_CANTIDAD_SURTIDA = FLOA_ped_CANTIDAD_SURTIDA + " + CStr(rsaux5!floa_Sal_Cantidad) + " WHERE INTE_ped_numero = " + CStr(rsaux5!inte_ped_numero) + " AND VCHA_ART_ARTICULO_ID = '" + rsaux5!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                             rsaux5.MoveNext
                                          Wend
                                          rsaux5.Close
                                          'rsaux5.Open "DELETE FROM TB_SALIDAS WHERE VCHA_SAL_ARCHIVO = '" + var_archivo_tabla + "'", cnnaccess, adOpenDynamic, adLockOptimistic
                                          cnnaccess.CommitTrans
                                          If var_empresa = "16" Then
                                             rsaux4.Open "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES_MULTIBONDEADOS '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", " + CStr(var_tipo_Cambio) + ",'" + var_catalogo_1 + "','" + var_catalogo_2 + "','" + var_clave_titular + "','" + var_clave_cliente + "'," + CStr(var_año_catalogo) + "," + CStr(var_mes_catalogo) + "," + CStr(var_si_surtir_catalogo), cnn, adOpenDynamic, adLockOptimistic
                                          Else
                                             If var_clave_movimiento = "FT" Then
                                                'MsgBox "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES_TIENDAS '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", " + CStr(var_tipo_Cambio) + ",'" + var_catalogo_1 + "','" + var_catalogo_2 + "','" + var_clave_titular + "','" + var_clave_cliente + "'," + CStr(var_año_catalogo) + "," + CStr(var_mes_catalogo) + "," + CStr(var_si_surtir_catalogo)
                                                'cnn.CommitTrans
                                                rsaux4.Open "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES_TIENDAS '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", " + CStr(var_tipo_Cambio) + ",'" + var_catalogo_1 + "','" + var_catalogo_2 + "','" + var_clave_titular + "','" + var_clave_cliente + "'," + CStr(var_año_catalogo) + "," + CStr(var_mes_catalogo) + "," + CStr(var_si_surtir_catalogo), cnn, adOpenDynamic, adLockOptimistic
                                             Else
                                                If parametros(0) = "SQLHOUSTON" Then
                                                   rsaux4.Open "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", " + CStr(var_tipo_Cambio) + ",'" + var_catalogo_1 + "','" + var_catalogo_2 + "','" + var_clave_titular + "','" + var_clave_cliente + "'," + CStr(var_año_catalogo) + "," + CStr(var_mes_catalogo) + "," + CStr(var_si_surtir_catalogo), cnn, adOpenDynamic, adLockOptimistic
                                                Else
                                                   rsaux4.Open "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", " + CStr(var_tipo_Cambio) + ",'" + var_catalogo_1 + "','" + var_catalogo_2 + "','" + var_clave_titular + "','" + var_clave_cliente + "'," + CStr(var_año_catalogo) + "," + CStr(var_mes_catalogo) + "," + CStr(var_si_surtir_catalogo), cnn, adOpenDynamic, adLockOptimistic
                                                End If
                                             End If
                                          End If
                                          Dim var_ttt As Integer
                                          var_ttt = 0
                                          'inicio traspaso tienda cantia
                                          If var_empresa = "310000" Then
                                             'If rsaux4.State = 1 Then
                                             '   rsaux4.Close
                                             'End If
                                             If var_ttt = 0 Then
                                                ok = False
                                                ok = TB_ENC_EMBARQUE_M.Anadir(var_empresa, var_unidad_organizacional, var_numero_embarque, "I")
                                                var_ttt = 1
                                             End If
                                             rsaux4.Open "select * from tb_detalle_Embarques where vcha_Emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque
                                             While Not rsaux4.EOF
                                                   var_primera_vez = True
                                                   rsaux5.Open "select * from tb_Salidas where vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + rsaux4!VCHA_MOV_MOVIMIENTO_ID + "' and inte_sal_numero =  " + CStr(rsaux4!INTE_SAL_NUMERO), cnn, adOpenDynamic, adLockOptimistic
                                                   While Not rsaux5.EOF
                                                         If var_primera_vez = True Then
                                                            var_inserta = False
                                                            var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, "CC_1", "ETC", Now, 0, 0, "", "", "", "CC_1", "", var_clave_usuario_global, fun_NombrePc, 0, "", "NUMERO EMBARQUE " + Me.txt_embarque, "", "B", "", "", 0, 0, 0, var_clave_moneda, 0)
                                                            var_numero_folio_CANTIA = var_numero_folio_regreso
                                                            var_primera_vez = False
                                                         End If
                                                         var_cantidad_leida = IIf(IsNull(rsaux5!floa_Sal_Cantidad), 0, rsaux5!floa_Sal_Cantidad)
                                                         var_año = 2005
                                                         var_costo = IIf(IsNull(rsaux5!floa_Sal_costo), 0, rsaux5!floa_Sal_costo)
                                                         var_precio = IIf(IsNull(rsaux5!floa_Sal_precio), 0, rsaux5!floa_Sal_precio)
                                                         Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = 'CC_1' and  VCHA_MOV_MOVIMIENTO_ID = 'ETC' and inte_ent_numero = " + Str(var_numero_folio_CANTIA) + " and vcha_art_articulo_id = '" + rsaux5!VCHA_ART_ARTICULO_ID + "'"
                                                         rsaux6.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                                         If Not rsaux6.EOF Then
                                                            var_inserta = False
                                                            var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, "CC_1", "ETC", CDbl(var_numero_folio_CANTIA), rsaux5!VCHA_ART_ARTICULO_ID, CDbl(var_cantidad_leida), CDbl(var_año))
                                                         Else
                                                            var_inserta = False
                                                            var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, "CC_1", "ETC", CDbl(var_numero_folio_CANTIA), rsaux5!VCHA_ART_ARTICULO_ID, CDbl(var_cantidad_leida), CDbl(var_costo), CDbl(var_precio), "0", "", CDbl(var_año))
                                                         End If
                                                         rsaux6.Close
                                                         rsaux5.MoveNext
                                                   Wend
                                                   rsaux5.Close
                                                   
                                                   Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = 'CC_1' and  VCHA_MOV_MOVIMIENTO_ID = 'ETC' and inte_ent_numero = " + Str(var_numero_folio_CANTIA)
                                                   rsaux5.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                                   While Not rsaux5.EOF
                                                         var_inserta = False
                                                         rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO) values ('" + rsaux5!VCHA_EMP_EMPRESA_ID + "', '" + rsaux5!VCHA_UOR_UNIDAD_ID + "', '" + rsaux5!VCHA_ALM_ALMACEN_ID + "', '" + rsaux5!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rsaux5!inte_ent_numero) + ", '" + rsaux5!VCHA_ART_ARTICULO_ID + "', " + CStr(rsaux5!floa_ent_cantidaD) + ", " + CStr(rsaux5!floa_ent_costo) + " , " + CStr(rsaux5!floa_ent_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
                                                         rsaux5.MoveNext
                                                   Wend
                                                   rsaux5.Close
                                                   var_estatus_movimiento = "I"
                                                   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, "CC_1", "ETC", CDbl(var_numero_folio_CANTIA), "", Now, 1)
                                                   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, "CC_1", "ETC", CDbl(var_numero_folio_CANTIA), "I", Now, 1)
                                                   
                                                   
                                                   Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_ENTRADAS.rpt")
                                                   reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_ENTRADA.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_MOVIMIENTOS_ENTRADA.VCHA_MOV_MOVIMIENTO_ID} = 'ETC' AND {VW_MOVIMIENTOS_ENTRADA.INTE_EMO_NUMERO} = " + Str(var_numero_folio_CANTIA) + " AND {VW_MOVIMIENTOS_ENTRADA.VCHA_ALM_ALMACEN_ID} = 'CC_1' and {VW_MOVIMIENTOS_ENTRADA.VCHA_eMP_EMPRESA_ID} = '" + var_empresa + "'"
                                                   frmvistasprevias.cr.ReportSource = reporte
                                                   For ntablas = 1 To reporte.Database.Tables.Count
                                                       reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                                   Next ntablas
                                                   frmvistasprevias.cr.ViewReport
                                                   frmvistasprevias.Caption = "Reporte de Movimientos"
                                                   frmvistasprevias.Show 1
                                                   Set reporte = Nothing
                                                   MsgBox "Se genero el movimiento de ENTRADA TIENDA CANTIA número " + CStr(var_numero_folio_CANTIA), vbOKOnly, "ATENCION"
                                                   rsaux4.MoveNext
                                             Wend
                                             rsaux4.Close
                                             
                                          End If
                                          ' fin de entrada tienda cantia
                                       
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux3.MoveNext
                           Wend
                           
                           
                           
                           
                           
                           
                           
                           rsaux3.Close
                        
                           var_cadena = "SELECT dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN, dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID , dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_SALIDAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO INNER JOIN dbo.TB_DETALLE_EMBARQUES ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID AND "
                           var_cadena = var_cadena + " dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID AND dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_SALIDAS.INTE_SAL_NUMERO = dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO WHERE (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque + ") "
                           rsaux9.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                           While Not rsaux9.EOF
                                 rsaux11.Open "update tb_det_orden_surtido set floa_ors_cantidad_SALIDA =  floa_ors_cantidad_salida + " + CStr(rsaux9!floa_Sal_Cantidad) + " where inte_ors_orden_surtido  = " + CStr(rsaux9!inte_emo_numero_origen) + " and vcha_Art_articulo_id = '" + rsaux9!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                 rsaux9.MoveNext
                           Wend
                           rsaux9.Close
                        
                           var_fecha_fin = CStr(Now)
                           Me.txt_codigo.Enabled = False
                           ok = False
                           ok = TB_ENC_EMBARQUE_M.Anadir(var_empresa, var_unidad_organizacional, var_numero_embarque, "I")
                           'var_si = MsgBox("¿Desea cerrar los pedidos del embarque?", vbYesNo, "ATENCION")
                           var_si = 6
                           If var_si = 6 Then
                              rsaux8.Open "SELECT * FROM VW_EMBARQUES_PEDIDOS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                              While Not rsaux8.EOF
                                    rsaux.Open "update tb_encabezado_pedidos set CHAR_PED_ESTATUS = 'E' where inte_ped_numero = " + CStr(rsaux8!inte_ped_numero), cnn, adOpenDynamic, adLockOptimistic
                                    rsaux8.MoveNext
                              Wend
                              rsaux8.Close
                           End If
                           var_estatus_movimiento = "I"
                           var_numero_folio = var_numero_folio_anterior
                           var_embarque_cerrado = "I"
                           rsaux5.Open "update tb_encabezado_embarques set dtim_emb_fecha_final = getdate() where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
                        
                           If var_empresa = "31" Then
                              On Error GoTo salir:
                           End If
                           
                           If var_empresa = "31" And var_clave_movimiento = "ET" Then
                              rsaux8.Open "select * from tb_detalle_Embarques where vcha_Emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque
                              While Not rsaux8.EOF
                                    var_primera_vez = True
                                    rsaux5.Open "select * from tb_Salidas where vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + rsaux8!VCHA_MOV_MOVIMIENTO_ID + "' and inte_sal_numero =  " + CStr(rsaux8!INTE_SAL_NUMERO), cnn, adOpenDynamic, adLockOptimistic
                                    While Not rsaux5.EOF
                                          If var_primera_vez = True Then
                                             var_inserta = False
                                             var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, "CC_1", "ETC", Now, 0, 0, "", "", "", "CC_1", "", var_clave_usuario_global, fun_NombrePc, 0, "", "NUMERO EMBARQUE " + Me.txt_embarque, "", "B", "", "", 0, 0, 0, var_clave_moneda, 0)
                                             var_numero_folio_CANTIA = var_numero_folio_regreso
                                             var_primera_vez = False
                                          End If
                                          var_cantidad_leida = IIf(IsNull(rsaux5!floa_Sal_Cantidad), 0, rsaux5!floa_Sal_Cantidad)
                                          var_año = 2005
                                          var_costo = IIf(IsNull(rsaux5!floa_Sal_costo), 0, rsaux5!floa_Sal_costo)
                                          var_precio = IIf(IsNull(rsaux5!floa_Sal_precio), 0, rsaux5!floa_Sal_precio)
                                          Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = 'CC_1' and  VCHA_MOV_MOVIMIENTO_ID = 'ETC' and inte_ent_numero = " + Str(var_numero_folio_CANTIA) + " and vcha_art_articulo_id = '" + rsaux5!VCHA_ART_ARTICULO_ID + "'"
                                          rsaux6.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux6.EOF Then
                                             var_inserta = False
                                             var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, "CC_1", "ETC", CDbl(var_numero_folio_CANTIA), rsaux5!VCHA_ART_ARTICULO_ID, CDbl(var_cantidad_leida), CDbl(var_año))
                                          Else
                                             var_inserta = False
                                             var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, "CC_1", "ETC", CDbl(var_numero_folio_CANTIA), rsaux5!VCHA_ART_ARTICULO_ID, CDbl(var_cantidad_leida), CDbl(var_costo), CDbl(var_precio), "0", "", CDbl(var_año))
                                          End If
                                          rsaux6.Close
                                          rsaux5.MoveNext
                                    Wend
                                    rsaux5.Close
                                    Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = 'CC_1' and  VCHA_MOV_MOVIMIENTO_ID = 'ETC' and inte_ent_numero = " + Str(var_numero_folio_CANTIA)
                                    rsaux5.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                    While Not rsaux5.EOF
                                          var_inserta = False
                                          rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO) values ('" + rsaux5!VCHA_EMP_EMPRESA_ID + "', '" + rsaux5!VCHA_UOR_UNIDAD_ID + "', '" + rsaux5!VCHA_ALM_ALMACEN_ID + "', '" + rsaux5!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rsaux5!inte_ent_numero) + ", '" + rsaux5!VCHA_ART_ARTICULO_ID + "', " + CStr(rsaux5!floa_ent_cantidaD) + ", " + CStr(rsaux5!floa_ent_costo) + " , " + CStr(rsaux5!floa_ent_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
                                          rsaux5.MoveNext
                                    Wend
                                    rsaux5.Close
                                    var_estatus_movimiento = "I"
                                    var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, "CC_1", "ETC", CDbl(var_numero_folio_CANTIA), "", Now, 1)
                                    var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, "CC_1", "ETC", CDbl(var_numero_folio_CANTIA), "I", Now, 1)
                                                  
                                                   
                                    Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_ENTRADAS.rpt")
                                    reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_ENTRADA.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_MOVIMIENTOS_ENTRADA.VCHA_MOV_MOVIMIENTO_ID} = 'ETC' AND {VW_MOVIMIENTOS_ENTRADA.INTE_EMO_NUMERO} = " + Str(var_numero_folio_CANTIA) + " AND {VW_MOVIMIENTOS_ENTRADA.VCHA_ALM_ALMACEN_ID} = 'CC_1' and {VW_MOVIMIENTOS_ENTRADA.VCHA_eMP_EMPRESA_ID} = '" + var_empresa + "'"
                                    frmvistasprevias.cr.ReportSource = reporte
                                    For ntablas = 1 To reporte.Database.Tables.Count
                                        reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                    Next ntablas
                                    frmvistasprevias.cr.ViewReport
                                    frmvistasprevias.Caption = "Reporte de Movimientos"
                                    frmvistasprevias.Show 1
                                    Set reporte = Nothing
                                    rsaux8.MoveNext
                              Wend
                              rsaux8.Close
                           End If
                           MsgBox var_fecha_inicio + " " + CStr(var_fecha_fin), vbOKOnly, ""
                           If var_clave_movimiento = "ET" Then
                              Call envio_tb_transito
                           End If
                           MsgBox "Se a cerrado el embarque", vbOKOnly, "ATENCION"
                        Else
                           rsaux3.Close
                           MsgBox "No es posible cerrar el embarque ya que no se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
                        End If
                     Else
                        MsgBox "Se a generado un error al procesar los kanban", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "El embaruqe ya habia sido cerrado con anterioridad", vbOKOnly, "ATENCION"
                  End If
                  var_clave_movimiento = var_clave_movimiento_anterior
               Else
                  MsgBox "El cerrado del embarque a sido cancelado", vbOKOnly, "ATENCION"
               End If
            End If
         Else
            MsgBox "No se puede cerrar el traspaso hasta que haya una autorización del almacén de tienda", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "No se a calculado los pagos del mes anterior", vbOKOnly, "ATENCION"
   End If
   frm_sellos.Visible = False
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
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

   frm_sellos.Visible = False
   Exit Sub
salir:
   If rs.State = 1 Then
      rs.Close
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

End Sub

Private Sub cmd_cerrar_embarque_Click()
   Dim var_busqueda_folio As Double
   Dim var_busqueda_numero As Double
   Dim var_busqueda_referencia As String
   Dim var_posible As Boolean
   Dim var_existen_cajas As Integer
   Dim var_numero_items As Integer
   lv_sellos.ListItems.Clear
   txt_sello = ""
   Cadena = "SELECT dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO, dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR, dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIDA , IsNull(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_NEGADA, 0) AS FLOA_ORS_CANTIDAD_NEGADA, dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO , dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_DETALLE_EMBARQUES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID AND"
   Cadena = Cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO INNER JOIN dbo.TB_DET_ORDEN_SURTIDO ON dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN = dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO AND dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_DET_ORDEN_SURTIDO.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_DET_ORDEN_SURTIDO.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_DET_ORDEN_SURTIDO.VCHA_ALM_ALMACEN_ID WHERE (dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIDA + ISNULL(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_NEGADA, 0) < dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR)"
   Cadena = Cadena + "  and dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + txt_embarque + " AND dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'"
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   rsaux4.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux4.EOF Then
      rsaux4.Close
      var_negado_desde = 1
      frmasignacion_negado.txt_numero_embarque = Me.txt_embarque
      frmasignacion_negado.txt_agente = Me.txt_agente
      var_activa_forma_asignacion_negado = Me.Name
      frmasignacion_negado.Show 1
      If Trim(txt_folio) <> "" Then
         txt_busqueda_folio = Me.txt_folio
         rs.Open "select * from tb_detalle_embarques where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + txt_busqueda_folio + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' AND INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If var_numero_embarque = rs!inte_emb_embarque Then
               var_posible = True
            Else
               MsgBox "Este movimiento se encuentra en el embarque No. " + Str(rs!inte_emb_embarque), vbOKOnly, "ATENCION"
               var_posible = False
               frm_movimientos.Visible = False
            End If
         Else
            MsgBox "El Movimiento no existe", vbOKOnly, "ATENCION"
            var_posible = False
            frm_movimientos.Visible = False
         End If
         rs.Close
         If var_posible = True Then
            rs.Open "select * from tb_detalle_cajas with (nolock) where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_paq_movimiento_destino = '" + var_clave_movimiento + "' and inte_paq_numero_destino = " + txt_busqueda_folio, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_cajas = True
            Else
               var_cajas = False
            End If
            rs.Close
            If var_numero_folio = CDbl(txt_busqueda_folio) Then
               rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            rs.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + txt_busqueda_folio + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               If var_numero_folio > 0 Then
                  rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
               End If
               var_movimiento_bloqueado = IIf(IsNull(rs!INTE_EMO_BLOQUEADO), 0, rs!INTE_EMO_BLOQUEADO)
               If var_movimiento_bloqueado = 0 Then
                  var_primera_vez = False
                  var_factura_ceros = IIf(IsNull(rs!inte_emo_factura_ceros), 0, rs!inte_emo_factura_ceros)
                  var_clave_moneda = rs!vcha_mon_moneda_id
                  var_orden_surtido = rs!inte_emo_numero_origen
                  var_numero_folio = rs!INTE_EMO_NUMERO
                  txt_folio = rs!INTE_EMO_NUMERO
                  var_almacen_origen = rs!VCHA_ALM_ALMACEN_ID
                  rsaux3.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_origen = rsaux3!VCHA_ALM_NOMBRE
                  rsaux3.Close
                  If IsNull(rs!char_Emo_estatus) Then
                     var_estatus_movimiento = ""
                  Else
                     var_estatus_movimiento = rs!char_Emo_estatus
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
                  rs.Close
                  rs.Open "select * from vw_orden_surtido where inte_ors_orden_surtido = " + Str(var_orden_surtido) + " and floa_ors_cantidad_surtir > 0", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     txt_archivo = var_orden_surtido
                     var_suma_cantidad_enviada = 0
                     var_suma_cantidad_recibida = 0
                     lbl_enviados.Caption = "0"
                     lbl_recibidos.Caption = "0"
                     lv_salidas.ListItems.Clear
                     If IsNull(rs!VCHA_TIT_NOMBRE) Then
                        GoTo no_titular:
                     Else
                        txt_titular = rs!VCHA_TIT_NOMBRE
                        var_clave_titular = rs!vcha_tit_titular_id
                     End If
                     If IsNull(rs!inte_ped_dias_condiciones) Then
                        var_plazo = 0
                     Else
                        var_plazo = rs!inte_ped_dias_condiciones
                     End If
                     If IsNull(rs!vcha_cli_email) Then
                        var_correo_electronico = ""
                     Else
                        var_correo_electronico = rs!vcha_cli_email
                     End If
                     If IsNull(rs!VCHA_ESB_NOMBRE) Then
                        GoTo no_establecimiento:
                     Else
                        txt_establecimiento = rs!VCHA_ESB_NOMBRE
                        var_clave_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
                     End If
                     If IsNull(rs!VCHA_AGE_NOMBRE) Then
                        GoTo no_agente:
                     Else
                        txt_agente = rs!VCHA_AGE_NOMBRE
                        var_clave_agente = rs!VCHA_AGE_AGENTE_ID
                     End If
                     var_almacen_Destino = ""
                     If var_tipo_documento = "V" Then
                        If IsNull(rs!almacen_agente) Then
                           GoTo no_almacen_agente:
                        Else
                           var_almacen_Destino = rs!almacen_agente
                        End If
                     End If
                     If IsNull(rs!VCHA_CLI_NOMBRE) Then
                        GoTo no_cliente:
                     Else
                        txt_cliente = rs!VCHA_CLI_NOMBRE
                        var_clave_cliente = rs!vcha_cli_clave_id
                     End If
                     If IsNull(rs!vcha_rut_nombre) Then
                        txt_ruta = ""
                        var_clave_ruta = ""
                     Else
                        txt_ruta = rs!vcha_rut_nombre
                        var_clave_ruta = rs!vcha_rut_ruta_id
                     End If
                     If IsNull(rs!inte_ped_numero) Then
                        GoTo no_Pedido:
                     Else
                        txt_pedido = rs!inte_ped_numero
                     End If
                     If IsNull(rs!FLOA_ORS_DESCUENTO_1) Then
                        txt_descuento1 = 0
                        var_descuento_1 = 0
                     Else
                        txt_descuento1 = rs!FLOA_ORS_DESCUENTO_1
                        var_descuento_1 = rs!FLOA_ORS_DESCUENTO_1
                     End If
                     If IsNull(rs!FLOA_ORS_DESCUENTO_2) Then
                        txt_descuento2 = 0
                        var_descuento_2 = 0
                     Else
                        txt_descuento2 = rs!FLOA_ORS_DESCUENTO_2
                        var_descuento_2 = rs!FLOA_ORS_DESCUENTO_2
                     End If
                     var_descuento_3 = 0
                     While Not rs.EOF
                        Set list_item = lv_salidas.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
                            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_nombre_español), "", Trim(rs!vcha_Art_nombre_español))
                            var_surtir = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR)
                            list_item.SubItems(2) = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0.00")
                            var_surtida = 0
                            list_item.SubItems(3) = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA) + IIf(IsNull(rs!floa_ors_cantidad_negada), 0, rs!floa_ors_cantidad_negada)
                            var_empacada = 0
                            list_item.SubItems(4) = Format(0, "###,###,##0.00")
                            list_item.SubItems(5) = Format(0, "###,###,##0.00")
                            var_falta = var_surtida + var_empacada
                            list_item.SubItems(6) = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR) - (IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA) + IIf(IsNull(rs!floa_ors_cantidad_negada), 0, rs!floa_ors_cantidad_negada))
                            list_item.SubItems(7) = IIf(IsNull(rs!floa_ors_costo), 0, rs!floa_ors_costo)
                            list_item.SubItems(8) = IIf(IsNull(rs!floa_ors_precio), 0, rs!floa_ors_precio)
                            list_item.SubItems(11) = IIf(IsNull(rs!char_ped_tipo), "P", rs!char_ped_tipo)
                            var_suma_cantidad_enviada = var_suma_cantidad_enviada + rs!FLOA_ORS_CANTIDAD_SURTIR
                            var_suma_cantidad_recibida = var_suma_cantidad_recibida + rs!FLOA_ORS_CANTIDAD_SURTIDA
                         rs.MoveNext:
                     Wend
                     If var_tipo_lectura = 0 Then
                        If rsaux4.State = 1 Then
                           rsaux4.Close
                        End If
                        var_nombre_tabla = "TEMP_" + Trim(var_empresa) + Trim(var_unidad_organizacional) + Trim(var_almacen_origen) + Trim(var_clave_movimiento) + Trim(CStr(var_numero_folio))
                        rsaux3.Open "select * from dbo.sysobjects where id = object_id(N'[dbo].[" + var_nombre_tabla + "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1", cnn, adOpenDynamic, adLockOptimistic
                        If rsaux3.EOF Then
                           rsaux3.Close
                           Cadena = "CREATE TABLE [dbo].[" + var_nombre_tabla + "] ([VCHA_EMP_EMPRESA_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_UOR_UNIDAD_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_ALM_ALMACEN_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_MOV_MOVIMIENTO_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[INTE_SAL_NUMERO] [int] NULL ,[VCHA_ART_ARTICULO_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[FLOA_SAL_CANTIDAD] [float] NULL ,[FLOA_SAL_COSTO] [float] NULL ,[FLOA_SAL_PRECIO] [float] NULL ,[FLOA_SAL_DESCUENTO] [float] NULL ,[FLOA_SAL_PROMOCION_1] [float] NULL ,[FLOA_SAL_PROMOCION_2] [float] NULL ,[VCHA_REE_FOLIO] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_SAL_REFERENCIA] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[CHAR_PED_TIPO] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_CAT_CATALOGO_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[FLOA_SAL_DESCUENTO_1] [float] NULL ,"
                           Cadena = Cadena + " [FLOA_SAL_DESCUENTO_2] [float] NULL ,[INTE_SAL_AÑO] [int] NULL , [INTE_SAL_CONSECUTIVO] [int] NULL) ON [PRIMARY]"
                           rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                           rsaux3.Open "INSERT INTO " + var_nombre_tabla + " (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, CHAR_PED_TIPO, INTE_SAL_CONSECUTIVO) select VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, CHAR_PED_TIPO, INTE_SAL_CONSECUTIVO from tb_temporal_salidas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id ='" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + Var_calave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux3.Close
                        End If
                        rsaux.Open "select * from " + var_nombre_tabla + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id ='" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux.EOF
                              valor = rsaux!VCHA_ART_ARTICULO_ID
                              var_n = lv_salidas.ListItems.Count
                              var_encontro = 0
                              var_i = 1
                              var_tipo_pedido = rsaux!char_ped_tipo
                              While (var_i <= var_n)
                                  lv_salidas.ListItems.item(var_i).Selected = True
                                  If valor = lv_salidas.selectedItem And var_tipo_pedido = lv_salidas.selectedItem.SubItems(11) Then
                                     var_encontro = 1
                                     var_i = var_n + 1
                                  Else
                                     var_encontro = 0
                                  End If
                                  var_i = var_i + 1
                              Wend
                              lv_salidas.selectedItem.SubItems(5) = Format(rsaux!floa_Sal_Cantidad, "###,###,##0.00")
                              rsaux.MoveNext
                        Wend
                        rsaux.Close
                        rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 1 WHERE vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + txt_busqueda_folio + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                        lbl_enviados = Format(Str(var_suma_cantidad_enviada), "###,###,##0.00")
                        lbl_recibidos = Format(Str(var_suma_cantidad_recibida), "###,###,##0.00")
                        txt_archivo.Enabled = False
                     Else
'''' metodo nuevo

                        var_archivo_tabla = Trim(var_empresa) + Trim(var_unidad_organizacional) + Trim(var_almacen_origen) + Trim(var_clave_movimiento) + Trim(CStr(var_numero_folio))
                        If rsaux4.State = 1 Then
                           rsaux4.Close
                        End If
                        
                        rsaux.Open "select * from tb_salidas where vcha_sal_Archivo = '" + var_archivo_tabla + "'", cnnaccess, adOpenDynamic, adLockOptimistic
                        While Not rsaux.EOF
                              valor = rsaux!VCHA_ART_ARTICULO_ID
                              var_n = lv_salidas.ListItems.Count
                              var_encontro = 0
                              var_i = 1
                              var_tipo_pedido = rsaux!vcha_sal_tipo
                              While (var_i <= var_n)
                                  lv_salidas.ListItems.item(var_i).Selected = True
                                  If valor = lv_salidas.selectedItem Then
                                     var_encontro = 1
                                     var_i = var_n + 1
                                  Else
                                     var_encontro = 0
                                  End If
                                  var_i = var_i + 1
                              Wend
                              lv_salidas.selectedItem.SubItems(5) = Format(rsaux!floa_Sal_Cantidad, "###,###,##0.00")
                              rsaux.MoveNext
                        Wend
                        rsaux.Close
                        rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 1 WHERE vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + txt_busqueda_folio + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                        lbl_enviados = Format(Str(var_suma_cantidad_enviada), "###,###,##0.00")
                        lbl_recibidos = Format(Str(var_suma_cantidad_recibida), "###,###,##0.00")
                        txt_archivo.Enabled = False


'''' fin metodo nuevo
                     End If
                  Else
                     MsgBox "Numero de Orden de surtido no existe", vbOKOnly, "ATENCION"
                  End If
                  rs.Close
                  var_renglon = -1
                  Call ilumina_grid
                  frm_movimientos.Visible = False
                  If txt_codigo.Enabled = True Then
                     txt_codigo.SetFocus
                  End If
               Else
                  rs.Close
                  MsgBox "El movimiento esta siendo usado por otro usuario", vbOKOnly, "ATENCION"
                  frm_movimientos.Visible = False
               End If
            Else
               rs.Close
               MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
               frm_movimientos.Visible = False
            End If
         End If
      var_n = lv_salidas.ListItems.Count
      var_numero_renglones = lv_salidas.Height / 312.5
      If var_n > var_numero_renglones Then
         lv_salidas.ColumnHeaders(2).Width = 4100.15
      Else
         lv_salidas.ColumnHeaders(2).Width = 4300.15
      End If
      End If
   Else
      rsaux4.Close
      rs.Open "select * from tb_encabezado_embarques where inte_emb_embarque = " + Str(var_numero_embarque) + " AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_embarque_cerrado = IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", Trim(rs!CHAR_EMB_ESTATUS))
      End If
      rs.Close
      rs.Open "select DISTINCT INTE_ORS_ORDEN_SURTIDO,INTE_PAQ_CAJA from tb_detalle_cajas where inte_emb_embarque = " + Me.txt_embarque + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and (char_paq_estatus <> 'S' and char_paq_estatus <> 'C')", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_existen_cajas = 1
      Else
         var_existen_cajas = 0
      End If
      rs.Close
      If var_existen_cajas = 0 Then
         If Trim(var_embarque_cerrado) = "" Then
            rs.Open "select * from tb_Sellos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Str(var_numero_embarque), cnn, adOpenDynamic, adLockOptimistic
            var_numero_items = 0
            If Not rs.EOF Then
               While Not rs.EOF
                     Set list_item = lv_sellos.ListItems.Add(, , rs!vcha_sel_Sello)
                     rs.MoveNext
                     var_numero_items = var_numero_items + 1
               Wend
            End If
            If var_numero_items > 5 Then
               lv_sellos.ColumnHeaders(1).Width = 2650
            Else
               lv_sellos.ColumnHeaders(1).Width = 2850
            End If
            rs.Close
            frm_sellos.Visible = True
            txt_sello.SetFocus
         Else
            MsgBox "El embarque ya habia sido cerrado con anterioridad", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Faltan cajas sin subir", vbOKOnly, "ATENCION"
      End If
   End If
   Exit Sub
no_almacen:
    MsgBox "Almacen Incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_Pedido:
    MsgBox "Pedido Incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_establecimiento:
    MsgBox "Establecimiento Incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_agente:
    MsgBox "Agente Incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_cliente:
    MsgBox "Cliente Incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_ruta:
    MsgBox "Ruta Incorrecta", vbOKOnly, "ATENCION"
    Exit Sub
no_titular:
    MsgBox "Titular incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_almacen_agente:
    MsgBox "No existe un almacen relacionado con este agente", vbOKOnly, "ATENCION"
    Exit Sub
End Sub

Private Sub cmd_imprimir_Click()
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
   If var_numero_folio > 0 Then
      If var_tipo_documento = "F" Then
         MsgBox "No existe una acción para este movimiento", vbOKOnly, "ATENCION"
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
               rsaux2.Open "delete from " + App.Path & "\temp_" + Trim(var_nombre_archivo) + ".dbf", var_tabla, adOpenDynamic, adLockOptimistic
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
            MsgBox "No se a cerrado el embarque", vbOKOnly, "ATENCION"
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
            MsgBox "No se a cerrado el embarque", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
   End If
archivo_ocupado:
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
End Sub

Private Sub cmd_mensaje_1_Click()
   Me.wmp1.Controls.Play
End Sub

Private Sub cmd_mensaje_2_Click()
   Me.wmp2.Controls.Play
End Sub

Private Sub cmd_mensaje_4_Click()
   Me.wmp4.Controls.Play
End Sub

Private Sub cmd_nuevo_Click()
   If var_numero_folio > 0 Then
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   var_consecutivo = 0
   lbl_cancelado = ""
   lv_salidas.ListItems.Clear
   var_primera_vez = True
   txt_origen = ""
   txt_archivo = ""
   txt_titular = ""
   txt_agente = ""
   txt_establecimiento = ""
   txt_cliente = ""
   txt_ruta = ""
   txt_pedido = ""
   txt_descuento1 = ""
   txt_descuento2 = ""
   lv_salidas.ListItems.Clear
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
   txt_codigo.Enabled = False
   txt_foco.Enabled = False
   If Me.txt_archivo.Enabled = True Then
      Me.txt_archivo.SetFocus
   End If
End Sub

Private Sub cmd_pasar_movimiento_Click()
   Dim pError As ADODB.Error
   Dim var_actualiza As Boolean
   Dim var_inserta As Boolean
   Dim bandera_suma As Boolean
   Dim var_cantidad As Variant
   Dim var_costo As Variant
   Dim var_precio As Variant
   Dim var_posible_caja As Boolean
   Dim var_cantidad_posible As Variant
   Dim var_embarque_paquete As Integer
   Dim var_embarque_caja As Integer
   Dim var_estatus_caja As String
   Dim var_orden_surtido_caja As Integer
   Dim var_posible_empaque As Boolean
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   Dim var_encontrado As Integer
   Dim var_canal_venta As String
   Dim var_i As Integer
   Dim var_n As Integer
   Dim var_j As Integer
   Dim var_tipo_pedido As String
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_DET_ORDEN_SURTIDO_M = New TB_DET_ORDEN_SURTIDO_M
   Set TB_DET_EMBARQUE_I = New TB_DET_EMBARQUE_I
   Set TB_DETALLE_CAJAS_M = New TB_DETALLE_CAJAS_M
   'On Error GoTo salir:
   z = 0
   If rsaux6.State = 1 Then
      rsaux6.Close
   End If
   rsaux6.Open "select * from TB_dET_ORDEN_SURTIDO  where INTE_ORS_ORDEN_SURTIDO = 3", cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux6.EOF
   txt_codigo = rsaux6!VCHA_ART_ARTICULO_ID
   var_cantidad_leida = rsaux6!FLOA_ORS_CANTIDAD_SURTIR
   cnn.CommandTimeout = 360
   If Trim(txt_codigo.Text) <> "" Then
      var_posible_empaque = False 'sirve para no meter articulos a granel con cajas
      If var_es_caja = True And var_cajas = True Then
         var_posible_empaque = True
      End If
      If var_es_caja = False And var_cajas = False Then
         var_posible_empaque = True
      End If
      If var_posible_empaque = True Then
         var_posible_caja = False
         bandera_suma = False
         If var_primera_vez = True Then
            var_inserta = False
            rsaux.Open "select vcha_can_canal_venta_id from tb_agentes where vcha_age_agente_id = '" + var_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
            var_canal_venta = IIf(IsNull(rsaux!vcha_can_canal_venta_id), "", rsaux!vcha_can_canal_venta_id)
            rsaux.Close
            var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, var_orden_surtido, var_clave_cliente, "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", txt_archivo, var_clave_establecimiento, "B", var_clave_titular, var_clave_agente, var_descuento_1, var_descuento_2, var_descuento_3, var_clave_moneda, 0)
            var_numero_folio = var_numero_folio_regreso
            If var_factura_ceros = 1 Then
               rsaux.Open "update tb_encabezado_movimientos set inte_emo_factura_ceros = 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            rsaux.Open "update tb_encabezado_movimientos set VCHA_CAN_CANAL_VENTA_ID = '" + var_canal_venta + "' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            var_inserta = False
            var_inserta = TB_DET_EMBARQUE_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_numero_embarque, var_clave_movimiento, var_numero_folio, "")
            txt_folio = var_numero_folio
            var_primera_vez = False
            '
            'var_nombre_tabla = "TEMP_" + Trim(var_empresa) + Trim(var_unidad_organizacional) + Trim(var_almacen_origen) + Trim(var_clave_movimiento) + Trim(CStr(var_numero_folio))
            'Cadena = "CREATE TABLE [dbo].[" + var_nombre_tabla + "] ([VCHA_EMP_EMPRESA_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_UOR_UNIDAD_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_ALM_ALMACEN_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_MOV_MOVIMIENTO_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[INTE_SAL_NUMERO] [int] NULL ,[VCHA_ART_ARTICULO_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[FLOA_SAL_CANTIDAD] [float] NULL ,[FLOA_SAL_COSTO] [float] NULL ,[FLOA_SAL_PRECIO] [float] NULL ,[FLOA_SAL_DESCUENTO] [float] NULL ,[FLOA_SAL_PROMOCION_1] [float] NULL ,[FLOA_SAL_PROMOCION_2] [float] NULL ,[VCHA_REE_FOLIO] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_SAL_REFERENCIA] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[CHAR_PED_TIPO] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_CAT_CATALOGO_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[FLOA_SAL_DESCUENTO_1] [float] NULL ,"
            'Cadena = Cadena + " [FLOA_SAL_DESCUENTO_2] [float] NULL ,[INTE_SAL_AÑO] [int] NULL , [INTE_SAL_CONSECUTIVO] [int] NULL) ON [PRIMARY]"
            'rsaux4.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            '
            If var_tipo_lectura = 1 Then
               var_i = 1
               For var_i = 1 To lv_salidas.ListItems.Count
                   lv_salidas.ListItems.item(var_i).Selected = True
                   If var_tipo_lectura = 1 Then
                      
                      var_precio = CDbl(lv_salidas.selectedItem.SubItems(8)) * 1
                      If var_factura_ceros = 1 Then
                         var_precio = 0
                      End If
                      
                      Cadena = "insert into tb_salidas (VCHA_SAL_ARCHIVO, INTE_PED_NUMERO, INTE_ORS_ORDEN_SURTIDO, VCHA_EMP_EMPRESA_ID, INTE_SAL_NUMERO,VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ORS_CANTIDAD_SURTIR, FLOA_ORS_CANTIDAD_SURTIDA, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, "
                      Cadena = Cadena + " VCHA_SAL_TIPO, INTE_SAL_CONSECUTIVO) VALUES ('" + Trim(var_empresa) + Trim(var_unidad_organizacional) + Trim(var_almacen_origen) + Trim(var_clave_movimiento) + Trim(CStr(var_numero_folio)) + "'," + Trim(txt_pedido) + "," + txt_archivo + ",'" + var_empresa + "'," + Trim(CStr(var_numero_folio)) + ",'" + lv_salidas.selectedItem + "',''," + CStr(CDbl(lv_salidas.selectedItem.SubItems(2)) * 1) + ", " + CStr(CDbl(lv_salidas.selectedItem.SubItems(3)) * 1) + ",0," + CStr(CDbl(lv_salidas.selectedItem.SubItems(7)) * 1) + "," + CStr(var_precio) + "," + CStr(CDbl(lv_salidas.selectedItem.SubItems(9)) * 1) + "," + CStr(CDbl(lv_salidas.selectedItem.SubItems(10)) * 1) + ",'" + lv_salidas.selectedItem.SubItems(11) + "',0)"
                      rsaux4.Open Cadena, cnnaccess, adOpenDynamic, adLockOptimistic
                   End If
               Next var_i
            End If
         End If
         If var_tipo_lectura = 0 Then
            If var_es_caja = False Then
               Cadena = "select * from tb_det_orden_surtido where inte_ors_orden_surtido = " + txt_archivo + " and vcha_art_articulo_id = '" + txt_codigo + "'"
               rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_promocion_1 = IIf(IsNull(rs!floa_ors_promocion_1), 0, rs!floa_ors_promocion_1)
                  var_promocion_2 = IIf(IsNull(rs!floa_ors_promocion_2), 0, rs!floa_ors_promocion_2)
                  valor = txt_codigo
                  var_n = lv_salidas.ListItems.Count
                  var_encontro = 0
                  var_i = 1
                  While (var_i <= var_n)
                        var_tipo_pedido = lv_salidas.selectedItem.SubItems(11)
                        lv_salidas.ListItems.item(var_i).Selected = True
                        valor = Trim(lv_salidas.selectedItem)
                        If txt_codigo = valor Then
                           var_cantidad_posible = lv_salidas.selectedItem.SubItems(2)
                           If var_cantidad_posible < lv_salidas.selectedItem.SubItems(3) + var_cantidad_leida Then
                              var_encontro = 0
                           Else
                              var_encontro = 1
                              var_i = var_n + 1
                           End If
                        End If
                        var_i = var_i + 1
                  Wend
                  If var_encontro = 1 Then
                     var_cantidad_posible = lv_salidas.selectedItem.SubItems(2)
                     
                     If var_cantidad_posible < lv_salidas.selectedItem.SubItems(3) + var_cantidad_leida Then
                        Call cmd_mensaje_2_Click
                        frmmensaje.lbl_mensaje = "Cantidad supera a la posible a surtir"
                        frmmensaje.Show 1
                     Else
                        var_tipo_pedido = lv_salidas.selectedItem.SubItems(11)
                        lv_salidas.selectedItem.SubItems(6) = Format(lv_salidas.selectedItem.SubItems(2) - (var_cantidad_leida + lv_salidas.selectedItem.SubItems(3) + lv_salidas.selectedItem.SubItems(4)), "###,###,##0.00")
                        lv_salidas.selectedItem.SubItems(4) = lv_salidas.selectedItem.SubItems(4)
                        lv_salidas.selectedItem.SubItems(3) = Format(lv_salidas.selectedItem.SubItems(3) + var_cantidad_leida, "###,###,##0.00")
                        lv_salidas.selectedItem.SubItems(5) = Format(lv_salidas.selectedItem.SubItems(5) + var_cantidad_leida, "###,###,##0.00")
                        var_renglon = lv_salidas.selectedItem.Index
                        Call ilumina_grid
                        var_costo = lv_salidas.selectedItem.SubItems(7)
                        var_precio = lv_salidas.selectedItem.SubItems(8)
                        var_cantidad = lv_salidas.selectedItem.SubItems(4)
                        lbl_recibidos = Format(CDbl(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                        var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                        If rsaux5.State = 1 Then
                           rsaux5.Close
                        End If
                        '26-09-08 para afectar orden surtido hasta que se cierra el embarque
                        rsaux5.Open "update tb_det_orden_surtido set floa_ors_cantidad_surtida = floa_ors_cantidad_surtida + " + CStr(var_cantidad_leida) + " where inte_ors_orden_surtido = " + CStr(var_orden_surtido) + " and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        If rsaux5.State = 1 Then
                           rsaux5.Close
                        End If
                        bandera_suma = True
                     End If
                  Else
                     Call cmd_mensaje_2_Click
                     txt_codigo = ""
                     frmmensaje.lbl_articulo = var_nombre_articulo_mensaje
                     frmmensaje.lbl_mensaje = "Cantidad supera a la posible a surtir"
                     frmmensaje.Show 1
                  End If
               Else
                  Call cmd_mensaje_2_Click
                  txt_codigo = ""
                  frmmensaje.lbl_articulo = var_nombre_articulo_mensaje
                  frmmensaje.lbl_mensaje = "El artículo no se encuentra dentro de la Orden de Surtido"
                  frmmensaje.Show 1
               End If
               rs.Close
               If bandera_suma = True Then
                  If var_factura_ceros = 1 Then
                     var_precio = 0
                  End If
                  Cadena = "select * from " + var_nombre_tabla + " where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "' and floa_sal_precio = " + CStr(var_precio) + " and char_ped_tipo = '" + var_tipo_pedido + "'"
                  rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_inserta = False
                     rsaux.Open "update " + var_nombre_tabla + " set floa_sal_cantidad = floa_sal_cantidad +" + CStr(var_cantidad_leida) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_Sal_Numero = " + CStr(var_numero_folio) + " and vcha_art_articulo_id = '" + lv_salidas.selectedItem + "' and round(floa_sal_precio,2) = round(" + CStr(var_precio) + ",2) and char_ped_tipo = '" + var_tipo_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
                     rs.Close
                  Else
                     var_inserta = False
                     var_consecutivo = var_consecutivo + 1
                     rsaux.Open "INSERT INTO " + var_nombre_tabla + " (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, CHAR_PED_TIPO, INTE_SAL_CONSECUTIVO) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ",  " + CStr(var_precio) + ", 0,  " + CStr(var_promocion_1) + ", " + CStr(var_promocion_2) + ",'" + var_tipo_pedido + "', " + CStr(var_consecutivo) + ") ", cnn, adOpenDynamic, adLockOptimistic
                     rs.Close
                  End If
                  bandera_suma = False
               End If
            Else
            End If
         Else
''''metodo nuevo
            'cnnaccess.BeginTrans
            If var_es_caja = False Then
               var_archivo_tabla = Trim(var_empresa) + Trim(var_unidad_organizacional) + Trim(var_almacen_origen) + Trim(var_clave_movimiento) + Trim(CStr(var_numero_folio))
               Cadena = "select * from tb_salidas where vcha_sal_archivo = '" + var_archivo_tabla + "' and inte_ors_orden_surtido = " + txt_archivo + " and vcha_art_articulo_id = '" + txt_codigo + "'"
               If rs.State = 1 Then
                  rs.Close
               End If
               rs.Open Cadena, cnnaccess, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_promocion_1 = IIf(IsNull(rs!floa_sal_promocion_1), 0, rs!floa_sal_promocion_1)
                  var_promocion_2 = IIf(IsNull(rs!FLOA_SAL_PROMOCION_2), 0, rs!FLOA_SAL_PROMOCION_2)
                  valor = txt_codigo
                  var_n = lv_salidas.ListItems.Count
                  var_encontro = 0
                  var_i = 1
                  While (var_i <= var_n)
                        var_tipo_pedido = lv_salidas.selectedItem.SubItems(11)
                        lv_salidas.ListItems.item(var_i).Selected = True
                        valor = Trim(lv_salidas.selectedItem)
                        If txt_codigo = valor Then
                           var_cantidad_posible = lv_salidas.selectedItem.SubItems(2)
                           If var_cantidad_posible < lv_salidas.selectedItem.SubItems(3) + var_cantidad_leida Then
                              var_encontro = 0
                           Else
                              var_encontro = 1
                              var_i = var_n + 1
                           End If
                        End If
                        var_i = var_i + 1
                  Wend
                  If var_encontro = 1 Then
                     var_cantidad_posible = lv_salidas.selectedItem.SubItems(2)
                     If var_cantidad_posible < lv_salidas.selectedItem.SubItems(3) + var_cantidad_leida Then
                        Call cmd_mensaje_2_Click
                        frmmensaje.lbl_mensaje = "Cantidad supera a la posible a surtir"
                        frmmensaje.Show 1
                     Else
                        var_tipo_pedido = lv_salidas.selectedItem.SubItems(11)
                        lv_salidas.selectedItem.SubItems(6) = Format(lv_salidas.selectedItem.SubItems(2) - (var_cantidad_leida + lv_salidas.selectedItem.SubItems(3) + lv_salidas.selectedItem.SubItems(4)), "###,###,##0.00")
                        lv_salidas.selectedItem.SubItems(4) = lv_salidas.selectedItem.SubItems(4)
                        lv_salidas.selectedItem.SubItems(3) = Format(lv_salidas.selectedItem.SubItems(3) + var_cantidad_leida, "###,###,##0.00")
                        lv_salidas.selectedItem.SubItems(5) = Format(lv_salidas.selectedItem.SubItems(5) + var_cantidad_leida, "###,###,##0.00")
                        var_renglon = lv_salidas.selectedItem.Index
                        Call ilumina_grid
                        var_costo = lv_salidas.selectedItem.SubItems(7)
                        var_precio = lv_salidas.selectedItem.SubItems(8)
                        var_cantidad = lv_salidas.selectedItem.SubItems(4)
                        lbl_recibidos = Format(CDbl(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                        var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                        If rsaux5.State = 1 Then
                           rsaux5.Close
                        End If
                        rsaux5.Open "update tb_Salidas set floa_ors_cantidad_surtida = floa_ors_cantidad_surtida + " + CStr(var_cantidad_leida) + " where inte_ors_orden_surtido = " + CStr(var_orden_surtido) + " and vcha_art_articulo_id = '" + txt_codigo + "' and vcha_sal_Archivo = '" + var_archivo_tabla + "'", cnnaccess, adOpenDynamic, adLockOptimistic
                        If rsaux5.State = 1 Then
                           rsaux5.Close
                        End If
                        bandera_suma = True
                     End If
                  Else
                     Call cmd_mensaje_2_Click
                     txt_codigo = ""
                     frmmensaje.lbl_articulo = var_nombre_articulo_mensaje
                     frmmensaje.lbl_mensaje = "Cantidad supera a la posible a surtir"
                     frmmensaje.Show 1
                  End If
               Else
                  Call cmd_mensaje_2_Click
                  txt_codigo = ""
                  frmmensaje.lbl_articulo = var_nombre_articulo_mensaje
                  frmmensaje.lbl_mensaje = "El artículo no se encuentra dentro de la Orden de Surtido"
                  frmmensaje.Show 1
               End If
               rs.Close
               If bandera_suma = True Then
                  If var_factura_ceros = 1 Then
                     var_precio = 0
                  End If
                  var_inserta = False
                  If rsaux4.State = 1 Then
                     rsaux4.Close
                  End If
                  '26-09-08 para afectar orden surtido hasta que se cierra el embarque
                  rsaux5.Open "update tb_det_orden_surtido set floa_ors_cantidad_surtida = floa_ors_cantidad_surtida + " + CStr(var_cantidad_leida) + " where inte_ors_orden_surtido = " + CStr(var_orden_surtido) + " and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux4.Open "SELECT * FROM TB_SALIDAS where vcha_sal_archivo = '" + var_archivo_tabla + "' and vcha_art_articulo_id = '" + lv_salidas.selectedItem + "' AND INTE_SAL_CONSECUTIVO = 0", cnnaccess, adOpenDynamic, adLockOptimistic
                  If Not rsaux4.EOF Then
                     var_consecutivo = var_consecutivo + 1
                     rsaux.Open "update tb_salidas set floa_sal_cantidad = floa_sal_cantidad +" + CStr(var_cantidad_leida) + ", INTE_SAL_CONSECUTIVO = " + CStr(var_consecutivo) + " where vcha_sal_archivo = '" + var_archivo_tabla + "' and vcha_art_articulo_id = '" + lv_salidas.selectedItem + "'", cnnaccess, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux.Open "update tb_salidas set floa_sal_cantidad = floa_sal_cantidad +" + CStr(var_cantidad_leida) + " where vcha_sal_archivo = '" + var_archivo_tabla + "' and vcha_art_articulo_id = '" + lv_salidas.selectedItem + "'", cnnaccess, adOpenDynamic, adLockOptimistic
                  End If
                  bandera_suma = False
               End If
            Else
            End If
            'cnnaccess.CommitTrans
''''' fin metodo nuevo
         End If
      Else
         txt_codigo = ""
         frmmensaje.lbl_mensaje = "No es posible mezclar mercancia a granel con mercancia empacada"
         frmmensaje.Show 1
      End If
      txt_codigo.SetFocus
   End If
'   Exit Sub
'salir:
'Resume
      rsaux6.MoveNext
   Wend
   i = i
End Sub

Private Sub cmd_restructuracion_Click()
   'aqui fallo la red y se restructuraron los movimientos de lo que se quedo grabado en la orden de surtido
   'para esto se creo un vista que se llama vw_restructurar_movimientos en donde se le pasa el numero del embarque a restructurar
     rs.Open "SELECT * FROM VW_RESTRUCTURAR_MOVIMIENTOS", cnn, adOpenDynamic, adLockOptimistic
     var_i = 1
     While Not rs.EOF
       Cadena = "insert into tb_salidas (VCHA_SAL_ARCHIVO, INTE_PED_NUMERO, INTE_ORS_ORDEN_SURTIDO, VCHA_EMP_EMPRESA_ID, INTE_SAL_NUMERO,VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ORS_CANTIDAD_SURTIR, FLOA_ORS_CANTIDAD_SURTIDA, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, "
                         Cadena = Cadena + " VCHA_SAL_TIPO, INTE_SAL_CONSECUTIVO) "
         Cadena = Cadena + "VALUES ('" + Trim(rs!VCHA_EMP_EMPRESA_ID) + Trim(rs!VCHA_UOR_UNIDAD_ID) + Trim(rs!VCHA_ALM_ALMACEN_ID) + Trim(rs!VCHA_MOV_MOVIMIENTO_ID) + Trim(CStr(rs!INTE_SAL_NUMERO)) + "'," + Trim(rs!inte_ped_numero) + "," + CStr(rs!INTE_ORS_ORDEN_SURTIDO) + ",'" + rs!VCHA_EMP_EMPRESA_ID + "'," + Trim(CStr(rs!INTE_SAL_NUMERO)) + ",'" + rs!VCHA_ART_ARTICULO_ID + "',''," + CStr(rs!FLOA_ORS_CANTIDAD_SURTIR) + ", " + CStr(rs!FLOA_ORS_CANTIDAD_SURTIDA) + "," + CStr(rs!FLOA_ORS_CANTIDAD_SURTIDA) + "," + CStr(rs!floa_ors_costo) + "," + CStr(rs!floa_ors_precio) + "," + CStr(rs!floa_ors_promocion_1) + "," + CStr(rs!floa_ors_promocion_2) + ",'" + rs!char_tpe_tipo_pedido_id + "'," + CStr(var_i) + ")"
         rsaux4.Open Cadena, cnnaccess, adOpenDynamic, adLockOptimistic
         rs.MoveNext
         var_i = var_i + 1
     Wend
     rs.Close
     
End Sub

Private Sub cmd_restructurar_Click()
   Set TB_DET_ORDEN_SURTIDO_M = New TB_DET_ORDEN_SURTIDO_M
   Set TB_DET_EMBARQUE_I = New TB_DET_EMBARQUE_I
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   rs.Open "select distinct inte_ped_numero, inte_ors_orden_surtido, inte_sal_numero from tb_salidas", cnnaccess, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_si = MsgBox("¿Deseas restructurar la orden de surtido número " + CStr(rs!INTE_ORS_ORDEN_SURTIDO), vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_orden_surtido = rs!INTE_ORS_ORDEN_SURTIDO
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open "select VCHA_ALM_ALMACEN_ID, vcha_age_agente_id, vcha_esb_establecimiento_id, vcha_cli_clave_id, vcha_tit_titular_id, floa_ped_descuento_1, floa_ped_descuento_2 from tb_encabezado_pedidos where inte_ped_numero = " + CStr(rs!inte_ped_numero), cnn, adOpenDynamic, adLockOptimistic
            var_almacen_origen = rsaux!VCHA_ALM_ALMACEN_ID
            var_clave_cliente = rsaux!vcha_cli_clave_id
            var_clave_establecimiento = rsaux!vcha_ESB_ESTABLECIMIENTO_id
            var_clave_titular = rsaux!vcha_tit_titular_id
            var_clave_agente = rsaux!VCHA_AGE_AGENTE_ID
            var_descuento_1 = rsaux!floa_ped_descuento_1
            var_descuento_2 = rsaux!FLOA_PED_DESCUENTO_2
            var_descuento_3 = 0
            rsaux.Close
            rsaux.Open "select * from tb_clientes where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
            var_clave_moneda = rsaux!vcha_mon_moneda_id
            rsaux.Close
            rsaux.Open "select vcha_can_canal_venta_id from tb_agentes where vcha_age_agente_id = '" + var_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
            var_canal_venta = IIf(IsNull(rsaux!vcha_can_canal_venta_id), "", rsaux!vcha_can_canal_venta_id)
            rsaux.Close
            txt_archivo = var_orden_surtido
         
            var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, var_orden_surtido, var_clave_cliente, "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", txt_archivo, var_clave_establecimiento, "B", var_clave_titular, var_clave_agente, var_descuento_1, var_descuento_2, var_descuento_3, var_clave_moneda, 0)
            var_numero_folio = var_numero_folio_regreso
            If var_factura_ceros = 1 Then
               rsaux.Open "update tb_encabezado_movimientos set inte_emo_factura_ceros = 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            rsaux.Open "update tb_encabezado_movimientos set VCHA_CAN_CANAL_VENTA_ID = '" + var_canal_venta + "' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            var_inserta = False
            var_inserta = TB_DET_EMBARQUE_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_numero_embarque, var_clave_movimiento, var_numero_folio, "")
            rsaux.Open "select vcha_Art_Articulo_ID, floa_Sal_Cantidad from tb_SaLidas where inte_ors_orden_surtido = " + CStr(rs!INTE_ORS_ORDEN_SURTIDO), cnnaccess, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  '26-09-08 para afectar orden surtido hasta que se cierra el embarque
                  rsaux2.Open "UPDATE TB_dET_ORDEN_SURTIDO SET FLOA_ORS_CANTIDAD_SURTIDA = " + CStr(rsaux!floa_Sal_Cantidad) + " WHERE INTE_ORS_ORDEN_SURTIDO = " + CStr(rs!INTE_ORS_ORDEN_SURTIDO) + " AND VCHA_aRT_ARTICULO_ID = '" + rsaux!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux.MoveNext
            Wend
         End If
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub cmd_resumen_Click()
   If Trim(Me.txt_archivo) <> "" Then
      rs.Open "select * from VW_RESUMEN_ORDEN_SURTIDO_LINEAS where inte_ors_orden_surtido = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
      lv_resumen.ListItems.Clear
      txt_cantidad_total_linea = 0
      If Not rs.EOF Then
         While Not rs.EOF
               Set list_item = lv_resumen.ListItems.Add(, , Trim(rs!VCHA_SUB_SUBDIVISION_ID))
               list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_sub_nombre), "", rs!vcha_sub_nombre))
               list_item.SubItems(2) = IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
               txt_cantidad_total_linea = Format(CDbl(txt_cantidad_total_linea) + IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
               rs.MoveNext
         Wend
      End If
      rs.Close
      If lv_resumen.ListItems.Count > 0 Then
         frm_resumen.Visible = True
         Me.lv_resumen.SetFocus
      Else
         MsgBox "La orden de surtido esta vacia", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 67 Then
      If Me.frm_sellos.Visible = False Then
         cmd_cerrar_embarque_Click
      Else
         cmd_cerrar_Click
      End If
   End If
   If Shift = 1 And KeyCode = 116 Then
      frmexisten_rapidas.Show 1
   End If
   If Shift = 1 And KeyCode = 117 Then
      Set reporte = appl.OpenReport(App.Path + "\rep_PROGRESO_EQUIPOS.rpt")
      reporte.RecordSelectionFormula = "{VW_PROGRESO_EQUIPOS.DTIM_EQU_FECHA} = CURRENTDATE"
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte de Progreso de Surtido"
      frmvistasprevias.Show 1
      Set reporte = Nothing
   End If
   If Shift = 1 And KeyCode = 118 Then
      lv_cajas_faltantes.ListItems.Clear
      rs.Open "select DISTINCT INTE_ORS_ORDEN_SURTIDO,INTE_PAQ_CAJA from tb_detalle_cajas where inte_emb_embarque = " + Me.txt_embarque + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and (char_paq_estatus <> 'S' and char_paq_estatus <> 'C')", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            Set list_item = lv_cajas_faltantes.ListItems.Add(, , rs!INTE_ORS_ORDEN_SURTIDO)
            list_item.SubItems(1) = IIf(IsNull(rs!inte_paq_caja), 0, Trim(rs!inte_paq_caja))
            rs.MoveNext
         Wend
         frm_cajas_faltantes.Visible = True
         lv_cajas_faltantes.SetFocus
      Else
         MsgBox "No existen cajas de esta orden de surtido", vbOKOnly, "ATENCION"
      End If
      rs.Close
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
   End If
   If Shift = 4 And KeyCode = 77 Then
      cmd_cerrar_embarque_Click
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
       If Me.frm_busqueda.Visible = True Then
          Me.frm_busqueda.Visible = False
       Else
          If Me.frm_cajas_faltantes.Visible = True Then
             Me.frm_cajas_faltantes.Visible = False
          Else
             If Me.frm_eliminar.Visible = True Then
                Me.frm_eliminar.Visible = False
             Else
                If Me.frm_movimientos.Visible = True Then
                   Me.frm_movimientos.Visible = False
                Else
                   If Me.frm_sellos.Visible = True Then
                      Me.frm_sellos.Visible = False
                   Else
                      If Trim(Me.txt_folio) <> "" Then
                         var_si = MsgBox("¿Deseas salir del movimiento?", vbYesNo, "ATENCION")
                         If var_si = 6 Then
                            Unload Me
                         End If
                      Else
                         Unload Me
                      End If
                   End If
                End If
             End If
          End If
       End If
    End If
End Sub

Private Sub Form_Load()
   var_posible_kanban = 0
   If var_empresa = "18" Then
      Me.cmd_resumen.Visible = True
   Else
      Me.cmd_resumen.Visible = False
   End If
   Me.frm_resumen.Visible = False
   frm_movimientos.Visible = False
   var_cadena_seguridad = ""
   var_factura_ceros = 0
   Top = 0
   Left = 0
   frm_sellos.Visible = False
   var_autorizo_embarque = False
   var_iva = 0
   var_agrupador = ""
   var_cantidad_leida = 1#
   var_estatus_movimiento = ""
   var_almacen_Destino = ""
   var_almacen_origen = ""
   var_proveedor = ""
   var_factura = ""
   frm_eliminar.Visible = False
   var_modifica = False
   txt_cantidad.Visible = False
   lbl_cantidad.Visible = False
   txt_codigo.Enabled = False
   var_primera_vez = True
   frm_busqueda.Visible = False
   var_suma_cantidad_enviada = 0
   var_suma_cantidad_recibida = 0
   txt_embarque = var_numero_embarque
   txt_jaula = var_numero_jaula
   frm_cajas_faltantes.Visible = False
   lbl_cancelado = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_numero_folio > 0 Then
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   Call activa_forma(var_activa_forma_salidas)
End Sub

Private Sub lv_cajas_faltantes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_cajas_faltantes.Visible = False
      If txt_codigo.Enabled = True Then
         txt_codigo.SetFocus
      End If
   End If
End Sub

Private Sub lv_cajas_faltantes_LostFocus()
   frm_cajas_faltantes.Visible = False
End Sub

Private Sub lv_movimientos_KeyPress(KeyAscii As Integer)
Dim var_busqueda_folio As Double
Dim var_busqueda_numero As Double
Dim var_busqueda_referencia As String
Dim var_posible As Boolean
Dim var_maquina_movimiento As String
   If KeyAscii = 13 Then
      If lv_movimientos.ListItems.Count > 0 Then
         txt_busqueda_folio = lv_movimientos.selectedItem.SubItems(1)
         rs.Open "select * from tb_detalle_embarques where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + txt_busqueda_folio + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' AND INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If var_numero_embarque = rs!inte_emb_embarque Then
               var_posible = True
            Else
               MsgBox "Este movimiento se encuentra en el embarque No. " + Str(rs!inte_emb_embarque), vbOKOnly, "ATENCION"
               var_posible = False
               frm_movimientos.Visible = False
            End If
         Else
            MsgBox "El Movimiento no existe", vbOKOnly, "ATENCION"
            var_posible = False
            frm_movimientos.Visible = False
         End If
         rs.Close
         If var_posible = True Then
            cnn.CommandTimeout = 360
            rs.Open "select * from tb_detalle_cajas with (nolock) where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_paq_movimiento_destino = '" + var_clave_movimiento + "' and inte_paq_numero_destino = " + txt_busqueda_folio, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_cajas = True
            Else
               var_cajas = False
            End If
            rs.Close
            If var_numero_folio = CDbl(txt_busqueda_folio) Then
               rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            If var_tipo_lectura = 0 Then
               rs.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + txt_busqueda_folio + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            Else
               rs.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + txt_busqueda_folio + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_aud_maquina = '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
               'rs.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + txt_busqueda_folio + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            If Not rs.EOF Then
               var_fecha_numero = CDbl(CStr(Year(rs!DTIM_EMO_FECHa)) + CStr(Month(rs!DTIM_EMO_FECHa)) + CStr(Day(rs!DTIM_EMO_FECHa)))

               If var_numero_folio > 0 Then
                  rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
               End If
               var_movimiento_bloqueado = IIf(IsNull(rs!INTE_EMO_BLOQUEADO), 0, rs!INTE_EMO_BLOQUEADO)
               If var_movimiento_bloqueado = 0 Then
                  var_primera_vez = False
                  var_factura_ceros = IIf(IsNull(rs!inte_emo_factura_ceros), 0, rs!inte_emo_factura_ceros)
                  var_clave_moneda = rs!vcha_mon_moneda_id
                  var_orden_surtido = rs!inte_emo_numero_origen
                  var_numero_folio = rs!INTE_EMO_NUMERO
                  txt_folio = rs!INTE_EMO_NUMERO
                  var_almacen_origen = rs!VCHA_ALM_ALMACEN_ID
                  rsaux3.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_origen = rsaux3!VCHA_ALM_NOMBRE
                  rsaux3.Close
                  If IsNull(rs!char_Emo_estatus) Then
                     var_estatus_movimiento = ""
                  Else
                     var_estatus_movimiento = rs!char_Emo_estatus
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
                  rs.Close
                  rs.Open "select * from vw_orden_surtido where inte_ors_orden_surtido = " + Str(var_orden_surtido) + " and floa_ors_cantidad_surtir > 0", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     txt_archivo = var_orden_surtido
                     var_suma_cantidad_enviada = 0
                     var_suma_cantidad_recibida = 0
                     lbl_enviados.Caption = "0"
                     lbl_recibidos.Caption = "0"
                     lv_salidas.ListItems.Clear
                     If IsNull(rs!VCHA_TIT_NOMBRE) Then
                        GoTo no_titular:
                     Else
                        txt_titular = rs!VCHA_TIT_NOMBRE
                        var_clave_titular = rs!vcha_tit_titular_id
                     End If
                     If IsNull(rs!inte_ped_dias_condiciones) Then
                        var_plazo = 0
                     Else
                        var_plazo = rs!inte_ped_dias_condiciones
                     End If
                     If IsNull(rs!vcha_cli_email) Then
                        var_correo_electronico = ""
                     Else
                        var_correo_electronico = rs!vcha_cli_email
                     End If
                     If IsNull(rs!VCHA_ESB_NOMBRE) Then
                        GoTo no_establecimiento:
                     Else
                        txt_establecimiento = rs!VCHA_ESB_NOMBRE
                        var_clave_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
                     End If
                     If IsNull(rs!VCHA_AGE_NOMBRE) Then
                        GoTo no_agente:
                     Else
                        txt_agente = rs!VCHA_AGE_NOMBRE
                        var_clave_agente = rs!VCHA_AGE_AGENTE_ID
                     End If
                     var_almacen_Destino = ""
                     If var_tipo_documento = "V" Then
                        If IsNull(rs!almacen_agente) Then
                           GoTo no_almacen_agente:
                        Else
                           var_almacen_Destino = rs!almacen_agente
                        End If
                     End If
                     If IsNull(rs!VCHA_CLI_NOMBRE) Then
                        GoTo no_cliente:
                     Else
                        txt_cliente = rs!VCHA_CLI_NOMBRE
                        var_clave_cliente = rs!vcha_cli_clave_id
                     End If
                     If IsNull(rs!vcha_rut_nombre) Then
                        txt_ruta = ""
                        var_clave_ruta = ""
                     Else
                        txt_ruta = rs!vcha_rut_nombre
                        var_clave_ruta = rs!vcha_rut_ruta_id
                     End If
                     If IsNull(rs!inte_ped_numero) Then
                        GoTo no_Pedido:
                     Else
                        txt_pedido = rs!inte_ped_numero
                     End If
                     If IsNull(rs!FLOA_ORS_DESCUENTO_1) Then
                        txt_descuento1 = 0
                        var_descuento_1 = 0
                     Else
                        txt_descuento1 = rs!FLOA_ORS_DESCUENTO_1
                        var_descuento_1 = rs!FLOA_ORS_DESCUENTO_1
                     End If
                     If IsNull(rs!FLOA_ORS_DESCUENTO_2) Then
                        txt_descuento2 = 0
                        var_descuento_2 = 0
                     Else
                        txt_descuento2 = rs!FLOA_ORS_DESCUENTO_2
                        var_descuento_2 = rs!FLOA_ORS_DESCUENTO_2
                     End If
                     var_descuento_3 = 0
                     
                     
                     While Not rs.EOF
                        Set list_item = lv_salidas.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
                            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_nombre_español), "", Trim(rs!vcha_Art_nombre_español))
                            var_surtir = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR)
                            list_item.SubItems(2) = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0.00")
                            var_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA) + IIf(IsNull(rs!floa_ors_cantidad_negada), 0, rs!floa_ors_cantidad_negada)
                            list_item.SubItems(3) = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA) + IIf(IsNull(rs!floa_ors_cantidad_negada), 0, rs!floa_ors_cantidad_negada), "###,###,##0.00")
                            var_empacada = IIf(IsNull(rs!floa_ors_Cantidad_empacada), 0, rs!floa_ors_Cantidad_empacada)
                            var_empacada = 0
                            list_item.SubItems(4) = Format(IIf(IsNull(rs!floa_ors_Cantidad_empacada), 0, rs!floa_ors_Cantidad_empacada), "###,###,##0.00")
                            list_item.SubItems(5) = Format(0, "###,###,##0.00")
                            var_falta = var_surtida + var_empacada
                            list_item.SubItems(6) = Format(var_surtir - var_surtida, "###,###,##0.00")
                            list_item.SubItems(7) = IIf(IsNull(rs!floa_ors_costo), 0, rs!floa_ors_costo)
                            list_item.SubItems(8) = IIf(IsNull(rs!floa_ors_precio), 0, rs!floa_ors_precio)
                            list_item.SubItems(9) = IIf(IsNull(rs!floa_ors_promocion_1), 0, rs!floa_ors_promocion_1)
                            list_item.SubItems(10) = IIf(IsNull(rs!floa_ors_promocion_2), 0, rs!floa_ors_promocion_2)
                            list_item.SubItems(11) = IIf(IsNull(rs!char_ped_tipo), "P", rs!char_ped_tipo)
                            var_suma_cantidad_enviada = var_suma_cantidad_enviada + rs!FLOA_ORS_CANTIDAD_SURTIR
                            var_suma_cantidad_recibida = var_suma_cantidad_recibida + rs!FLOA_ORS_CANTIDAD_SURTIDA
                         rs.MoveNext:
                     Wend
                     If var_tipo_lectura = 0 Then
                        rsaux.Open "SELECT MAX(INTE_SAL_CONSECUTIVO) FROM TB_TEMPORAL_SALIDAS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id ='" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                        Else
                           var_consecutivo = 0
                        End If
                        rsaux.Close
                        If rsaux4.State = 1 Then
                           rsaux4.Close
                        End If
                        
                        var_nombre_tabla = "TEMP_" + Trim(var_empresa) + Trim(var_unidad_organizacional) + Trim(var_almacen_origen) + Trim(var_clave_movimiento) + Trim(CStr(var_numero_folio))
                        rsaux3.Open "select * from dbo.sysobjects where id = object_id(N'[dbo].[" + var_nombre_tabla + "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1", cnn, adOpenDynamic, adLockOptimistic
                        If rsaux3.EOF Then
                           rsaux3.Close
                           Cadena = "CREATE TABLE [dbo].[" + var_nombre_tabla + "] ([VCHA_EMP_EMPRESA_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_UOR_UNIDAD_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_ALM_ALMACEN_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_MOV_MOVIMIENTO_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[INTE_SAL_NUMERO] [int] NULL ,[VCHA_ART_ARTICULO_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[FLOA_SAL_CANTIDAD] [float] NULL ,[FLOA_SAL_COSTO] [float] NULL ,[FLOA_SAL_PRECIO] [float] NULL ,[FLOA_SAL_DESCUENTO] [float] NULL ,[FLOA_SAL_PROMOCION_1] [float] NULL ,[FLOA_SAL_PROMOCION_2] [float] NULL ,[VCHA_REE_FOLIO] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_SAL_REFERENCIA] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[CHAR_PED_TIPO] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_CAT_CATALOGO_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[FLOA_SAL_DESCUENTO_1] [float] NULL ,"
                           Cadena = Cadena + " [FLOA_SAL_DESCUENTO_2] [float] NULL ,[INTE_SAL_AÑO] [int] NULL , [INTE_SAL_CONSECUTIVO] [int] NULL) ON [PRIMARY]"
                           rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                           rsaux3.Open "INSERT INTO " + var_nombre_tabla + " (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, CHAR_PED_TIPO, INTE_SAL_CONSECUTIVO) select VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, CHAR_PED_TIPO, INTE_SAL_CONSECUTIVO from tb_temporal_salidas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id ='" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux3.Close
                        End If
                          
                        rsaux.Open "select * from " + var_nombre_tabla + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id ='" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux.EOF
                              valor = rsaux!VCHA_ART_ARTICULO_ID
                              var_n = lv_salidas.ListItems.Count
                              var_encontro = 0
                              var_i = 1
                              var_tipo_pedido = rsaux!char_ped_tipo
                              While (var_i <= var_n)
                                  lv_salidas.ListItems.item(var_i).Selected = True
                                  If valor = lv_salidas.selectedItem And var_tipo_pedido = lv_salidas.selectedItem.SubItems(11) Then
                                     var_encontro = 1
                                     var_i = var_n + 1
                                  Else
                                     var_encontro = 0
                                  End If
                                  var_i = var_i + 1
                              Wend
                              lv_salidas.selectedItem.SubItems(5) = Format(rsaux!floa_Sal_Cantidad, "###,###,##0.00")
                              rsaux.MoveNext
                        Wend
                        rsaux.Close
                        rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 1 WHERE vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + txt_busqueda_folio + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                        lbl_enviados = Format(Str(var_suma_cantidad_enviada), "###,###,##0.00")
                        lbl_recibidos = Format(Str(var_suma_cantidad_recibida), "###,###,##0.00")
                        txt_archivo.Enabled = False
                     Else
'''' nuevo metodo
                        var_archivo_tabla = Trim(var_empresa) + Trim(var_unidad_organizacional) + Trim(var_almacen_origen) + Trim(var_clave_movimiento) + Trim(CStr(var_numero_folio))
                        If var_estatus_movimiento = "I" Then
                           rsaux.Open "SELECT MAX(INTE_SAL_CONSECUTIVO) FROM TB_TEMPORAL_SALIDAS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id ='" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux.EOF Then
                              var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                           Else
                              var_consecutivo = 0
                           End If
                           rsaux.Close
                           If rsaux4.State = 1 Then
                              rsaux4.Close
                           End If

                           var_i = 1
                           For var_i = 1 To lv_salidas.ListItems.Count
                              lv_salidas.ListItems.item(var_i).Selected = True
                              If var_tipo_lectura = 1 Then
                                 Cadena = "insert into tb_salidas (VCHA_SAL_ARCHIVO, INTE_PED_NUMERO, INTE_ORS_ORDEN_SURTIDO, VCHA_EMP_EMPRESA_ID, INTE_SAL_NUMERO,VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ORS_CANTIDAD_SURTIR, FLOA_ORS_CANTIDAD_SURTIDA, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, "
                                 Cadena = Cadena + " VCHA_SAL_TIPO, INTE_SAL_CONSECUTIVO) VALUES ('" + Trim(var_empresa) + Trim(var_unidad_organizacional) + Trim(var_almacen_origen) + Trim(var_clave_movimiento) + Trim(CStr(var_numero_folio)) + "'," + Trim(txt_pedido) + "," + txt_archivo + ",'" + var_empresa + "'," + Trim(CStr(var_numero_folio)) + ",'" + lv_salidas.selectedItem + "',''," + lv_salidas.selectedItem.SubItems(2) + ", " + lv_salidas.selectedItem.SubItems(3) + ",0," + lv_salidas.selectedItem.SubItems(7) + "," + lv_salidas.selectedItem.SubItems(8) + "," + lv_salidas.selectedItem.SubItems(9) + "," + lv_salidas.selectedItem.SubItems(10) + ",'" + lv_salidas.selectedItem.SubItems(11) + "',0)"
                                 rsaux4.Open Cadena, cnnaccess, adOpenDynamic, adLockOptimistic
                              End If
                           Next var_i
                           rsaux.Open "select * from tb_temporal_salidas with (nolock) where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id ='" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                           While Not rsaux.EOF
                                 'rsaux4.Open "update tb_salidas set floa_sal_cantidad = " + CStr(rsaux!floa_sal_cantidad) + " where vcha_sal_archivo = '" + Trim(var_empresa) + Trim(var_unidad_organizacional) + Trim(var_almacen_origen) + Trim(var_clave_movimiento) + Trim(CStr(var_numero_folio)) + "' and vcha_art_articulo_id = '" + rsaux!VCHA_ART_aRTICULO_ID + "'", cnnaccess, adOpenDynamic, adLockOptimistic
                                 rsaux.MoveNext
                           Wend
                           rsaux.Close
                        Else
                           rsaux.Open "SELECT MAX(INTE_SAL_CONSECUTIVO) FROM TB_SALIDAS where vcha_sal_archivo = '" + var_archivo_tabla + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnnaccess, adOpenDynamic, adLockOptimistic
                           If Not rsaux.EOF Then
                              var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                           Else
                              var_consecutivo = 0
                           End If
                           rsaux.Close
                           If rsaux4.State = 1 Then
                              rsaux4.Close
                           End If
                        End If
                        rsaux.Open "select * from tb_salidas where vcha_sal_archivo = '" + var_archivo_tabla + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnnaccess, adOpenDynamic, adLockOptimistic
                        While Not rsaux.EOF
                              valor = rsaux!VCHA_ART_ARTICULO_ID
                              var_n = lv_salidas.ListItems.Count
                              var_encontro = 0
                              var_i = 1
                              var_tipo_pedido = rsaux!vcha_sal_tipo
                              While (var_i <= var_n)
                                  lv_salidas.ListItems.item(var_i).Selected = True
                                  If valor = lv_salidas.selectedItem Then
                                     var_encontro = 1
                                     var_i = var_n + 1
                                  Else
                                     var_encontro = 0
                                  End If
                                  var_i = var_i + 1
                              Wend
                              lv_salidas.selectedItem.SubItems(5) = Format(rsaux!floa_Sal_Cantidad, "###,###,##0.00")
                              rsaux.MoveNext
                        Wend
                        rsaux.Close
                        
                        If var_estatus_movimiento = "I" Then
                           'rsaux5.Open "delete from tb_salidas where vcha_sal_archivo = '" + var_archivo_tabla + "'", cnnaccess, adOpenDynamic, adLockOptimistic
                        End If
                        
                        rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 1 WHERE vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + txt_busqueda_folio + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                        lbl_enviados = Format(Str(var_suma_cantidad_enviada), "###,###,##0.00")
                        lbl_recibidos = Format(Str(var_suma_cantidad_recibida), "###,###,##0.00")
                        txt_archivo.Enabled = False
                     End If
                  Else
                     MsgBox "Numero de Orden de surtido no existe", vbOKOnly, "ATENCION"
                  End If
                  rs.Close
                  var_renglon = -1
                  Call ilumina_grid
                  frm_movimientos.Visible = False
                  If txt_codigo.Enabled = True Then
                     txt_codigo.SetFocus
                  End If
               Else
                  rs.Close
                  MsgBox "El movimiento esta siendo usado por otro usuario", vbOKOnly, "ATENCION"
                  frm_movimientos.Visible = False
               End If
            Else
               rs.Close
               MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
               frm_movimientos.Visible = False
            End If
         End If
      End If
      var_n = lv_salidas.ListItems.Count
      var_numero_renglones = lv_salidas.Height / 312.5
      If var_n > var_numero_renglones Then
         lv_salidas.ColumnHeaders(2).Width = 4100.15
      Else
         lv_salidas.ColumnHeaders(2).Width = 4300.15
      End If
   End If
   If KeyAscii = 27 Then
      frm_movimientos.Visible = False
      If txt_codigo.Enabled = True Then
         txt_codigo.SetFocus
      End If
   End If
   Exit Sub
no_almacen:
    MsgBox "Almacen Incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_Pedido:
    MsgBox "Pedido Incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_establecimiento:
    MsgBox "Establecimiento Incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_agente:
    MsgBox "Agente Incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_cliente:
    MsgBox "Cliente Incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_ruta:
    MsgBox "Ruta Incorrecta", vbOKOnly, "ATENCION"
    Exit Sub
no_titular:
    MsgBox "Titular incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_almacen_agente:
    MsgBox "No existe un almacen relacionado con este agente", vbOKOnly, "ATENCION"
    Exit Sub

End Sub

Private Sub lv_movimientos_LostFocus()
   frm_movimientos.Visible = False
End Sub

Private Sub lv_resumen_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_resumen, ColumnHeader)
End Sub

Private Sub lv_resumen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_resumen.Visible = False
   End If
End Sub

Private Sub lv_resumen_LostFocus()
   frm_resumen.Visible = False
End Sub

Private Sub lv_salidas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_salidas, ColumnHeader)
End Sub

Private Sub lv_salidas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         MsgBox "Imporsible modificar el movimiento", vbOKOnly, "ATENCION"
      Else
         frm_eliminar.Visible = True
         txt_cantidad_eliminar.SetFocus
      End If
   End If
End Sub


Private Sub lv_sellos_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      rs.Open "delete from tb_sellos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + CStr(var_numero_embarque) + " and vcha_sel_sello = '" + lv_sellos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
      lv_sellos.ListItems.Remove (lv_sellos.selectedItem.Index)
   End If
End Sub

Private Sub lv_sellos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_sellos.Visible = False
   End If
End Sub

Private Sub txt_archivo_KeyPress(KeyAscii As Integer)
   Dim var_clave_movimiento_tem As String
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If IsNumeric(txt_archivo) Then
         ejecuta
         var_renglon = -1
         Call ilumina_grid
      Else
         MsgBox "Embarque incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
Dim var_busqueda_folio As Double
Dim var_busqueda_numero As Double
Dim var_busqueda_referencia As String
Dim var_posible As Boolean
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      rs.Open "select * from tb_detalle_embarques where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + txt_busqueda_folio + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         If var_numero_embarque = rs!inte_emb_embarque Then
            var_posible = True
         Else
            MsgBox "Este movimiento se encuentra en el embarque No. " + Str(rs!inte_emb_embarque), vbOKOnly, "ATENCION"
            var_posible = False
            frm_busqueda.Visible = False
         End If
      Else
         MsgBox "El Movimiento no existe", vbOKOnly, "ATENCION"
         var_posible = False
         frm_busqueda.Visible = False
      End If
      rs.Close
      If var_posible = True Then
         rs.Open "select * from tb_detalle_cajas with (nolock) where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_paq_movimiento_destino = '" + var_clave_movimiento + "' and inte_paq_numero_destino = " + txt_busqueda_folio, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_cajas = True
         Else
            var_cajas = False
         End If
         rs.Close
         If var_numero_folio = CDbl(txt_busqueda_folio) Then
            rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
         rs.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + txt_busqueda_folio + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If var_numero_folio > 0 Then
               rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            var_movimiento_bloqueado = IIf(IsNull(rs!INTE_EMO_BLOQUEADO), 0, rs!INTE_EMO_BLOQUEADO)
            If var_movimiento_bloqueado = 0 Then
               var_primera_vez = False
               var_factura_ceros = IIf(IsNull(rs!inte_emo_factura_ceros), 0, rs!inte_emo_factura_ceros)
               var_clave_moneda = rs!vcha_mon_moneda_id
               var_orden_surtido = rs!inte_emo_numero_origen
               var_numero_folio = rs!INTE_EMO_NUMERO
               txt_folio = rs!INTE_EMO_NUMERO
               var_almacen_origen = rs!VCHA_ALM_ALMACEN_ID
               rsaux3.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               txt_origen = rsaux3!VCHA_ALM_NOMBRE
               rsaux3.Close
               If IsNull(rs!char_Emo_estatus) Then
                  var_estatus_movimiento = ""
               Else
                  var_estatus_movimiento = rs!char_Emo_estatus
               End If
               If rs!char_Emo_estatus = "I" Then
                  txt_codigo.Enabled = False
               Else
                  txt_codigo.Enabled = True
               End If
               rs.Close
               rs.Open "select * from vw_orden_surtido where inte_ors_orden_surtido = " + Str(var_orden_surtido) + " and floa_ors_cantidad_surtir > 0", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  txt_archivo = var_orden_surtido
                  var_suma_cantidad_enviada = 0
                  var_suma_cantidad_recibida = 0
                  lbl_enviados.Caption = "0"
                  lbl_recibidos.Caption = "0"
                  lv_salidas.ListItems.Clear
                  If IsNull(rs!VCHA_TIT_NOMBRE) Then
                     GoTo no_titular:
                  Else
                     txt_titular = rs!VCHA_TIT_NOMBRE
                     var_clave_titular = rs!vcha_tit_titular_id
                  End If
                  If IsNull(rs!inte_ped_dias_condiciones) Then
                     var_plazo = 0
                  Else
                     var_plazo = rs!inte_ped_dias_condiciones
                  End If
                  If IsNull(rs!vcha_cli_email) Then
                     var_correo_electronico = ""
                  Else
                     var_correo_electronico = rs!vcha_cli_email
                  End If
                  If IsNull(rs!VCHA_ESB_NOMBRE) Then
                     GoTo no_establecimiento:
                  Else
                     txt_establecimiento = rs!VCHA_ESB_NOMBRE
                     var_clave_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
                  End If
                  If IsNull(rs!VCHA_AGE_NOMBRE) Then
                     GoTo no_agente:
                  Else
                     txt_agente = rs!VCHA_AGE_NOMBRE
                     var_clave_agente = rs!VCHA_AGE_AGENTE_ID
                  End If
                  var_almacen_Destino = ""
                  If var_tipo_documento = "V" Then
                     If IsNull(rs!almacen_agente) Then
                        GoTo no_almacen_agente:
                     Else
                        var_almacen_Destino = rs!almacen_agente
                     End If
                  End If
                  If IsNull(rs!VCHA_CLI_NOMBRE) Then
                     GoTo no_cliente:
                  Else
                     txt_cliente = rs!VCHA_CLI_NOMBRE
                     var_clave_cliente = rs!vcha_cli_clave_id
                  End If
                  If IsNull(rs!vcha_rut_nombre) Then
                     txt_ruta = ""
                     var_clave_ruta = ""
                  Else
                     txt_ruta = rs!vcha_rut_nombre
                     var_clave_ruta = rs!vcha_rut_ruta_id
                  End If
                  If IsNull(rs!inte_ped_numero) Then
                     GoTo no_Pedido:
                  Else
                     txt_pedido = rs!inte_ped_numero
                  End If
                  If IsNull(rs!FLOA_ORS_DESCUENTO_1) Then
                     txt_descuento1 = 0
                     var_descuento_1 = 0
                  Else
                     txt_descuento1 = rs!FLOA_ORS_DESCUENTO_1
                     var_descuento_1 = rs!FLOA_ORS_DESCUENTO_1
                  End If
                  If IsNull(rs!FLOA_ORS_DESCUENTO_2) Then
                     txt_descuento2 = 0
                     var_descuento_2 = 0
                  Else
                     txt_descuento2 = rs!FLOA_ORS_DESCUENTO_2
                     var_descuento_2 = rs!FLOA_ORS_DESCUENTO_2
                  End If
                  var_descuento_3 = 0
                  While Not rs.EOF
                     Set list_item = lv_salidas.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
                         list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_nombre_español), "", Trim(rs!vcha_Art_nombre_español))
                         var_surtir = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR)
                         list_item.SubItems(2) = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0.00")
                         var_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
                         list_item.SubItems(3) = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA), "###,###,##0.00")
                         var_empacada = IIf(IsNull(rs!floa_ors_Cantidad_empacada), 0, rs!floa_ors_Cantidad_empacada)
                         list_item.SubItems(4) = Format(IIf(IsNull(rs!floa_ors_Cantidad_empacada), 0, rs!floa_ors_Cantidad_empacada), "###,###,##0.00")
                         list_item.SubItems(5) = Format(0, "###,###,##0.00")
                         var_falta = var_surtida + var_empacada
                         list_item.SubItems(6) = Format(var_surtir - var_falta, "###,###,##0.00")
                         list_item.SubItems(7) = IIf(IsNull(rs!floa_ors_costo), "", rs!floa_ors_costo)
                         'If var_factura_ceros = 1 Then
                            list_item.SubItems(8) = 0
                         'Else
                         '   list_item.SubItems(8) = IIf(IsNull(rs!floa_ors_precio), "", rs!floa_ors_precio)
                         'End If
                         list_item.SubItems(11) = IIf(IsNull(rs!char_ped_tipo), "P", rs!char_ped_tipo)
                         var_suma_cantidad_enviada = var_suma_cantidad_enviada + rs!FLOA_ORS_CANTIDAD_SURTIR
                         var_suma_cantidad_recibida = var_suma_cantidad_recibida + rs!FLOA_ORS_CANTIDAD_SURTIDA
                      rs.MoveNext:
                  Wend
                  rsaux.Open "SELECT MAX(INTE_SAL_CONSECUTIVO) FROM TB_TEMPORAL_SALIDAS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id ='" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                  Else
                     var_consecutivo = 0
                  End If
                  rsaux.Close
                  rsaux.Open "select * from tb_temporal_salidas with (nolock) where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id ='" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux.EOF
                     valor = rsaux!VCHA_ART_ARTICULO_ID
                     'Set itmfound = lv_salidas.FindItem(valor, lvwText, , lvwPartial)
                     'itmfound.EnsureVisible
                     'itmfound.Selected = True
                     var_n = lv_salidas.ListItems.Count
                     var_encontro = 0
                     var_i = 1
                     var_tipo_pedido = rsaux!char_ped_tipo
                     While (var_i <= var_n)
                         lv_salidas.ListItems.item(var_i).Selected = True
                         'If var_cantidad_posible < lv_salidas.SelectedItem.SubItems(3) + var_cantidad_leida Then
                         If valor = lv_salidas.selectedItem And var_tipo_pedido = lv_salidas.selectedItem.SubItems(11) Then
                            var_encontro = 1
                            var_i = var_n + 1
                         Else
                            var_encontro = 0
                         End If
                         var_i = var_i + 1
                     Wend
                     lv_salidas.selectedItem.SubItems(5) = Format(rsaux!floa_Sal_Cantidad, "###,###,##0.00")
                     rsaux.MoveNext
                  Wend
                  rsaux.Close
                  rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 1 WHERE vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + txt_busqueda_folio + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  lbl_enviados = Format(Str(var_suma_cantidad_enviada), "###,###,##0.00")
                  lbl_recibidos = Format(Str(var_suma_cantidad_recibida), "###,###,##0.00")
                  txt_archivo.Enabled = False
               Else
                  MsgBox "Numero de Orden de surtido no existe", vbOKOnly, "ATENCION"
               End If
               rs.Close
               frm_busqueda.Visible = False
            Else
               rs.Close
               MsgBox "El movimiento esta siendo usado por otro usuario", vbOKOnly, "ATENCION"
            End If
         Else
            rs.Close
            MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
            frm_busqueda.Visible = False
         End If
      End If
   End If
   var_n = lv_salidas.ListItems.Count
   var_numero_renglones = lv_salidas.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_salidas.ColumnHeaders(2).Width = 4100.15
   Else
      lv_salidas.ColumnHeaders(2).Width = 4300.15
   End If
   If KeyAscii = 27 Then
      frm_busqueda.Visible = False
   End If
   Exit Sub
no_almacen:
    MsgBox "Almacen Incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_Pedido:
    MsgBox "Pedido Incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_establecimiento:
    MsgBox "Establecimiento Incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_agente:
    MsgBox "Agente Incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_cliente:
    MsgBox "Cliente Incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_ruta:
    MsgBox "Ruta Incorrecta", vbOKOnly, "ATENCION"
    Exit Sub
no_titular:
    MsgBox "Titular incorrecto", vbOKOnly, "ATENCION"
    Exit Sub
no_almacen_agente:
    MsgBox "No existe un almacen relacionado con este agente", vbOKOnly, "ATENCION"
    Exit Sub
End Sub

Private Sub txt_busqueda_folio_LostFocus()
      frm_busqueda.Visible = False
End Sub

Private Sub txt_cantidad_eliminar_GotFocus()
   txt_cantidad_eliminar = ""
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   Dim var_embarque_paquete As Integer
   Dim var_embarque_caja As Integer
   Dim var_encontrado As Integer
   Dim var_i As Integer
   Dim var_n As Integer
   Dim var_j As Integer
   Dim var_precio As Variant
   Dim var_tipo_pedido As String
   Dim var_posible As Boolean
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If var_cajas = True Then
         var_es_caja = False
         If Trim(txt_cantidad_eliminar) <> "" Then
            If Left(Trim(txt_cantidad_eliminar), 1) = "C" Then
               var_es_caja = True
            Else
               var_es_caja = False
            End If
            If var_es_caja = True Then
               x = Mid(txt_cantidad_eliminar, 2, 6)
               If IsNumeric(x) Then
                  var_embarque_paquete = x
                  x = Mid(txt_cantidad_eliminar, 8, 3)
                  If IsNumeric(x) Then
                     var_embarque_caja = x
                     var_posible_caja = True
                  Else
                     var_posible_caja = False
                  End If
               Else
                  var_posible_caja = False
               End If
               If var_posible_caja = True Then
                  var_embarque_paquete = txt_embarque
                  rsaux3.Open "select * from tb_detalle_cajas with (nolock) where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and INTE_EMB_EMBARQUE = " + CStr(var_numero_embarque) + " and inte_paq_caja = " + Str(var_embarque_caja), cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     If rsaux3!char_paq_estatus = "S" Then
                        Set TB_DETALLE_CAJAS_M = New TB_DETALLE_CAJAS_M
                        ok = False
                        ok = TB_DETALLE_CAJAS_M.Anadir(txt_archivo, var_embarque_caja, var_empresa, var_unidad_organizacional, var_almacen_origen, "I", "", 0, var_numero_embarque)
                        While Not rsaux3.EOF
                           Set TB_DET_ORDEN_SURTIDO_M = New TB_DET_ORDEN_SURTIDO_M
                           Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
                           valor = rsaux3!VCHA_ART_ARTICULO_ID
                           var_precio = rsaux3!floa_paq_precio
                           var_tipo_pedido = rsaux3!char_ped_tipo
                           var_n = lv_salidas.ListItems.Count
                           var_encontro = 0
                           var_i = 1
                           While (var_i <= var_n)
                                 lv_salidas.ListItems.item(var_i).Selected = True
                                 If (lv_salidas.selectedItem.SubItems(8) * 1) = var_precio And lv_salidas.selectedItem = valor And lv_salidas.selectedItem.SubItems(11) = var_tipo_pedido Then
                                    var_encontro = 1
                                    var_i = var_n + 1
                                 End If
                                 var_i = var_i + 1
                           Wend
                           
                           
                           var_cantidad_eliminar = rsaux3!floa_paq_cantidad
                           var_cantidad_eliminar_arch = lv_salidas.selectedItem.SubItems(3) - var_cantidad_eliminar
                           var_cantidad_eliminar_mov = lv_salidas.selectedItem.SubItems(5) - var_cantidad_eliminar
                           If var_cantidad_eliminar_arch < 0 Or var_cantidad_eliminar_mov < 0 Then
                              MsgBox "No esposible eliminar esta cantidad", vbOKOnly, "ATENCION"
                           Else
                              var_precio = lv_salidas.selectedItem.SubItems(8)
                              var_actualiza = TB_DET_ORDEN_SURTIDO_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_OS, var_orden_surtido, lv_salidas.selectedItem, 0 - var_cantidad_eliminar, var_cantidad_eliminar, var_precio, var_tipo_pedido)
                              lv_salidas.selectedItem.SubItems(3) = Format(lv_salidas.selectedItem.SubItems(3) - var_cantidad_eliminar, "###,###,##0.00")
                              lv_salidas.selectedItem.SubItems(4) = Format(lv_salidas.selectedItem.SubItems(4) + var_cantidad_eliminar, "###,###,##0.00")
                              lv_salidas.selectedItem.SubItems(5) = Format(lv_salidas.selectedItem.SubItems(5) - var_cantidad_eliminar, "###,###,##0.00")
                              lv_salidas.selectedItem.SubItems(6) = Format(lv_salidas.selectedItem.SubItems(2) - lv_salidas.selectedItem.SubItems(3) - lv_salidas.selectedItem.SubItems(4), "###,###,##0.00")
                              rsaux.Open "update " + var_nombre_tabla + " set floa_sal_cantidad = floa_sal_cantidad -" + CStr(var_cantidad_eliminar) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_Sal_Numero = " + CStr(var_numero_folio) + " and vcha_art_articulo_id = '" + lv_salidas.selectedItem + "' and round(floa_sal_precio,2) = round(" + CStr(var_precio) + ",2) AND CHAR_PED_TIPO = '" + var_tipo_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
                              lbl_recibidos = Format(CDbl(lbl_recibidos) - var_cantidad_eliminar, "###,###,##0.00")
                              frm_eliminar.Visible = False
                              txt_codigo.SetFocus
                           End If
                           rsaux3.MoveNext
                        Wend
                     End If
                  End If
                  rsaux3.Close
               End If
            End If
         End If
      Else
         var_archivo_tabla = Trim(var_empresa) + Trim(var_unidad_organizacional) + Trim(var_almacen_origen) + Trim(var_clave_movimiento) + Trim(CStr(var_numero_folio))
         Set TB_DET_ORDEN_SURTIDO_M = New TB_DET_ORDEN_SURTIDO_M
         Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
         var_posible = True
         If var_posible_kanban = 1 Then
            If IsNumeric(Me.txt_cantidad_eliminar) Then
               Set TB_CANCELAR_RES_FUERA_DE_KANBAN = New TB_CANCELAR_RES_FUERA_DE_KANBAN
               var_inserta = TB_CANCELAR_RES_FUERA_DE_KANBAN.Anadir(var_almacen_origen, var_clave_movimiento, var_numero_folio, Me.lv_salidas.selectedItem, CDbl(Me.txt_cantidad_eliminar), "", "")
               var_kanban_es_un_kanban = var_kanban_es_un_kanban
               var_kanban_almacen_id = var_kanban_almacen_id
               var_kanban_articulo_id = var_kanban_articulo_id
               var_kanban_exito = var_kanban_exito
               var_kanban_mensaje = var_kanban_mensaje
               If var_kanban_exito = "S" Then
                  var_posible = True
               Else
                  Call cmd_mensaje_1_Click
                  frmmensaje.lbl_mensaje = var_kanban_mensaje
                  frmmensaje.Show 1
                  var_posible = False
               End If
               
            Else
            
               Set TB_ES_UN_KANBAN = New TB_ES_UN_KANBAN
               var_kanban = Me.txt_codigo
               var_inserta = TB_ES_UN_KANBAN.Anadir(Me.txt_cantidad_eliminar, "", "", "", "", "")
               var_kanban_es_un_kanban = var_kanban_es_un_kanban
               var_kanban_almacen_id = var_kanban_almacen_id
               var_kanban_articulo_id = var_kanban_articulo_id
               var_kanban_exito = var_kanban_exito
               var_kanban_mensaje = var_kanban_mensaje
               
               If var_kanban_es_un_kanban = "S" Then
                  If lv_salidas.selectedItem = var_kanban_articulo_id Then
                     Set TB_CANCELAR_RESERVACION_KANBAN = New TB_CANCELAR_RESERVACION_KANBAN
                     var_kanban = Me.txt_codigo
                     var_inserta = TB_CANCELAR_RESERVACION_KANBAN.Anadir(var_almacen_origen, var_clave_movimiento, var_numero_folio, Me.txt_cantidad_eliminar, "", "")
                     var_kanban_es_un_kanban = var_kanban_es_un_kanban
                     var_kanban_almacen_id = var_kanban_almacen_id
                     var_kanban_articulo_id = var_kanban_articulo_id
                     var_kanban_exito = var_kanban_exito
                     var_kanban_mensaje = var_kanban_mensaje
                     If var_kanban_exito = "S" Then
                        var_posible = True
                     Else
                        Call cmd_mensaje_1_Click
                        frmmensaje.lbl_mensaje = var_kanban_mensaje
                        frmmensaje.Show 1
                        var_posible = False
                     End If
                  Else
                     Call cmd_mensaje_1_Click
                     frmmensaje.lbl_mensaje = "El codigo de kanban no corresponde al del artículo seleccionado"
                     frmmensaje.Show 1
                     var_posible = False
                  End If
               Else
                  Call cmd_mensaje_1_Click
                  frmmensaje.lbl_mensaje = var_kanban_mensaje
                  frmmensaje.Show 1
                  var_posible = False
               End If
            End If
         Else
            var_posible = True
         End If
         
         If var_posible = True Then
            If var_posible_kanban = 1 Then
               If Not IsNumeric(txt_cantidad_eliminar) Then
                  Me.txt_cantidad_eliminar = 1
               End If
            End If
            var_cantidad_eliminar = CDbl(txt_cantidad_eliminar)
            var_cantidad_eliminar_arch = lv_salidas.selectedItem.SubItems(3) - CDbl(txt_cantidad_eliminar)
            var_cantidad_eliminar_mov = lv_salidas.selectedItem.SubItems(5) - CDbl(txt_cantidad_eliminar)
            If var_cantidad_eliminar_arch < 0 Or var_cantidad_eliminar_mov < 0 Then
               MsgBox "No esposible eliminar esta cantidad", vbOKOnly, "ATENCION"
            Else
               If var_tipo_lectura = 0 Then
                  var_tipo_pedido = lv_salidas.selectedItem.SubItems(11)
                  var_precio = lv_salidas.selectedItem.SubItems(8)
                  var_actualiza = TB_DET_ORDEN_SURTIDO_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_orden_surtido, lv_salidas.selectedItem, 0 - var_cantidad_eliminar, 0, var_precio, var_tipo_pedido)
                  lv_salidas.selectedItem.SubItems(3) = Format(lv_salidas.selectedItem.SubItems(3) - CDbl(txt_cantidad_eliminar), "###,###,##0.00")
                  lv_salidas.selectedItem.SubItems(5) = Format(lv_salidas.selectedItem.SubItems(5) - CDbl(txt_cantidad_eliminar), "###,###,##0.00")
                  lv_salidas.selectedItem.SubItems(6) = Format(lv_salidas.selectedItem.SubItems(2) - lv_salidas.selectedItem.SubItems(3) - lv_salidas.selectedItem.SubItems(4), "###,###,##0.00")
                  rsaux5.Open "update TB_DETALLE_EQUIPOS_ORDEN_SURTIDO set FLOA_ORS_CANTIDAD_SURTIDA = isnull(FLOA_ORS_CANTIDAD_SURTIDA,0) - " + CStr(var_cantidad_eliminar) + " where INTE_ORS_ORDEN_SURTIDO = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
                  var_renglon = lv_salidas.selectedItem.Index
                  rsaux.Open "update " + var_nombre_tabla + " set floa_sal_cantidad = floa_sal_cantidad -" + CStr(var_cantidad_eliminar) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_Sal_Numero = " + CStr(var_numero_folio) + " and vcha_art_articulo_id = '" + lv_salidas.selectedItem + "' and round(floa_sal_precio,2) = round(" + CStr(var_precio) + ",2) AND CHAR_PED_TIPO= '" + var_tipo_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
                  Call ilumina_grid
                  lbl_recibidos = Format(CDbl(lbl_recibidos) - var_cantidad_eliminar, "###,###,##0.00")
                  frm_eliminar.Visible = False
                  txt_codigo.SetFocus
               Else
                  var_tipo_pedido = lv_salidas.selectedItem.SubItems(11)
                  var_precio = lv_salidas.selectedItem.SubItems(8)
                  'var_actualiza = TB_DET_ORDEN_SURTIDO_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_orden_surtido, lv_salidas.selectedItem, 0 - var_cantidad_eliminar, 0, var_precio, var_tipo_pedido)
                  '26-09-08 para afectar orden surtido hasta que se cierra el embarque
                  rsaux5.Open "update tb_det_orden_surtido set floa_ors_cantidad_surtida = floa_ors_cantidad_surtida - " + CStr(var_cantidad_eliminar) + " where inte_ors_orden_surtido = " + CStr(var_orden_surtido) + " and vcha_art_articulo_id = '" + lv_salidas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                  
                  rsaux5.Open "update TB_DETALLE_EQUIPOS_ORDEN_SURTIDO set FLOA_ORS_CANTIDAD_SURTIDA = isnull(FLOA_ORS_CANTIDAD_SURTIDA,0) - " + CStr(var_cantidad_eliminar) + " where INTE_ORS_ORDEN_SURTIDO = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
                  lv_salidas.selectedItem.SubItems(3) = Format(lv_salidas.selectedItem.SubItems(3) - CDbl(txt_cantidad_eliminar), "###,###,##0.00")
                  lv_salidas.selectedItem.SubItems(5) = Format(lv_salidas.selectedItem.SubItems(5) - CDbl(txt_cantidad_eliminar), "###,###,##0.00")
                  lv_salidas.selectedItem.SubItems(6) = Format(lv_salidas.selectedItem.SubItems(2) - lv_salidas.selectedItem.SubItems(3) - lv_salidas.selectedItem.SubItems(4), "###,###,##0.00")
                  var_renglon = lv_salidas.selectedItem.Index
                  rsaux.Open "update tb_salidas set FLOA_ORS_CANTIDAD_SURTIDA = FLOA_ORS_CANTIDAD_SURTIDA - " + CStr(var_cantidad_eliminar) + ", floa_sal_cantidad  = floa_sal_cantidad - " + CStr(var_cantidad_eliminar) + " where vcha_sal_archivo  = '" + var_archivo_tabla + "' and vcha_art_articulo_id = '" + lv_salidas.selectedItem + "' and inte_ors_orden_surtido = " + txt_archivo, cnnaccess, adOpenDynamic, adLockOptimistic
                  Call ilumina_grid
                  lbl_recibidos = Format(CDbl(lbl_recibidos) - var_cantidad_eliminar, "###,###,##0.00")
                  frm_eliminar.Visible = False
                  txt_codigo.SetFocus
               End If
            End If
         End If
      End If
      frm_eliminar.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_eliminar.Visible = False
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   frm_eliminar.Visible = False
   If Me.txt_codigo.Enabled = True Then
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_cantidad_GotFocus()
   txt_cantidad = ""
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_cantidad) <> "" Then
         var_cantidad_leida = txt_cantidad
         txt_foco.Enabled = True
         txt_foco.SetFocus
         lbl_cantidad.Visible = False
         txt_cantidad.Visible = False
      End If
   End If
End Sub

Private Sub txt_codigo_GotFocus()
   var_cantidad_multibondeados = 0
   var_nombre_articulo_mensaje = ""
   txt_codigo = ""
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
   Dim var_caja As String
   Dim var_cantidad_caja As Integer
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If var_posible_kanban = 1 Then
      If var_empresa <> "16" Then
         If KeyAscii = 39 Or KeyAscii = 61 Then
            KeyAscii = 0
         End If
      End If
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      If KeyAscii = 13 Then
         If var_empresa = "16" Then
            'If Mid(Me.txt_codigo, 1, 3) <> "TMG" Then
            '   var_cadena_X = ""
            '   For var_jj = 1 To Len(Me.txt_codigo)
            '       If Mid(Me.txt_codigo, var_jj, 1) = "'" Then
            '          var_cadena_X = var_cadena_X + "-"
            '       Else
            '          var_cadena_X = var_cadena_X + Mid(Me.txt_codigo, var_jj, 1)
            '       End If
            '   Next var_jj
            '   Me.txt_codigo = var_cadena_X
           '
            'End If
         End If
         var_cantidad_multibondeados = 0
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
         'If var_empresa = 16 Then
         '   If Len(Me.txt_codigo) = 6 Then
         '      'Me.txt_codigo = Mid(Me.txt_codigo, 1, 3) + "-" + Mid(Me.txt_codigo, 4, 3) + "-"
         '   Else
         '      If Len(Me.txt_codigo) = 7 Then
         '         Me.txt_codigo = Mid(Me.txt_codigo, 1, 3) + "-" + Mid(Me.txt_codigo, 4, 3) + "-" + Mid(Me.txt_codigo, 7, 1)
         '      End If
         '   End If
         'End If
          
         var_verificador = True
         If Len(Trim(txt_codigo)) = 12 Then
            Call calcula_verificador(Trim(txt_codigo))
         End If
         If var_empresa = "31" Then
            var_verificador = True
         End If
         If var_verificador = True Then
            var_es_caja = False
            If Trim(txt_codigo) <> "" Then
               If Left(Trim(txt_codigo), 1) = "C" Then
                  x = Mid(txt_codigo, 2, 6)
                  var_embarque_caja = 0
                  If IsNumeric(x) Then
                     var_embarque_caja = CDbl(x)
                     If var_embarque_caja = var_numero_embarque Then
                        var_es_caja = True
                     Else
                        frmmensaje.lbl_mensaje = "La caja pertenece a otro embarque"
                        frmmensaje.Show 1
                        'MsgBox "La caja pertenece al embarque " + CStr(var_embarque_caja)
                        var_es_caja = False
                     End If
                  Else
                     frmmensaje.lbl_mensaje = "Caja incorrecta"
                     frmmensaje.Show 1
                     'MsgBox "Caja incorrecta", vbOKOnly, "ATENCION"
                     var_es_caja = False
                  End If
               Else
                  var_es_caja = False
               End If
               If var_es_caja = True Then
                  txt_foco.Enabled = True
                  txt_foco.SetFocus
               Else
                  var_caja = Left(txt_codigo, 6)
                  If var_caja = "000005" Or var_caja = "000010" Or var_caja = "000015" Or var_caja = "000020" Or var_caja = "000001" Or var_caja = "000002" Or var_caja = "000003" Or var_caja = "000004" Or var_caja = "000006" Or var_caja = "000007" Or var_caja = "000008" Or var_caja = "000009" Or var_caja = "000010" Or var_caja = "000011" Or var_caja = "000012" Or var_caja = "000013" Or var_caja = "000014" Or var_caja = "000015" Or var_caja = "000016" Or var_caja = "000017" Or var_caja = "000018" Or var_caja = "000019" Or var_caja = "000020" Or var_caja = "000021" Or var_caja = "000022" Or var_caja = "000023" Or var_caja = "000024" Or var_caja = "000025" Or var_caja = "000026" Or var_caja = "000027" Or var_caja = "000028" Or var_caja = "000029" Or var_caja = "000030" Then
                     var_cantidad_caja = CInt(var_caja)
                     txt_codigo = Mid(txt_codigo, 7, 5)
                  End If
                  rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_nombre_articulo_mensaje = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
                     If IsNull(rs(43).Value) Then
                        var_recontable = 0
                     Else
                        var_recontable = rs(43).Value
                     End If
                     rs.Close
                     If var_recontable = 1 Then
                        var_cantidad_leida = 1#
                        lbl_cantidad.Visible = True
                        txt_cantidad.Visible = True
                        txt_cantidad.SetFocus
                     Else
                        var_cantidad_leida = 1#
                        txt_foco.Enabled = True
                        txt_foco.SetFocus
                     End If
                  Else
                     rs.Close
                     rs.Open "select * from tb_equivalencias where VCHA_EQU_CODIGO_EQUIVALENTE = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        txt_codigo = rs(0).Value
                        rs.Close
                        rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_nombre_articulo_mensaje = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
                           If var_cantidad_caja = 0 Then
                              If IsNull(rs(43).Value) Then
                                 var_recontable = 0
                              Else
                                 var_recontable = rs(43).Value
                              End If
                           Else
                              var_recontable = 0
                           End If
                           rs.Close
                           If var_recontable = 1 Then
                              var_cantidad_leida = 1#
                              lbl_cantidad.Visible = True
                              txt_cantidad.Visible = True
                              txt_cantidad.SetFocus
                           Else
                              If var_cantidad_caja = 0 Then
                                 var_cantidad_leida = 1#
                              Else
                                 var_cantidad_leida = var_cantidad_caja
                              End If
                              txt_foco.Enabled = True
                              txt_foco.SetFocus
                           End If
                        Else
                           Call cmd_mensaje_1_Click
                           frmmensaje.lbl_mensaje = "El artículo no existe"
                           frmmensaje.Show 1
                           'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                           txt_codigo = ""
                        End If
                     Else
                        Call cmd_mensaje_1_Click
                        frmmensaje.lbl_mensaje = "El artículo no existe"
                        frmmensaje.Show 1
                       'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                        txt_codigo = ""
                        rs.Close
                     End If
                  End If
               End If
            End If
         Else
            txt_codigo = ""
            Call cmd_mensaje_1_Click
            frmmensaje.lbl_mensaje = "Error en Código"
            frmmensaje.Show 1
            ' MsgBox "Error en Código", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      txt_codigo = Trim(txt_codigo)
      If var_empresa <> "16" Then
         If KeyAscii = 39 Or KeyAscii = 61 Then
            KeyAscii = 0
         End If
      End If
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      If KeyAscii = 13 Then
         'If var_empresa = "16" Then
         '   If Mid(Me.txt_codigo, 1, 3) <> "TMG" Then
         '      var_cadena_X = ""
         '      For var_jj = 1 To Len(Me.txt_codigo)
         '          If Mid(Me.txt_codigo, var_jj, 1) = "'" Then
         '             var_cadena_X = var_cadena_X + "-"
         '          Else
         '             var_cadena_X = var_cadena_X + Mid(Me.txt_codigo, var_jj, 1)
         '          End If
         '      Next var_jj
         '      Me.txt_codigo = var_cadena_X
         '   End If
         '   'MsgBox Me.txt_codigo
         'End If
       
      
      
var_cantidad_multibondeados = 0
         'If var_empresa = 16 Then
         '  If Len(Me.txt_codigo) = 6 Then
         '     'Me.txt_codigo = Mid(Me.txt_codigo, 1, 3) + "-" + Mid(Me.txt_codigo, 4, 3) + "-"
         '  Else
         '     If Len(Me.txt_codigo) = 7 Then
         '        If Mid(Me.txt_codigo, 1, 3) <> "TMG" Then
         '           Me.txt_codigo = Mid(Me.txt_codigo, 1, 3) + "-" + Mid(Me.txt_codigo, 4, 3) + "-" + Mid(Me.txt_codigo, 7, 1)
         '        End If
         '     End If
         '  End If
         'End If
         If var_empresa = "16" Then
            If Mid(Me.txt_codigo, 1, 3) <> "TMG" Then
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
               'If UCase(x) = "B" Or UCase(x) = "R" Then
               '   If Len(Me.txt_codigo) > 9 Then
               '      var_c = Mid(Me.txt_codigo, 1, Len(Me.txt_codigo) - 9)
               '      If IsNumeric(var_c) Then
               '         var_cantidad_multibondeados = CDbl(var_c)
               '         Me.txt_codigo = Right(Me.txt_codigo, 9)
               '         'MsgBox Me.txt_codigo
               '      End If
               '   End If
               'End If
            End If
         End If
         var_verificador = True
         If Len(Trim(txt_codigo)) = 12 Then
            Call calcula_verificador(Trim(txt_codigo))
         End If
         If var_empresa = "31" Then
            var_verificador = True
         End If
         
         If var_verificador = True Then
            var_es_caja = False
            If Trim(txt_codigo) <> "" Then
               If Left(Trim(txt_codigo), 1) = "C" Then
                  x = Mid(txt_codigo, 2, 6)
                  var_embarque_caja = 0
                  If IsNumeric(x) Then
                     var_embarque_caja = CDbl(x)
                     If var_embarque_caja = var_numero_embarque Then
                        var_es_caja = True
                     Else
                        frmmensaje.lbl_mensaje = "La caja pertenece a otro embarque"
                        frmmensaje.Show 1
                        'MsgBox "La caja pertenece al embarque " + CStr(var_embarque_caja)
                        var_es_caja = False
                     End If
                  Else
                     frmmensaje.lbl_mensaje = "Caja incorrecta"
                     frmmensaje.Show 1
                     'MsgBox "Caja incorrecta", vbOKOnly, "ATENCION"
                     var_es_caja = False
                  End If
               Else
                  var_es_caja = False
               End If
               If var_es_caja = True Then
                  txt_foco.Enabled = True
                  txt_foco.SetFocus
               Else
                  var_caja = Left(txt_codigo, 6)
                  If var_caja = "000005" Or var_caja = "000010" Or var_caja = "000015" Or var_caja = "000020" Or var_caja = "000001" Or var_caja = "000002" Or var_caja = "000003" Or var_caja = "000004" Or var_caja = "000006" Or var_caja = "000007" Or var_caja = "000008" Or var_caja = "000009" Or var_caja = "000010" Or var_caja = "000011" Or var_caja = "000012" Or var_caja = "000013" Or var_caja = "000014" Or var_caja = "000015" Or var_caja = "000016" Or var_caja = "000017" Or var_caja = "000018" Or var_caja = "000019" Or var_caja = "000020" Or var_caja = "000021" Or var_caja = "000022" Or var_caja = "000023" Or var_caja = "000024" Or var_caja = "000025" Or var_caja = "000026" Or var_caja = "000027" Or var_caja = "000028" Or var_caja = "000029" Or var_caja = "000030" Then
                     var_cantidad_caja = CInt(var_caja)
                     txt_codigo = Mid(txt_codigo, 7, 5)
                  End If
                  rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_nombre_articulo_mensaje = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
                     If IsNull(rs(43).Value) Then
                        var_recontable = 0
                     Else
                        If var_cantidad_multibondeados = 0 Then
                           var_recontable = rs(43).Value
                        Else
                           var_recontable = 0
                        End If
                     End If
                     rs.Close
                     If var_recontable = 1 Then
                        var_cantidad_leida = 1#
                        lbl_cantidad.Visible = True
                        txt_cantidad.Visible = True
                        txt_cantidad.SetFocus
                     Else
                        If var_cantidad_multibondeados = 0 Then
                           var_cantidad_leida = 1#
                        Else
                           var_cantidad_leida = var_cantidad_multibondeados
                        End If
                        txt_foco.Enabled = True
                        txt_foco.SetFocus
                     End If
                  Else
                     rs.Close
                     rs.Open "select * from tb_equivalencias where VCHA_EQU_CODIGO_EQUIVALENTE = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        txt_codigo = rs(0).Value
                        rs.Close
                        rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_nombre_articulo_mensaje = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
                           If var_cantidad_caja = 0 Then
                              If IsNull(rs(43).Value) Then
                                 var_recontable = 0
                              Else
                                 var_recontable = rs(43).Value
                              End If
                           Else
                              var_recontable = 0
                           End If
                           rs.Close
                           If var_recontable = 1 Then
                              var_cantidad_leida = 1#
                              lbl_cantidad.Visible = True
                              txt_cantidad.Visible = True
                              txt_cantidad.SetFocus
                           Else
                              If var_cantidad_caja = 0 Then
                                 var_cantidad_leida = 1#
                              Else
                                 var_cantidad_leida = var_cantidad_caja
                              End If
                              txt_foco.Enabled = True
                              txt_foco.SetFocus
                           End If
                        Else
                           Call cmd_mensaje_1_Click
                           frmmensaje.lbl_mensaje = "El artículo no existe"
                           frmmensaje.Show 1
                           'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                           txt_codigo = ""
                        End If
                     Else
                        Call cmd_mensaje_1_Click
                        frmmensaje.lbl_mensaje = "El artículo no existe"
                        frmmensaje.Show 1
                       'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                        txt_codigo = ""
                        rs.Close
                     End If
                  End If
               End If
            End If
         Else
            txt_codigo = ""
            Call cmd_mensaje_1_Click
            frmmensaje.lbl_mensaje = "Error en Código"
            frmmensaje.Show 1
            'MsgBox "Error en Código", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub




Private Sub txt_foco_GotFocus()
   Dim pError As ADODB.Error
   Dim var_actualiza As Boolean
   Dim var_inserta As Boolean
   Dim bandera_suma As Boolean
   Dim var_cantidad As Variant
   Dim var_costo As Variant
   Dim var_precio As Variant
   Dim var_posible_caja As Boolean
   Dim var_cantidad_posible As Variant
   Dim var_embarque_paquete As Integer
   Dim var_embarque_caja As Integer
   Dim var_estatus_caja As String
   Dim var_orden_surtido_caja As Integer
   Dim var_posible_empaque As Boolean
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   Dim var_encontrado As Integer
   Dim var_canal_venta As String
   Dim var_i As Integer
   Dim var_n As Integer
   Dim var_j As Integer
   Dim var_tipo_pedido As String
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_DET_ORDEN_SURTIDO_M = New TB_DET_ORDEN_SURTIDO_M
   Set TB_DET_EMBARQUE_I = New TB_DET_EMBARQUE_I
   Set TB_DETALLE_CAJAS_M = New TB_DETALLE_CAJAS_M
   'On Error GoTo salir:
   z = 0
   cnn.CommandTimeout = 360
   If Trim(txt_codigo.Text) <> "" Then
      var_posible_empaque = False 'sirve para no meter articulos a granel con cajas
      If var_es_caja = True And var_cajas = True Then
         var_posible_empaque = True
      End If
      If var_es_caja = False And var_cajas = False Then
         var_posible_empaque = True
      End If
      If var_posible_empaque = True Then
         rsaux.Open "SELECT * FROM TB_EMBARQUES_ORDENES_SURTIDO WHERE VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND INTE_ORS_ORDEN_SURTIDO = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
         If rsaux.EOF Then
            rsaux5.Open "INSERT INTO TB_EMBARQUES_ORDENES_SURTIDO (VCHA_EMP_EMPRESA_ID, INTE_EMB_EMBARQUE, INTE_ORS_ORDEN_SURTIDO) VALUES ('" + var_empresa + "', " + Me.txt_embarque + "," + Me.txt_archivo + ")", cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux.Close
         var_posible_caja = False
         bandera_suma = False
         If var_primera_vez = True Then
            var_inserta = False
            rsaux.Open "select vcha_can_canal_venta_id from tb_agentes where vcha_age_agente_id = '" + var_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
            var_canal_venta = IIf(IsNull(rsaux!vcha_can_canal_venta_id), "", rsaux!vcha_can_canal_venta_id)
            rsaux.Close
            var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, var_orden_surtido, var_clave_cliente, "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", txt_archivo, var_clave_establecimiento, "B", var_clave_titular, var_clave_agente, var_descuento_1, var_descuento_2, var_descuento_3, var_clave_moneda, 0)
            var_numero_folio = var_numero_folio_regreso
            If var_factura_ceros = 1 Then
               rsaux.Open "update tb_encabezado_movimientos set inte_emo_factura_ceros = 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            var_pedido_credito = 1
            If var_clave_movimiento = "FT" Then
               rsaux.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + CStr(var_orden_surtido), cnn, adOpenDynamic, adLockOptimistic
               var_pedido_credito = 1
               If Not rsaux.EOF Then
                  var_pedido_credito = IIf(IsNull(rsaux!inte_ors_pedido_credito), 1, rsaux!inte_ors_pedido_credito)
               End If
               rsaux.Close
            End If
            rsaux.Open "update tb_encabezado_movimientos set VCHA_CAN_CANAL_VENTA_ID = '" + var_canal_venta + "', inte_emo_pedido_credito = " + CStr(var_pedido_credito) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            
            var_inserta = False
            var_inserta = TB_DET_EMBARQUE_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_numero_embarque, var_clave_movimiento, var_numero_folio, "")
            txt_folio = var_numero_folio
            var_primera_vez = False
            var_fecha_numero = CDbl(CStr(Year(Date)) + CStr(Month(Date)) + CStr(Day(Date)))
            
            '
            'var_nombre_tabla = "TEMP_" + Trim(var_empresa) + Trim(var_unidad_organizacional) + Trim(var_almacen_origen) + Trim(var_clave_movimiento) + Trim(CStr(var_numero_folio))
            'Cadena = "CREATE TABLE [dbo].[" + var_nombre_tabla + "] ([VCHA_EMP_EMPRESA_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_UOR_UNIDAD_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_ALM_ALMACEN_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_MOV_MOVIMIENTO_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[INTE_SAL_NUMERO] [int] NULL ,[VCHA_ART_ARTICULO_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[FLOA_SAL_CANTIDAD] [float] NULL ,[FLOA_SAL_COSTO] [float] NULL ,[FLOA_SAL_PRECIO] [float] NULL ,[FLOA_SAL_DESCUENTO] [float] NULL ,[FLOA_SAL_PROMOCION_1] [float] NULL ,[FLOA_SAL_PROMOCION_2] [float] NULL ,[VCHA_REE_FOLIO] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_SAL_REFERENCIA] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[CHAR_PED_TIPO] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,[VCHA_CAT_CATALOGO_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[FLOA_SAL_DESCUENTO_1] [float] NULL ,"
            'Cadena = Cadena + " [FLOA_SAL_DESCUENTO_2] [float] NULL ,[INTE_SAL_AÑO] [int] NULL , [INTE_SAL_CONSECUTIVO] [int] NULL) ON [PRIMARY]"
            'rsaux4.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            '
            If var_tipo_lectura = 1 Then
               var_i = 1
               For var_i = 1 To lv_salidas.ListItems.Count
                   lv_salidas.ListItems.item(var_i).Selected = True
                   If var_tipo_lectura = 1 Then
                      
                      var_precio = CDbl(lv_salidas.selectedItem.SubItems(8)) * 1
                      If var_factura_ceros = 1 Then
                         var_precio = 0
                      End If
                      'rsaux5.Open "select * from tb_Salidas where vcha_sal_archivo = '" + Trim(var_empresa) + Trim(var_unidad_organizacional) + Trim(var_almacen_origen) + Trim(var_clave_movimiento) + Trim(CStr(var_numero_folio)) + "' and vcha_art_articulo_id = '" + lv_salidas.selectedItem + "'", cnnaccess, adOpenDynamic, adLockOptimistic
                      
                      'If rsaux5.EOF Then
                         Call cmd_mensaje_4_Click
                         Cadena = "insert into tb_salidas (VCHA_SAL_ARCHIVO, INTE_PED_NUMERO, INTE_ORS_ORDEN_SURTIDO, VCHA_EMP_EMPRESA_ID, INTE_SAL_NUMERO,VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ORS_CANTIDAD_SURTIR, FLOA_ORS_CANTIDAD_SURTIDA, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, "
                         Cadena = Cadena + " VCHA_SAL_TIPO, INTE_SAL_CONSECUTIVO) VALUES ('" + Trim(var_empresa) + Trim(var_unidad_organizacional) + Trim(var_almacen_origen) + Trim(var_clave_movimiento) + Trim(CStr(var_numero_folio)) + "'," + Trim(txt_pedido) + "," + txt_archivo + ",'" + var_empresa + "'," + Trim(CStr(var_numero_folio)) + ",'" + lv_salidas.selectedItem + "',''," + CStr(CDbl(lv_salidas.selectedItem.SubItems(2)) * 1) + ", " + CStr(CDbl(lv_salidas.selectedItem.SubItems(3)) * 1) + ",0," + CStr(CDbl(lv_salidas.selectedItem.SubItems(7)) * 1) + "," + CStr(var_precio) + "," + CStr(CDbl(lv_salidas.selectedItem.SubItems(9)) * 1) + "," + CStr(CDbl(lv_salidas.selectedItem.SubItems(10)) * 1) + ",'" + lv_salidas.selectedItem.SubItems(11) + "',0)"
                         Call cmd_mensaje_4_Click
                         rsaux4.Open Cadena, cnnaccess, adOpenDynamic, adLockOptimistic
                      'End If
                      'rsaux5.Close
                   End If
               Next var_i
            End If
         End If
         If var_tipo_lectura = 0 Then
            If var_es_caja = False Then
               Cadena = "select * from tb_det_orden_surtido where inte_ors_orden_surtido = " + txt_archivo + " and vcha_art_articulo_id = '" + txt_codigo + "'"
               rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_promocion_1 = IIf(IsNull(rs!floa_ors_promocion_1), 0, rs!floa_ors_promocion_1)
                  var_promocion_2 = IIf(IsNull(rs!floa_ors_promocion_2), 0, rs!floa_ors_promocion_2)
                  valor = txt_codigo
                  var_n = lv_salidas.ListItems.Count
                  var_encontro = 0
                  var_i = 1
                  While (var_i <= var_n)
                        var_tipo_pedido = lv_salidas.selectedItem.SubItems(11)
                        lv_salidas.ListItems.item(var_i).Selected = True
                        valor = Trim(lv_salidas.selectedItem)
                        If txt_codigo = valor Then
                           var_cantidad_posible = lv_salidas.selectedItem.SubItems(2)
                           If var_cantidad_posible < lv_salidas.selectedItem.SubItems(3) + var_cantidad_leida Then
                              var_encontro = 0
                           Else
                              var_encontro = 1
                              var_i = var_n + 1
                           End If
                        End If
                        var_i = var_i + 1
                  Wend
                  If var_encontro = 1 Then
                     var_cantidad_posible = lv_salidas.selectedItem.SubItems(2)
                     If var_cantidad_posible < lv_salidas.selectedItem.SubItems(3) + var_cantidad_leida Then
                        Call cmd_mensaje_2_Click
                        frmmensaje.lbl_mensaje = "Cantidad supera a la posible a surtir"
                        frmmensaje.Show 1
                     Else
                        var_tipo_pedido = lv_salidas.selectedItem.SubItems(11)
                        lv_salidas.selectedItem.SubItems(6) = Format(lv_salidas.selectedItem.SubItems(2) - (var_cantidad_leida + lv_salidas.selectedItem.SubItems(3) + lv_salidas.selectedItem.SubItems(4)), "###,###,##0.00")
                        lv_salidas.selectedItem.SubItems(4) = lv_salidas.selectedItem.SubItems(4)
                        lv_salidas.selectedItem.SubItems(3) = Format(lv_salidas.selectedItem.SubItems(3) + var_cantidad_leida, "###,###,##0.00")
                        lv_salidas.selectedItem.SubItems(5) = Format(lv_salidas.selectedItem.SubItems(5) + var_cantidad_leida, "###,###,##0.00")
                        var_renglon = lv_salidas.selectedItem.Index
                        Call ilumina_grid
                        var_costo = lv_salidas.selectedItem.SubItems(7)
                        var_precio = lv_salidas.selectedItem.SubItems(8)
                        var_cantidad = lv_salidas.selectedItem.SubItems(4)
                        lbl_recibidos = Format(CDbl(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                        var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                        If rsaux5.State = 1 Then
                           rsaux5.Close
                        End If
                        '26-09-08 para afectar orden surtido hasta que se cierra el embarque
                        rsaux5.Open "update tb_det_orden_surtido set floa_ors_cantidad_surtida = floa_ors_cantidad_surtida + " + CStr(var_cantidad_leida) + " where inte_ors_orden_surtido = " + CStr(var_orden_surtido) + " and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        If rsaux5.State = 1 Then
                           rsaux5.Close
                        End If
                        bandera_suma = True
                     End If
                  Else
                     Call cmd_mensaje_2_Click
                     txt_codigo = ""
                     frmmensaje.lbl_articulo = var_nombre_articulo_mensaje
                     frmmensaje.lbl_mensaje = "Cantidad supera a la posible a surtir"
                     frmmensaje.Show 1
                  End If
               Else
                  Call cmd_mensaje_2_Click
                  txt_codigo = ""
                  frmmensaje.lbl_articulo = var_nombre_articulo_mensaje
                  frmmensaje.lbl_mensaje = "El artículo no se encuentra dentro de la Orden de Surtido"
                  frmmensaje.Show 1
               End If
               rs.Close
               If bandera_suma = True Then
                  If var_factura_ceros = 1 Then
                     var_precio = 0
                  End If
                  Cadena = "select * from " + var_nombre_tabla + " where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "' and floa_sal_precio = " + CStr(var_precio) + " and char_ped_tipo = '" + var_tipo_pedido + "'"
                  rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_inserta = False
                     rsaux.Open "update " + var_nombre_tabla + " set floa_sal_cantidad = floa_sal_cantidad +" + CStr(var_cantidad_leida) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_Sal_Numero = " + CStr(var_numero_folio) + " and vcha_art_articulo_id = '" + lv_salidas.selectedItem + "' and round(floa_sal_precio,2) = round(" + CStr(var_precio) + ",2) and char_ped_tipo = '" + var_tipo_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
                     rs.Close
                  Else
                     var_inserta = False
                     var_consecutivo = var_consecutivo + 1
                     rsaux.Open "INSERT INTO " + var_nombre_tabla + " (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, CHAR_PED_TIPO, INTE_SAL_CONSECUTIVO) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ",  " + CStr(var_precio) + ", 0,  " + CStr(var_promocion_1) + ", " + CStr(var_promocion_2) + ",'" + var_tipo_pedido + "', " + CStr(var_consecutivo) + ") ", cnn, adOpenDynamic, adLockOptimistic
                     rs.Close
                  End If
                  bandera_suma = False
               End If
            Else
            End If
         Else
''''metodo nuevo
            'cnnaccess.BeginTrans
            If var_es_caja = False Then
               var_archivo_tabla = Trim(var_empresa) + Trim(var_unidad_organizacional) + Trim(var_almacen_origen) + Trim(var_clave_movimiento) + Trim(CStr(var_numero_folio))
               Cadena = "select * from tb_salidas where vcha_sal_archivo = '" + var_archivo_tabla + "' and inte_ors_orden_surtido = " + txt_archivo + " and vcha_art_articulo_id = '" + txt_codigo + "'"
               If rs.State = 1 Then
                  rs.Close
               End If
               Call cmd_mensaje_4_Click
               rs.Open Cadena, cnnaccess, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_promocion_1 = IIf(IsNull(rs!floa_sal_promocion_1), 0, rs!floa_sal_promocion_1)
                  var_promocion_2 = IIf(IsNull(rs!FLOA_SAL_PROMOCION_2), 0, rs!FLOA_SAL_PROMOCION_2)
                  valor = txt_codigo
                  var_n = lv_salidas.ListItems.Count
                  var_encontro = 0
                  var_i = 1
                  While (var_i <= var_n)
                        var_tipo_pedido = lv_salidas.selectedItem.SubItems(11)
                        lv_salidas.ListItems.item(var_i).Selected = True
                        valor = Trim(lv_salidas.selectedItem)
                        If UCase(txt_codigo) = UCase(valor) Then
                           var_cantidad_posible = lv_salidas.selectedItem.SubItems(2)
                           If var_cantidad_posible < lv_salidas.selectedItem.SubItems(3) + var_cantidad_leida Then
                              var_encontro = 0
                           Else
                              var_encontro = 1
                              var_i = var_n + 1
                           End If
                        End If
                        var_i = var_i + 1
                  Wend
                  If var_encontro = 1 Then
                     var_cantidad_posible = lv_salidas.selectedItem.SubItems(2)
                     If var_cantidad_posible < lv_salidas.selectedItem.SubItems(3) + var_cantidad_leida Then
                        Call cmd_mensaje_2_Click
                        frmmensaje.lbl_mensaje = "Cantidad supera a la posible a surtir"
                        frmmensaje.Show 1
                     Else
                        var_tipo_pedido = lv_salidas.selectedItem.SubItems(11)
                        lv_salidas.selectedItem.SubItems(6) = Format(lv_salidas.selectedItem.SubItems(2) - (var_cantidad_leida + lv_salidas.selectedItem.SubItems(3) + lv_salidas.selectedItem.SubItems(4)), "###,###,##0.00")
                        lv_salidas.selectedItem.SubItems(4) = lv_salidas.selectedItem.SubItems(4)
                        lv_salidas.selectedItem.SubItems(3) = Format(lv_salidas.selectedItem.SubItems(3) + var_cantidad_leida, "###,###,##0.00")
                        lv_salidas.selectedItem.SubItems(5) = Format(lv_salidas.selectedItem.SubItems(5) + var_cantidad_leida, "###,###,##0.00")
                        var_renglon = lv_salidas.selectedItem.Index
                        Call ilumina_grid
                        var_costo = lv_salidas.selectedItem.SubItems(7)
                        var_precio = lv_salidas.selectedItem.SubItems(8)
                        var_cantidad = lv_salidas.selectedItem.SubItems(4)
                        lbl_recibidos = Format(CDbl(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                        var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                        If rsaux5.State = 1 Then
                           rsaux5.Close
                        End If
                        rsaux5.Open "update TB_DETALLE_EQUIPOS_ORDEN_SURTIDO set FLOA_ORS_CANTIDAD_SURTIDA = isnull(FLOA_ORS_CANTIDAD_SURTIDA,0) + " + CStr(var_cantidad_leida) + " where INTE_ORS_ORDEN_SURTIDO = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
                        Call cmd_mensaje_4_Click
                        rsaux5.Open "update tb_Salidas set floa_ors_cantidad_surtida = floa_ors_cantidad_surtida + " + CStr(var_cantidad_leida) + " where inte_ors_orden_surtido = " + CStr(var_orden_surtido) + " and vcha_art_articulo_id = '" + txt_codigo + "' and vcha_sal_Archivo = '" + var_archivo_tabla + "'", cnnaccess, adOpenDynamic, adLockOptimistic
                        If rsaux5.State = 1 Then
                           rsaux5.Close
                        End If
                        bandera_suma = True
                     End If
                  Else
                     Call cmd_mensaje_2_Click
                     txt_codigo = ""
                     frmmensaje.lbl_articulo = var_nombre_articulo_mensaje
                     frmmensaje.lbl_mensaje = "Cantidad supera a la posible a surtir"
                     frmmensaje.Show 1
                  End If
               Else
                  Call cmd_mensaje_2_Click
                  txt_codigo = ""
                  frmmensaje.lbl_articulo = var_nombre_articulo_mensaje
                  frmmensaje.lbl_mensaje = "El artículo no se encuentra dentro de la Orden de Surtido"
                  frmmensaje.Show 1
               End If
               rs.Close
               If bandera_suma = True Then
                  If var_factura_ceros = 1 Then
                     var_precio = 0
                  End If
                  var_inserta = False
                  If rsaux4.State = 1 Then
                     rsaux4.Close
                  End If
                  If var_posible_kanban = 1 Then
                     Set TB_RESERVAR_FUERA_DE_KANBAN = New TB_RESERVAR_FUERA_DE_KANBAN
                     Set TB_RESERVAR_KANBAN = New TB_RESERVAR_KANBAN
                     If var_kanban_es_un_kanban = "S" Then
                        var_inserta = TB_RESERVAR_KANBAN.Anadir(var_kanban, var_clave_movimiento, var_numero_folio, var_almacen_origen, Me.txt_codigo, "", "")
                        If var_kanban_exito = "S" Then
                           var_posible_leido = 1
                        Else
                           var_posible_leido = 0
                        End If
                     Else
                        var_inserta = TB_RESERVAR_FUERA_DE_KANBAN.Anadir(var_numero_folio, var_clave_movimiento, var_almacen_origen, Me.txt_codigo, "", "")
                        If var_kanban_exito = "S" Then
                           var_posible_leido = 1
                        Else
                           var_posible_leido = 0
                        End If
                     End If
                     
                  Else
                     var_kanban_mensaje = ""
                     var_posible_leido = 1
                  End If
                  If var_posible_leido = 1 Then
                     '26-09-08 para afectar orden surtido hasta que se cierra el embarque
                     'MsgBox cnn.ConnectionString
                     rsaux5.Open "update tb_det_orden_surtido set floa_ors_cantidad_surtida = floa_ors_cantidad_surtida + " + CStr(var_cantidad_leida) + " where inte_ors_orden_surtido = " + CStr(var_orden_surtido) + " and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     Call cmd_mensaje_4_Click
                     rsaux4.Open "SELECT * FROM TB_SALIDAS where vcha_sal_archivo = '" + var_archivo_tabla + "' and vcha_art_articulo_id = '" + lv_salidas.selectedItem + "' AND INTE_SAL_CONSECUTIVO = 0", cnnaccess, adOpenDynamic, adLockOptimistic
                     If Not rsaux4.EOF Then
                        var_consecutivo = var_consecutivo + 1
                        Call cmd_mensaje_4_Click
                        rsaux.Open "update tb_salidas set floa_sal_cantidad = floa_sal_cantidad +" + CStr(var_cantidad_leida) + ", INTE_SAL_CONSECUTIVO = " + CStr(var_consecutivo) + " where vcha_sal_archivo = '" + var_archivo_tabla + "' and vcha_art_articulo_id = '" + lv_salidas.selectedItem + "'", cnnaccess, adOpenDynamic, adLockOptimistic
                     Else
                        Call cmd_mensaje_4_Click
                        rsaux.Open "update tb_salidas set floa_sal_cantidad = floa_sal_cantidad +" + CStr(var_cantidad_leida) + " where vcha_sal_archivo = '" + var_archivo_tabla + "' and vcha_art_articulo_id = '" + lv_salidas.selectedItem + "'", cnnaccess, adOpenDynamic, adLockOptimistic
                     End If
                     bandera_suma = False
                  Else
                     lv_salidas.selectedItem.SubItems(6) = Format(lv_salidas.selectedItem.SubItems(2) - (CDbl(lv_salidas.selectedItem.SubItems(3)) + CDbl(lv_salidas.selectedItem.SubItems(4)) - var_cantidad_leida), "###,###,##0.00")
                     lv_salidas.selectedItem.SubItems(4) = lv_salidas.selectedItem.SubItems(4)
                     lv_salidas.selectedItem.SubItems(3) = Format(lv_salidas.selectedItem.SubItems(3) - var_cantidad_leida, "###,###,##0.00")
                     lv_salidas.selectedItem.SubItems(5) = Format(lv_salidas.selectedItem.SubItems(5) - var_cantidad_leida, "###,###,##0.00")
                     var_renglon = lv_salidas.selectedItem.Index
                     lbl_recibidos = Format(CDbl(lbl_recibidos) - var_cantidad_leida, "###,###,##0.00")
                     var_cantidad_recibida = var_cantidad_recibida - var_cantidad_leida
                     
                     
                     txt_codigo = ""
                     frmmensaje.lbl_mensaje = var_kanban_mensaje
                     frmmensaje.Show 1
                  End If
               End If
            Else
            End If
            'cnnaccess.CommitTrans
''''' fin metodo nuevo
         End If
      Else
         txt_codigo = ""
         frmmensaje.lbl_mensaje = "No es posible mezclar mercancia a granel con mercancia empacada"
         frmmensaje.Show 1
      End If
      txt_codigo.SetFocus
   End If
'   Exit Sub
'salir:
'Resume
End Sub


Sub ejecuta()
   Dim var_embarque_agente As String
   Dim var_embarque_almacen As String
   Dim var_movimiento_agente As String
   Dim var_embarque_cerrado As String
   Dim var_almacen_empaque_nombre As String
   Dim var_almacen_empaque As String
   Dim var_posible_lectura As Boolean
   var_autorizo_embarque = False
   Dim var_liberada As Integer
   If Dir(App.Path & "\bd_salidas.mdb") <> "" Then
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open "select * from tb_encabezado_embarques where inte_emb_embarque = " + Str(var_numero_embarque) + " and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   var_embarque_agente = rs!VCHA_AGE_AGENTE_ID
   var_embarque_cerrado = Trim(rs!CHAR_EMB_ESTATUS)
   If Not rs.EOF Then
      var_embarque_agente = rs!VCHA_AGE_AGENTE_ID
      rs.Close
      var_clave_movimiento = txt_clave_movimiento
      rs.Open "select * from tb_detalle_cajas with (nolock) where inte_ors_orden_surtido = " + txt_archivo + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and char_paq_estatus <> 'S'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         si = MsgBox("La orden de surtido a sido empaquetada, ¿Desea subir las cajas?", vbYesNo, "ATENCION")
         If si = 6 Then
            var_cajas = True
         Else
            var_cajas = False
         End If
      Else
         var_cajas = False
      End If
      rs.Close
      rs.Open "select * from vw_orden_surtido where inte_ors_orden_surtido = " + txt_archivo + " and floa_ors_cantidad_surtir > 0", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_liberada = IIf(IsNull(rs!inte_ors_liberada), 0, rs!inte_ors_liberada)
         If var_liberada = 1 Then
            If var_clave_movimiento = rs!VCHA_MOV_MOVIMIENTO_ID Then
               var_posible_lectura = True
               If var_tipo_lectura = 1 Then
                  If rsaux4.State = 1 Then
                     rsaux4.Close
                  End If
                  rsaux4.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "' and inte_emo_numero_origen = " + txt_archivo + " and char_emo_estatus = ''", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux4.EOF Then
                     var_posible_lectura = False
                  End If
                  rsaux4.Close
               End If
               If var_posible_lectura = True Then
                  var_clave_moneda = rs!vcha_mon_moneda_id
                  var_movimiento_agente = rs!VCHA_AGE_AGENTE_ID
                  If var_movimiento_agente = var_embarque_agente Then
                     var_autorizo_embarque = True
                  Else
                     si = MsgBox("La orden de surtido no corresponde al agente con el que se inicio el embarque, ¿desea agregarlo?", vbYesNo, "ATENCION")
                     If si = 6 Then
                        var_autorizo_embarque = False
                        var_autoriza_mov = False
                        var_opcion_seguridad = 3
                        frmpasswords.Show 1
                        var_autorizo_embarque = var_autoriza_mov
                     Else
                        var_autorizo_embarque = False
                     End If
                  End If
                  If var_clave_movimiento = rs!VCHA_MOV_MOVIMIENTO_ID Then
                     var_orden_surtido = txt_archivo
                     var_suma_cantidad_enviada = 0
                     var_suma_cantidad_recibida = 0
                     lbl_enviados.Caption = "0"
                     lbl_recibidos.Caption = "0"
                     lv_salidas.ListItems.Clear
               
                     If IsNull(rs!vcha_cli_email) Then
                        var_correo_electronico = ""
                     Else
                        var_correo_electronico = rs!vcha_cli_email
                     End If
                     If IsNull(rs!VCHA_ALM_NOMBRE) Then
                        GoTo no_almacen:
                     Else
                       var_almacen_OS = rs!VCHA_ALM_ALMACEN_ID
                       var_almacen_origen = rs!VCHA_ALM_ALMACEN_ID
                       txt_origen = rs!VCHA_ALM_NOMBRE
                     End If
                     If IsNull(rs!VCHA_TIT_NOMBRE) Then
                        GoTo no_titular:
                     Else
                        txt_titular = rs!VCHA_TIT_NOMBRE
                        var_clave_titular = rs!vcha_tit_titular_id
                     End If
                     If IsNull(rs!inte_ped_dias_condiciones) Then
                        var_plazo = 0
                     Else
                        var_plazo = rs!inte_ped_dias_condiciones
                     End If
                     If IsNull(rs!VCHA_ESB_NOMBRE) Then
                        GoTo no_establecimiento:
                     Else
                        txt_establecimiento = rs!VCHA_ESB_NOMBRE
                        var_clave_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
                     End If
                     If IsNull(rs!VCHA_AGE_NOMBRE) Then
                        GoTo no_agente:
                     Else
                        txt_agente = rs!VCHA_AGE_NOMBRE
                        var_clave_agente = rs!VCHA_AGE_AGENTE_ID
                     End If
                     var_almacen_Destino = ""
                     If var_tipo_documento = "V" Then
                        If IsNull(rs!almacen_agente) Then
                           GoTo no_almacen_agente:
                        Else
                           var_almacen_Destino = rs!almacen_agente
                        End If
                     End If
                     If IsNull(rs!VCHA_CLI_NOMBRE) Then
                        GoTo no_cliente:
                     Else
                        txt_cliente = rs!VCHA_CLI_NOMBRE
                        var_clave_cliente = rs!vcha_cli_clave_id
                     End If
                     If IsNull(rs!vcha_rut_nombre) Then
                        txt_ruta = ""
                        var_clave_ruta = ""
                     Else
                        txt_ruta = rs!vcha_rut_nombre
                        var_clave_ruta = rs!vcha_rut_ruta_id
                     End If
                     If IsNull(rs!inte_ped_numero) Then
                        GoTo no_Pedido:
                     Else
                        txt_pedido = rs!inte_ped_numero
                     End If
                     If IsNull(rs!FLOA_ORS_DESCUENTO_1) Then
                        txt_descuento1 = 0
                        var_descuento_1 = 0
                     Else
                        txt_descuento1 = rs!FLOA_ORS_DESCUENTO_1
                        var_descuento_1 = rs!FLOA_ORS_DESCUENTO_1
                     End If
                     If IsNull(rs!FLOA_ORS_DESCUENTO_2) Then
                        txt_descuento2 = 0
                        var_descuento_2 = 0
                     Else
                        txt_descuento2 = rs!FLOA_ORS_DESCUENTO_2
                        var_descuento_2 = rs!FLOA_ORS_DESCUENTO_2
                     End If
                     var_descuento_3 = 0
                     While Not rs.EOF
                        var_factura_ceros = IIf(IsNull(rs!inte_ors_factura_ceros), 0, rs!inte_ors_factura_ceros)
                        Set list_item = lv_salidas.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
                        list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_nombre_español), "", Trim(rs!vcha_Art_nombre_español))
                        list_item.SubItems(2) = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0.00")
                        var_surtir = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR)
                        list_item.SubItems(3) = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA) + IIf(IsNull(rs!floa_ors_cantidad_negada), 0, rs!floa_ors_cantidad_negada), "###,###,##0.00")
                        list_item.SubItems(4) = 0
                        var_empacada = 0
                        list_item.SubItems(5) = Format(0, "###,###,##0.00")
                        var_falta = 0
                        list_item.SubItems(6) = (list_item.SubItems(2) * 1) - (list_item.SubItems(3) * 1)
                        list_item.SubItems(7) = IIf(IsNull(rs!floa_ors_costo), 0, rs!floa_ors_costo)
                        list_item.SubItems(8) = IIf(IsNull(rs!floa_ors_precio), 0, rs!floa_ors_precio)
                        list_item.SubItems(9) = IIf(IsNull(rs!floa_ors_promocion_1), 0, rs!floa_ors_promocion_1)
                        list_item.SubItems(10) = IIf(IsNull(rs!floa_ors_promocion_2), 0, rs!floa_ors_promocion_2)
                        list_item.SubItems(11) = IIf(IsNull(rs!char_ped_tipo), "P", rs!char_ped_tipo)
                        var_suma_cantidad_enviada = var_suma_cantidad_enviada + rs!FLOA_ORS_CANTIDAD_SURTIR
                        var_suma_cantidad_recibida = var_suma_cantidad_recibida + rs!FLOA_ORS_CANTIDAD_SURTIDA
                        rs.MoveNext:
                     Wend
                     lbl_enviados = Format(Str(var_suma_cantidad_enviada), "###,###,##0.00")
                     lbl_recibidos = Format(Str(var_suma_cantidad_recibida), "###,###,##0.00")
                     txt_codigo.Enabled = True
                     txt_archivo.Enabled = False
                     lbl_enviados = Format(var_suma_cantidad_enviada, "###,###,##0.00")
                     lbl_recibidos = Format(var_suma_cantidad_recibida, "###,###,##0.00")
                     If var_autorizo_embarque = True Then
                        txt_codigo.Enabled = True
                        txt_archivo.Enabled = False
                     Else
                        txt_codigo.Enabled = False
                        txt_archivo.Enabled = False
                     End If
                     If var_embarque_cerrado = "I" Then
                        MsgBox "El embarque ya fue cerrado", vbOKOnly, "ATENCION"
                        txt_codigo.Enabled = False
                        txt_archivo.Enabled = False
                    End If
                  Else
                     MsgBox "Orden de surtido incorrecta para este movimiento", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "La orden de surtido se encuentra abierta en otro movimiento", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Nota de envio incorrecta para este movimiento", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "La orden de surtido no a sido liberada", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Número de Orden de surtido no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      rs.Close
   End If
   var_renglon = -1
   Call ilumina_grid
   var_n = lv_salidas.ListItems.Count
   var_numero_renglones = Me.lv_salidas.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_salidas.ColumnHeaders(6).Width = 930
   Else
      lv_salidas.ColumnHeaders(6).Width = 1130
   End If
   If Me.txt_codigo.Enabled = True Then
      Me.txt_codigo.SetFocus
   End If
   Else
      MsgBox "La maquina " + fun_NombrePc + " no cuenta con el archivo bd_salidas.mdb, favor de copiarlo de el servido", vbOKOnly, "ATENCION"
   End If
   Exit Sub
no_almacen:
   MsgBox "Almacen Incorrecto", vbOKOnly, "ATENCION"
   Exit Sub
no_Pedido:
   MsgBox "Pedido Incorrecto", vbOKOnly, "ATENCION"
   Exit Sub
no_establecimiento:
   MsgBox "Establecimiento Incorrecto", vbOKOnly, "ATENCION"
   Exit Sub
no_agente:
   MsgBox "Agente Incorrecto", vbOKOnly, "ATENCION"
   Exit Sub
no_cliente:
   MsgBox "Cliente Incorrecto", vbOKOnly, "ATENCION"
   Exit Sub
no_ruta:
   MsgBox "Ruta Incorrecta", vbOKOnly, "ATENCION"
   Exit Sub
no_titular:
   MsgBox "Titular incorrecto", vbOKOnly, "ATENCION"
   Exit Sub
no_almacen_agente:
   MsgBox "No existe un almacen relacionado a este agente", vbOKOnly, "ATENCION"
   Exit Sub
End Sub

Private Sub txt_sello_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmd_aceptar_sello.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_sellos.Visible = False
   End If
End Sub

