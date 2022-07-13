VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_salidas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Embarques"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_mensaje_4 
      Caption         =   "mensaje 4"
      Height          =   195
      Left            =   1965
      TabIndex        =   54
      Top             =   675
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton cmd_mensaje_2 
      Caption         =   "mensaje 2"
      Height          =   195
      Left            =   1800
      TabIndex        =   52
      Top             =   675
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11175
      Picture         =   "frmoracle_salidas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Salir"
      Top             =   630
      Width           =   330
   End
   Begin VB.Frame frm_sellos 
      Height          =   2340
      Left            =   795
      TabIndex        =   3
      Top             =   855
      Width           =   3045
      Begin VB.CommandButton cmd_cancelar_sello 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmoracle_salidas.frx":063A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton cmd_aceptar_sello 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   45
         Picture         =   "frmoracle_salidas.frx":0784
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   360
         Width           =   330
      End
      Begin VB.TextBox txt_sello 
         Height          =   315
         Left            =   585
         TabIndex        =   6
         Top             =   795
         Width           =   2385
      End
      Begin VB.CommandButton cmd_cerrar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   705
         Picture         =   "frmoracle_salidas.frx":08CE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cerrar Alt + C"
         Top             =   360
         Width           =   330
      End
      Begin VB.Frame Frame4 
         Height          =   75
         Left            =   30
         TabIndex        =   4
         Top             =   645
         Width           =   2970
      End
      Begin MSComctlLib.ListView lv_sellos 
         Height          =   1200
         Left            =   30
         TabIndex        =   9
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
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   "Sellos"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   7
         Left            =   30
         TabIndex        =   11
         Top             =   120
         Width           =   2970
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sello:"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   840
         Width           =   390
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmoracle_salidas.frx":09D0
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   630
      Width           =   330
   End
   Begin VB.CommandButton cmd_cerrar_embarque 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmoracle_salidas.frx":0AD2
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cerrar Embarque"
      Top             =   630
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   4320
      Left            =   105
      TabIndex        =   38
      Top             =   2985
      Width           =   11475
      Begin VB.TextBox txt_foco 
         Height          =   315
         Left            =   11655
         TabIndex        =   51
         Top             =   525
         Width           =   1650
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
         TabIndex        =   2
         Top             =   495
         Width           =   1890
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   4440
         TabIndex        =   40
         Top             =   1575
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   75
            TabIndex        =   41
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
            TabIndex        =   42
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
         TabIndex        =   1
         Top             =   465
         Width           =   3390
      End
      Begin VB.CommandButton cmd_pasar_movimiento 
         Height          =   330
         Left            =   8880
         Picture         =   "frmoracle_salidas.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   540
         Visible         =   0   'False
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_salidas 
         Height          =   3225
         Left            =   15
         TabIndex        =   43
         Top             =   1035
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   5689
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "   Código"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   10583
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
            Text            =   "Faltan    "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "inventory item id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "delivery detail id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "source line number"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "delivery_id"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Lectura de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   0
         Left            =   30
         TabIndex        =   46
         Top             =   120
         Width           =   11400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   615
         Width           =   1395
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   5115
         TabIndex        =   44
         Top             =   615
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   45
      TabIndex        =   16
      Top             =   870
      Width           =   11505
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   60
      TabIndex        =   15
      Top             =   510
      Width           =   11505
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   75
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
            Picture         =   "frmoracle_salidas.frx":0CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas.frx":15B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas.frx":1E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas.frx":2426
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas.frx":2D02
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas.frx":35DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas.frx":3EB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas.frx":3FC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas.frx":40DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas.frx":41EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas.frx":42FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas.frx":4410
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas.frx":4522
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas.frx":46C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas.frx":5516
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas.frx":56EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas.frx":57FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   915
      Width           =   6975
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
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   375
         Width           =   1620
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
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   375
         Width           =   1140
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   1
         Left            =   30
         TabIndex        =   22
         Top             =   120
         Width           =   6900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   840
         TabIndex        =   21
         Top             =   525
         Width           =   765
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Jaula:"
         Height          =   195
         Left            =   3810
         TabIndex        =   20
         Top             =   525
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Index           =   3
      Left            =   7140
      TabIndex        =   23
      Top             =   915
      Width           =   2220
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
         TabIndex        =   25
         Top             =   420
         Width           =   1845
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Cantidad a Surtir"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   4
         Left            =   30
         TabIndex        =   24
         Top             =   120
         Width           =   2145
      End
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Index           =   4
      Left            =   9450
      TabIndex        =   26
      Top             =   915
      Width           =   2115
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
         TabIndex        =   28
         Top             =   420
         Width           =   1770
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Cantidad Surtida"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   5
         Left            =   30
         TabIndex        =   27
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1185
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   1800
      Width           =   11460
      Begin VB.TextBox txt_origen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1395
         TabIndex        =   32
         Top             =   420
         Width           =   4230
      End
      Begin VB.TextBox txt_archivo 
         Height          =   315
         Left            =   7080
         TabIndex        =   0
         Top             =   750
         Width           =   1170
      End
      Begin VB.TextBox txt_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1395
         TabIndex        =   31
         Top             =   750
         Width           =   4230
      End
      Begin VB.TextBox txt_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7080
         TabIndex        =   30
         Top             =   420
         Width           =   4230
      End
      Begin VB.Label lbl_archivo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "O. de Surtido:"
         Height          =   195
         Left            =   6075
         TabIndex        =   37
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   36
         Top             =   420
         Width           =   660
      End
      Begin VB.Label label 
         BackColor       =   &H000000C0&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   30
         TabIndex        =   35
         Top             =   120
         Width           =   11385
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   195
         TabIndex        =   34
         Top             =   750
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   6075
         TabIndex        =   33
         Top             =   420
         Width           =   525
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   75
      Left            =   0
      TabIndex        =   55
      Top             =   0
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
   Begin WMPLibCtl.WindowsMediaPlayer wmp2 
      Height          =   135
      Left            =   1515
      TabIndex        =   53
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
      Left            =   60
      TabIndex        =   50
      Top             =   30
      Width           =   11445
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   30
      Left            =   8520
      TabIndex        =   49
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
   Begin WMPLibCtl.WindowsMediaPlayer wmp3 
      Height          =   30
      Left            =   4740
      TabIndex        =   48
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
   Begin WMPLibCtl.WindowsMediaPlayer wmp4 
      Height          =   75
      Left            =   10215
      TabIndex        =   47
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
End
Attribute VB_Name = "frmoracle_salidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_localizador_subinventario As String
Dim var_localizador As Integer
Dim var_encontro As Integer
Dim var_cantidad_leida As Double
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim var_suma As Integer
Dim var_renglon As Integer

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
          If lv_salidas.ListItems.Item(var_i).ListSubItems(4) * 1 = 0 Then
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
   var_renglon = 0
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_estatus_embarque = IIf(IsNull(rs!char_Emb_estatus), "", rs!char_Emb_estatus)
      var_tipo_pedido_embarque = IIf(IsNull(rs!tipo_pedido), "", rs!tipo_pedido)
   Else
      var_estatus_embarque = "I"
   End If
   rs.Close
   If var_estatus_embarque = "" Then
      If IsNumeric(Me.txt_archivo) Then
         var_orden = CDbl(Me.txt_archivo)
         rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_cadena = "SELECT HCAS.CUST_ACCOUNT_ID, a.source_header_type_name, oha.source_document_id, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, oha.attribute8, oha.attribute9 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND  A.SOURCE_HEADER_NUMBER = '" + CStr(var_orden) + "'"
         var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID  AND released_status = 'Y' AND HCAS.ORG_ID = " + var_empresa
         rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If var_tipo_pedido_embarque = "" Then
               var_tipo_pedido_embarque = rs!source_header_type_name
               rsaux.Open "update XXVIA_TB_ENCABEZADO_EMBARQUES set tipo_pedido = '" + rs!source_header_type_name + "' WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
            End If
            If rs!source_header_type_name = var_tipo_pedido_embarque Then
               rsaux.Open "SELECT oha.header_id, oha.ordered_date, oha.order_number,  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  f.orig_system_reference from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU,  hz_cust_acct_sites_all f Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND  HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND  oha.order_type_id in (1106, 1049) and HCSU.site_use_code = 'BILL_TO' and f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID and order_number  = '" + Me.txt_archivo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
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
                     If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rs!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                        If var_pedido_tienda = 0 Then
                           rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rs!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux2.EOF Then
                              Me.txt_agente = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                           Else
                              rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              Me.txt_agente = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                              rsaux4.Close
                           End If
                           rsaux2.Close
                        Else
                           Me.txt_agente = IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9)
                        End If
                     Else
                        rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        Me.txt_agente = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                        rsaux4.Close
                     End If
                     Me.txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                     Me.txt_origen = IIf(IsNull(rs!subinventory), "", rs!subinventory)
                     Me.lv_salidas.ListItems.Clear
                     var_cantidad_enviada = 0
                     While Not rs.EOF
                           Set list_item = lv_salidas.ListItems.Add(, , rs!SEGMENT1)
                           list_item.SubItems(1) = IIf(IsNull(rs!Description), "", rs!Description)
                           list_item.SubItems(2) = Format(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity), "###,###,##0.00")
                           var_cantidad_enviada = var_cantidad_enviada + IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)
                           list_item.SubItems(3) = 0
                           list_item.SubItems(4) = Format(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity), "###,###,##0.00")
                           list_item.SubItems(5) = IIf(IsNull(rs!inventory_item_id), 0, rs!inventory_item_id)
                           list_item.SubItems(6) = IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)
                           list_item.SubItems(7) = IIf(IsNull(rs!SOURCE_LINE_NUMBER), 0, rs!SOURCE_LINE_NUMBER)
                           list_item.SubItems(8) = IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)
                           rs.MoveNext
                     Wend
                     Me.lbl_enviados = Format(var_cantidad_enviada, "###,###,##0.00")
                     Me.lbl_recibidos = Format(0, "###,###,##0.00")
                     Me.txt_archivo.Enabled = False
                     var_cantidad_recibida = 0
                     rsaux2.Open "SELECT * FROM XXVIA_TB_SALIDAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND  source_header_number = " + CStr(CDbl(Me.txt_archivo)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     While Not rsaux2.EOF
                           var_codigo = rsaux2!SEGMENT1
                           For var_j = 1 To Me.lv_salidas.ListItems.Count
                               Me.lv_salidas.ListItems.Item(var_j).Selected = True
                               If Me.lv_salidas.selectedItem = var_codigo And CDbl(Me.lv_salidas.selectedItem.SubItems(6)) = CDbl(rsaux2!delivery_detail_id) Then
                                  Me.lv_salidas.selectedItem.SubItems(3) = Format(rsaux2!FLOA_SAL_CANTIDAD_LEIDA, "###,###,##0.00")
                                  Me.lv_salidas.selectedItem.SubItems(4) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(2)) - CDbl(Me.lv_salidas.selectedItem.SubItems(3)), "###,###,##0.00")
                               End If
                           Next var_j
                           var_cantidad_recibida = var_cantidad_recibida + rsaux2!FLOA_SAL_CANTIDAD_LEIDA
                           rsaux2.MoveNext
                     Wend
                     rsaux2.Close
                     Me.lbl_recibidos = Format(var_cantidad_recibida, "###,###,##0.00")
                     Me.txt_codigo.Enabled = True
                     Me.txt_codigo.SetFocus
                  Else
                     If rsaux!inte_emb_embarque = CDbl(Me.txt_embarque) Or rsaux.EOF Then
                        If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rs!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                           If var_pedido_tienda = 0 Then
                              rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rs!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 Me.txt_agente = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                              Else
                                 rsaux9.Open "SELECT * FROM XXVIA_VW_AGENTES WHERE CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 Me.txt_agente = IIf(IsNull(rsaux9!Name), "", rsaux9!Name)
                                 rsaux9.Close
                              End If
                              rsaux2.Close
                           Else
                              Me.txt_agente = IIf(IsNull(rs!attribuete9), "", rs!ATTRIBUTE9)
                           End If
                        Else
                           rsaux9.Open "SELECT * FROM XXVIA_VW_AGENTES WHERE CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           Me.txt_agente = IIf(IsNull(rsaux9!Name), "", rsaux9!Name)
                           rsaux9.Close
                        End If
                        
                        
                        Me.txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                        Me.txt_origen = IIf(IsNull(rs!subinventory), "", rs!subinventory)
                        Me.lv_salidas.ListItems.Clear
                        var_cantidad_enviada = 0
                        While Not rs.EOF
                              Set list_item = lv_salidas.ListItems.Add(, , rs!SEGMENT1)
                              list_item.SubItems(1) = IIf(IsNull(rs!Description), "", rs!Description)
                              list_item.SubItems(2) = Format(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity), "###,###,##0.00")
                              var_cantidad_enviada = var_cantidad_enviada + IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)
                              list_item.SubItems(3) = 0
                              list_item.SubItems(4) = Format(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity), "###,###,##0.00")
                              list_item.SubItems(5) = IIf(IsNull(rs!inventory_item_id), 0, rs!inventory_item_id)
                              list_item.SubItems(6) = IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)
                              list_item.SubItems(7) = IIf(IsNull(rs!SOURCE_LINE_NUMBER), 0, rs!SOURCE_LINE_NUMBER)
                              list_item.SubItems(8) = IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)
                              rs.MoveNext
                        Wend
                        var_cantidad_recibida = 0
                        rsaux2.Open "SELECT * FROM XXVIA_TB_SALIDAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND  source_header_number = " + CStr(CDbl(Me.txt_archivo)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                        
                              var_codigo = rsaux2!SEGMENT1
                              For var_j = 1 To Me.lv_salidas.ListItems.Count
                                  Me.lv_salidas.ListItems.Item(var_j).Selected = True
                                  If Me.lv_salidas.selectedItem = var_codigo And CDbl(Me.lv_salidas.selectedItem.SubItems(6)) = CDbl(rsaux2!delivery_detail_id) Then
                                     Me.lv_salidas.selectedItem.SubItems(3) = Format(rsaux2!FLOA_SAL_CANTIDAD_LEIDA, "###,###,##0.00")
                                     Me.lv_salidas.selectedItem.SubItems(4) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(2)) - CDbl(Me.lv_salidas.selectedItem.SubItems(3)), "###,###,##0.00")
                                  End If
                              Next var_j
                              var_cantidad_recibida = var_cantidad_recibida + rsaux2!FLOA_SAL_CANTIDAD_LEIDA
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        Me.lbl_recibidos = Format(var_cantidad_recibida, "###,###,##0.00")
                        Me.lbl_enviados = Format(var_cantidad_enviada, "###,###,##0.00")
                        Me.txt_archivo.Enabled = False
                        Me.txt_codigo.Enabled = True
                        Me.txt_codigo.SetFocus
                     Else
                        rsaux1.Open "select * from TB_ORACLE_EMBARQUES_ORDENES where source_header_number = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
                        MsgBox "La orden de surtido se encuentra en el embarque " + CStr(rsaux1!inte_emb_embarque), vbOKOnly, "ATENCION"
                        rsaux1.Close
                        Me.txt_agente = ""
                        Me.txt_archivo = ""
                        Me.txt_cliente = ""
                        Me.txt_origen = ""
                        Me.lbl_enviados = ""
                        Me.lbl_recibidos = ""
                        Me.txt_codigo.Enabled = False
                        Me.lv_salidas.ListItems.Clear
                     End If
                  End If
                  rsaux.Close
                  Call ilumina_grid
               Else
                  MsgBox "El pedido corresponde a ventas por telefono y no a sido liberado", vbOKOnly, "ATENCION"
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
            Me.txt_codigo.Enabled = False
            Me.lv_salidas.ListItems.Clear
         End If
         rs.Close
      Else
         MsgBox "Número de orden de surtido incorrecta", vbOKOnly, "ATENCION"
         Me.txt_agente = ""
         Me.txt_archivo = ""
         Me.txt_cliente = ""
         Me.txt_origen = ""
         Me.txt_codigo.Enabled = False
         Me.lv_salidas.ListItems.Clear
      End If
   Else
      MsgBox "El embarque ya fue cerrado", vbOKOnly, "ATENCION"
   End If
End Sub


Private Sub cmd_aceptar_sello_Click()
   If Trim(txt_sello) <> "" Then
      rs.Open "insert into tb_Sellos (inte_emb_embarque, vcha_Sel_Sello) values (" + Me.txt_embarque + ",'" + Me.txt_sello + "')", cnn, adOpenDynamic, adLockOptimistic
      Set list_item = lv_sellos.ListItems.Add(, , txt_sello)
      Me.txt_sello = ""
      Me.txt_sello.SetFocus
   Else
      MsgBox "No se indico un sello", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_cancelar_sello_Click()
   Me.frm_sellos.Visible = False
End Sub

Private Sub cmd_cerrar_Click()
Dim clnt As New SoapClient30
Dim var_arreglo() As String
Dim var_trip_id As String
Dim var_b As Boolean
Dim var_con As String
rs.Open "SELECT CHAR_EMB_ESTATUS FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
VAR_ESTATUS = IIf(IsNull(rs(0).Value), "", rs(0).Value)
rs.Close
If VAR_ESTATUS = "" Then
   x = 1
Else
   x = 0
End If
If x = 1 Then
   var_si = MsgBox("¿Desea cerrar el embarque?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_si = MsgBox("Confirmar el cerrado del embarque", vbYesNo, "ATENCION")
      If var_si = 6 Then
         If rs.State = 1 Then
            rs.Close
         End If
         If Me.txt_archivo = "" Then
            rs.Open "SELECT * FROM XXVIA_TB_SALIDAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               Me.txt_archivo = rs!source_header_number
            End If
            rs.Close
         End If
         var_orden = CDbl(Me.txt_archivo)
         rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If rs!tipo_embarque = 1 Then
            rsaux.Open "select distinct source_header_number from xxvia_tb_salidas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
         var_Cadena_pedidos = ""
         var_j = 0
         While Not rsaux.EOF
               If var_Cadena_pedidos = "" Then
                  var_Cadena_pedidos = "'" + CStr(rsaux!source_header_number) + "'"
               Else
                  var_Cadena_pedidos = var_Cadena_pedidos + ", '" + CStr(rsaux!source_header_number) + "'"
               End If
               var_j = var_j + 1
         

               var_cadena = "SELECT a.source_header_type_name, oha.source_document_id, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND A.SOURCE_HEADER_NUMBER = '" + CStr(rsaux!source_header_number) + "'"
               var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y'"
               
               
               
               
               rsaux4.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               rsaux4.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                  While Not rsaux4.EOF
                        rsaux3.Open "SELECT * FROM XXVIA_TB_SALIDAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND source_header_number = " + CStr(CDbl(rsaux4!source_header_number)) + " AND DELIVERY_DETAIL_ID = " + CStr(IIf(IsNull(rsaux4!delivery_detail_id), 0, rsaux4!delivery_detail_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If rsaux3.EOF Then
                           var_cadena = "INSERT INTO XXVIA_TB_SALIDAS (INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER, SEGMENT1, FLOA_SAL_CANTIDAD_LEIDA, INVENTORY_ITEM_ID, DELIVERY_DETAIL_ID, SOURCE_LINE_NUMBER, DELIVERY_ID, CONSECUTIVO)"
                           var_cadena = var_cadena + " values (" + Me.txt_embarque + "," + CStr(CDbl(rsaux!source_header_number)) + ",'" + rsaux4!SEGMENT1 + "',0," + CStr(IIf(IsNull(rsaux4!inventory_item_id), 0, rsaux4!inventory_item_id)) + "," + CStr(IIf(IsNull(rsaux4!delivery_detail_id), 0, rsaux4!delivery_detail_id)) + ",'" + CStr(IIf(IsNull(rsaux4!SOURCE_LINE_NUMBER), 0, rsaux4!SOURCE_LINE_NUMBER)) + "'," + CStr(IIf(IsNull(rsaux4!delivery_id), 0, rsaux4!delivery_id)) + ",0) "
                           rsaux5.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux3.Close
                        rsaux4.MoveNext
                  Wend
               End If
               rsaux4.Close
               
               rsaux.MoveNext
         Wend
         rsaux.Close
         
         
         
         
         
         
         
         
         If var_Cadena_pedidos <> "" Then
         x = 0
         If x = 1 Then
            
            var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, E.NAME from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_Cadena_pedidos + ")"
            var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
            
            
            
            
            rsaux9.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux9.EOF Then
               rsaux9.Close
               var_sigue = 1
               var_n = 0
               While var_sigue = 1
                     var_n = var_n + 1
                     If rsaux8.State = 1 Then
                        rsaux8.Close
                     End If
                     var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, E.NAME from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_Cadena_pedidos + ")"
                     var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
                     
                     
                     
                     
                     rsaux8.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If rsaux8.EOF Then
                        var_sigue = 0
                     Else
                       rsaux10.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                       If Not rsaux10.EOF Then
                          VAR_USER_ID = rsaux10!user_id
                          VAR_RESP_ID = rsaux10!resp_id
                          VAR_RESP_APPL_ID = rsaux10!resp_appl_id
                       End If
                       rsaux10.Close
                     
                        While Not rsaux8.EOF
                              'MsgBox rsaux8!SOURCE_HEADER_NUMBER
                              rsaux7.Open "SELECT * FROM TB_ORACLE_NEGADO WHERE PEDIDO IN (" + CStr(rsaux8!source_header_number) + ") AND INVENTORY_ITEM_ID = " + CStr(rsaux8!inventory_item_id), cnn, adOpenDynamic, adLockOptimistic
                              rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              rsaux6.Open "call XXVIA_SP_DEPURA_ORDEN_SURTIDO (" + CStr(CDbl(rsaux8!header_id)) + ", " + CStr(CDbl(rsaux8!source_LINE_ID)) + ", 'PRODUCCION'," + CStr(VAR_USER_ID) + "," + CStr(VAR_RESP_ID) + "," + CStr(VAR_RESP_APPL_ID) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              rsaux7.Close
                              rsaux8.MoveNext
                        Wend
                     End If
                     If var_n = 100 Then
                        var_sigue = 0
                     End If
                     rsaux8.Close
               Wend
            Else
               rsaux9.Close
            End If
         End If
         rs.Close
            
            
            '----------
            
          x = 0
            If x = 1 Then
         rsaux.Open "select distinct source_header_number from xxvia_tb_salidas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_Cadena_pedidos = ""
         var_j = 0
         While Not rsaux.EOF
               'If var_cadena_pedidos = "" Then
               '   var_cadena_pedidos = "'" + CStr(rsaux!source_header_number) + "'"
               'Else
               '   var_cadena_pedidos = var_cadena_pedidos + ", '" + CStr(rsaux!source_header_number) + "'"
               'End If
               'var_j = var_j + 1
         
               var_cadena = "SELECT a.source_header_type_name, oha.source_document_id, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, E.NAME from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER = '" + CStr(rsaux!source_header_number) + "'"
               var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B'"
               
               

               
               
               
               rsaux4.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               rsaux4.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                  While Not rsaux4.EOF
                        rsaux3.Open "SELECT * FROM XXVIA_TB_SALIDAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND source_header_number = " + CStr(CDbl(rsaux4!source_header_number)) + " AND DELIVERY_DETAIL_ID = " + CStr(IIf(IsNull(rsaux4!delivery_detail_id), 0, rsaux4!delivery_detail_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        
                        If rsaux3.EOF Then
                           var_cadena = "INSERT INTO XXVIA_TB_SALIDAS (INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER, SEGMENT1, FLOA_SAL_CANTIDAD_LEIDA, INVENTORY_ITEM_ID, DELIVERY_DETAIL_ID, SOURCE_LINE_NUMBER, DELIVERY_ID, CONSECUTIVO)"
                           var_cadena = var_cadena + " values (" + Me.txt_embarque + "," + CStr(CDbl(rsaux!source_header_number)) + ",'" + rsaux4!SEGMENT1 + "',0," + CStr(IIf(IsNull(rsaux4!inventory_item_id), 0, rsaux4!inventory_item_id)) + "," + CStr(IIf(IsNull(rsaux4!delivery_detail_id), 0, rsaux4!delivery_detail_id)) + ",'" + CStr(IIf(IsNull(rsaux4!SOURCE_LINE_NUMBER), 0, rsaux4!SOURCE_LINE_NUMBER)) + "'," + CStr(IIf(IsNull(rsaux4!delivery_id), 0, rsaux4!delivery_id)) + ",0) "
                           rsaux5.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux3.Close
                        rsaux4.MoveNext
                  Wend
               End If
               rsaux4.Close
               
               rsaux.MoveNext
         Wend
         rsaux.Close
            End If
            
            
            
            
            
            
            '----------
            rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
            VAR_X_TRIP_ID = rs!ARREGLO_0
            var_x_trip_name = rs!ARREGLO_1
            rs.Close
            
            'clnt.MSSoapInit var_webservice
            'rs.Open "SELECT * FROM XXVIA_TB_SALIDAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
            'While Not rs.EOF
            '      'var_b = clnt.ACTUALIZAR_DETALLE(Val(rs!DELIVERY_DETAIL_ID), CDbl(rs!FLOA_sAL_cANTIDAD_LEIDA), "OE", IIf(IsNull(rs!consecutivo), 0, rs!consecutivo))
            '      rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            '      rsaux10.Open "call XXVIA_SP_ACTUALIZA_DETALLE (1.0," + CStr(CDbl(rs!DELIVERY_DETAIL_ID)) + ", " + CStr(CDbl(rs!FLOA_sAL_cANTIDAD_LEIDA)) + ", 'OE'," + CStr(IIf(IsNull(rs!consecutivo), 0, rs!consecutivo)) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
            '      rs.MoveNext
            'Wend
            'rs.Close
            'Set clnt = Nothing
            
            rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If rs.State = 1 Then
               rs.Close
            End If
            'rs.Open "call xxvia_sp_actualiza_detalle(" + Me.txt_embarque + ",1)", cnnoracle_4, adOpenDynamic, adLockOptimistic
            
            
            clnt.MSSoapInit var_webservice
            rs.Open "SELECT delivery_detail_id, MIN(CONSECUTIVO) AS CONSECUTIVO, sum(floa_sal_Cantidad_leida) as floa_sal_Cantidad_leida FROM XXVIA_TB_SALIDAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " group by delivery_detail_id", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux.Open "SELECT * FROM WSH_DELIVERABLES_V WHERE delivery_detail_id = " + CStr(rs!delivery_detail_id) + " AND RELEASED_STATUS = 'Y'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     'var_b = clnt.ACTUALIZAR_DETALLE(Val(rs!delivery_detail_id), CDbl(rs!FLOA_sAL_cANTIDAD_LEIDA), "OE", CDbl(IIf(IsNull(rs!consecutivo), 0, rs!consecutivo)))
                     On Error GoTo salir2:
                     rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     rsaux6.Open "CALL xxvia_pk_interfaces_om.actualizar_detalle (1.0, " + CStr(rs!delivery_detail_id) + "," + CStr(rs!FLOA_SAL_CANTIDAD_LEIDA) + ",'OE'," + CStr(IIf(IsNull(rs!CONSECUTIVO), 0, rs!CONSECUTIVO)) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux.Close
                  rs.MoveNext
            Wend
            rs.Close
            Set clnt = Nothing
            
            
            
            'clnt.MSSoapInit var_webservice
            'rs.Open "SELECT DISTINCT DELIVERY_ID FROM XXVIA_TB_SALIDAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
            'While Not rs.EOF
            '      rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            '      'MsgBox CStr(rs!delivery_id) + " " + CStr(Val(VAR_X_TRIP_ID))
            '      var_arreglo = clnt.ASIGNAR_embarque(rs!delivery_id, Val(VAR_X_TRIP_ID), "CONFIRM")
            '      rs.MoveNext
            'Wend
            'rs.Close
            'Set clnt = Nothing
            rs.Open "UPDATE XXVIA_TB_ENCABEZADO_EMBARQUES SET CHAR_EMB_ESTATUS = 'I', FECHA_FIN = SYSDATE WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
            rs.Open "UPDATE TB_ORACLE_EMBARQUES_ORDENES SET estatus = 'I' WHERE inte_emb_EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
            x = 1
            If x = 0 Then
            rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               If IIf(IsNull(rs!char_Emb_estatus), "", rs!char_Emb_estatus) = "I" Then
                  If rs!tipo_embarque = 1 Then
                      rsaux.Open "select distinct source_header_number from xxvia_tb_salidas where inte_emb_embarque = " + Me.txt_embarque + " order by source_header_number", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  var_Cadena_pedidos = ""
                  var_j = 0
                  While Not rsaux.EOF
                        If var_Cadena_pedidos = "" Then
                           var_Cadena_pedidos = "'" + CStr(rsaux!source_header_number) + "'"
                        Else
                           var_Cadena_pedidos = var_Cadena_pedidos + ", '" + CStr(rsaux!source_header_number) + "'"
                        End If
                        var_j = var_j + 1
                        rsaux2.Open "SELECT * FROM TB_EMBARQUES_PEDIDOS_DESPACHADOS where embarque = " + Me.txt_embarque + " and pedido = " + CStr(rsaux!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        If rsaux2.EOF Then
                           rsaux3.Open "insert into TB_EMBARQUES_PEDIDOS_DESPACHADOS (embarque, pedido) values (" + Me.txt_embarque + "," + CStr(rsaux!source_header_number) + ")", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux2.Close
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  var_i = 0
                  '''' hasta aqui
                  While var_j <> var_i
                        var_i = 0
                        var_cadena = "SELECT e.collector_id, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  E.NAME , sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors e Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id "
                        var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + var_Cadena_pedidos + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id AND released_status = 'C' group by  e.collector_id, A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status,  E.NAME"
                        
                        'var_cadena = "SELECT DISTINCT D.collector_id, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  D.NAME , sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_vw_agentes D Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCAS.CUST_ACCOUNT_ID = D.CUST_ACCOUNT_ID AND A.SOURCE_HEADER_NUMBER in (" + VAR_CADENA_PEDIDOS + ")"
                        'var_cadena = " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND released_status = 'C' group by  D.collector_id, A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status,  D.NAME "
                        
                        rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        While Not rsaux.EOF
                              rsaux2.Open "update TB_EMBARQUES_PEDIDOS_DESPACHADOS set marca = '*' where embarque = " + Me.txt_embarque + " and pedido = " + CStr(rsaux!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              var_i = var_i + 1
                              rsaux.MoveNext
                        Wend
                        rsaux.Close
                        rsaux.Open "select * from TB_EMBARQUES_PEDIDOS_DESPACHADOS where embarque = " + Me.txt_embarque + " and marca is null", cnn, adOpenDynamic, adLockOptimistic
                        var_cadena_pedidos_no_desp = ""
                        While Not rsaux.EOF
                              If var_cadena_pedidos_no_desp = "" Then
                                 var_cadena_pedidos_no_desp = CStr(rsaux!pedido)
                              Else
                                 var_cadena_pedidos_no_desp = var_cadena_pedidos_no_desp + ", " + CStr(rsaux!pedido)
                              End If
                              rsaux.MoveNext
                        Wend
                        rsaux.Close
                        If var_cadena_pedidos_no_desp <> "" Then
                           MsgBox "los pedidos " + var_cadena_pedidos_no_desp + " no se pudieron cerrar en oracle", vbOKOnly, "ATENCION"
                        End If
                  Wend
                  
                  var_cadena_pedidos_global = var_Cadena_pedidos
                  var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos_global + ") "
                  var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
                  
                  'var_cadena = "SELECT DISTINCT a.source_line_id, OHA.HEADER_ID from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, xxvia_vw_agentes D Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCAS.CUST_ACCOUNT_ID = D.CUST_ACCOUNT_ID AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos_global + ")"
                  'var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND  released_status = 'B' order by A.source_header_number"
                  
                  
                  rsaux7.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux7.EOF Then
                     var_tipo_depurado = 1
                     frmoracle_depurar_pedidos.Show 1
                  End If
                  rsaux7.Close
                  var_tipo_depurado = 0
                  
                  var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, E.NAME from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos_global + ")"
                  var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
                  rsaux9.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux9.EOF Then
                     rsaux9.Close
                     var_sigue = 1
                     var_i = 0
                     While var_sigue = 1 And var_i < 50
                           If rsaux8.State = 1 Then
                              rsaux8.Close
                           End If
                           var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, E.NAME from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos_global + ")"
                           var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
                           
                           
                           rsaux8.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If rsaux8.EOF Then
                              var_sigue = 0
                           Else
                              rsaux9.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux9.EOF Then
                                 VAR_USER_ID = rsaux9!user_id
                                 VAR_RESP_ID = rsaux9!resp_id
                                 VAR_RESP_APPL_ID = rsaux9!resp_appl_id
                              End If
                              rsaux9.Close
                              While Not rsaux8.EOF
                                    rsaux7.Open "SELECT * FROM TB_ORACLE_NEGADO WHERE PEDIDO IN (" + CStr(rsaux8!source_header_number) + ") AND INVENTORY_ITEM_ID = " + CStr(rsaux8!inventory_item_id), cnn, adOpenDynamic, adLockOptimistic
                                    rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    'clnt.MSSoapInit var_webservice
                                    'var_s = clnt.cancelar_back_order(CDbl(rsaux8!HEADER_ID), CDbl(rsaux8!SOURCE_LINE_ID), rsaux7!CAUSA_NEGADO)
                                    'Set clnt = Nothing
                                    rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    rsaux10.Open "call XXVIA_SP_DEPURA_ORDEN_SURTIDO (" + CStr(CDbl(rsaux8!header_id)) + ", " + CStr(CDbl(rsaux8!source_LINE_ID)) + ", '" + rsaux7!CAUSA_NEGADO + "'," + CStr(VAR_USER_ID) + "," + CStr(VAR_RESP_ID) + "," + CStr(VAR_RESP_APPL_ID) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    rsaux7.Close
                                    rsaux8.MoveNext
                              Wend
                           End If
                           If rsaux8.State = 1 Then
                              rsaux8.Close
                           End If
                           var_i = var_i + 1
                           If var_i = 100 Then
                              MsgBox "No se depuraron las ordenes de surtido"
                           End If
                     Wend
                  Else
                     rsaux9.Close
                  End If
               End If
            End If
            End If
            Me.txt_codigo.Enabled = False
            'MsgBox "Se a cerrado el embarque", vbOKOnly, "ATENCION"
            Me.frm_sellos.Visible = False
            If rs.State = 1 Then
              rs.Close
            End If
            
            '--------------- confirmar pedidos
            
            
   rsaux.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      VAR_X_TRIP_ID = rs!ARREGLO_0
      var_x_trip_name = rs!ARREGLO_1
      VAR_ESTATUS = IIf(IsNull(rs!char_Emb_estatus), "", rs!char_Emb_estatus)
      If IIf(IsNull(rs!char_Emb_estatus), "", rs!char_Emb_estatus) = "I" Then
         If rs!tipo_embarque = 1 Then
            rsaux.Open "select distinct source_header_number from xxvia_tb_salidas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
         If rs!tipo_embarque = 2 Then
            rsaux.Open "select distinct source_header_number from xxvia_tb_SAlidas_cajas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
         VAR_CADENA_PEDIDOS_M = ""
         While Not rsaux.EOF
               If VAR_CADENA_PEDIDOS_M = "" Then
                  VAR_CADENA_PEDIDOS_M = CStr(rsaux!source_header_number)
               Else
                  VAR_CADENA_PEDIDOS_M = VAR_CADENA_PEDIDOS_M + ", " + CStr(rsaux!source_header_number)
               End If
               rsaux.MoveNext
         Wend
         var_Cadena_pedidos = ""
         rsaux.MoveFirst
         While Not rsaux.EOF
               rsaux1.Open "select distinct delivery_id from wsh_deliverables_v where SOURCE_HEADER_NUMBER = " + CStr(rsaux!source_header_number) + " AND delivery_id IS NOT NULL", cnnoracle_4, adOpenDynamic, adLockOptimistic
               VAR_ENTREGA = rsaux1!delivery_id
               rsaux1.Close
               rsaux1.Open "select distinct source_header_number from wsh_deliverables_v where delivery_id = " + CStr(VAR_ENTREGA), cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  var_j = 0
                  While Not rsaux1.EOF
                        var_j = var_j + 1
                        rsaux1.MoveNext
                  Wend
                  If var_j > 1 Then
                     If var_Cadena_pedidos = "" Then
                        var_Cadena_pedidos = CStr(rsaux!source_header_number) + " ENTREGA: " + CStr(VAR_ENTREGA)
                     Else
                        var_Cadena_pedidos = var_Cadena_pedidos + ", " + CStr(rsaux!source_header_number) + " ENTREGA: " + CStr(VAR_ENTREGA)
                     End If
                  End If
               End If
               rsaux1.Close
               rsaux.MoveNext
         Wend
         rsaux.MoveFirst
         
         
         If var_Cadena_pedidos <> "" Then
            MsgBox "Los pedidos siguientes tienen dos entregas " + var_Cadena_pedidos
         Else
            cnn.BeginTrans
            rsaux8.Open "SELECT MAX(CONSECUTIVO) FROM TB_ORACLE_COMPARACION_PEDIDO_AFECTACION", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
               var_consecutivo = IIf(IsNull(rsaux8(0).Value), 0, rsaux8(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rsaux8.Close
            rsaux8.Open "insert into TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            
            
            
            rsaux8.Open "SELECT inte_emb_embarque, SOURCE_HEADER_NUMBER, SUM(FLOA_sal_cANTIDAD_LEIDA) AS CANTIDAD FROM XXVIA_TB_sALIDAS where source_header_number in (" + VAR_CADENA_PEDIDOS_M + ") GROUP BY inte_emb_embarque, SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux8.EOF
                  rsaux2.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(rsaux8!inte_emb_embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!Cantidad) + ",0, '" + CStr(rsaux2!FECHA_INiCIO) + "'," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!Cantidad) + ",0, ''," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux2.Close
                  rsaux8.MoveNext
            Wend
            rsaux8.Close
            rsaux8.Open "SELECT inte_emb_embarque, SOURCE_HEADER_NUMBER, SUM(FLOA_sal_cANTIDAD_LEIDA) AS CANTIDAD FROM XXVIA_TB_sALIDAS_CAJAS where source_header_number in (" + VAR_CADENA_PEDIDOS_M + ") GROUP BY inte_emb_embarque, SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux8.EOF
                  rsaux2.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(rsaux8!inte_emb_embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!Cantidad) + ",0, '" + CStr(rsaux2!FECHA_INiCIO) + "'," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!Cantidad) + ",0, ''," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux2.Close
                  rsaux8.MoveNext
            Wend
            rsaux8.Close
            rsaux8.Open "SELECT pedido FROM TB_ORACLE_COMPARACION_PEDIDO_AFECTACION WHERE CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux8.EOF
                  rsaux1.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  rsaux10.Open "SELECT SOURCE_HEADER_NUMBER, SUM(SHIPPED_QUANTITY) AS CANTIDAD FROM WSH_DELIVERABLES_V WHERE SOURCE_HEADER_NUMBER = " + CStr(IIf(IsNull(rsaux8!pedido), 0, rsaux8!pedido)) + " GROUP BY SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux10.EOF Then
                     rsaux1.Open "UPDATE TB_ORACLE_COMPARACION_PEDIDO_AFECTACION SET CANTIDAD_AFECTADA = " + CStr(IIf(IsNull(rsaux10!Cantidad), 0, rsaux10!Cantidad)) + " WHERE PEDIDO = " + CStr(rsaux8!pedido) + " AND CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux10.Close
                  rsaux8.MoveNext
            Wend
            rsaux8.Close
            rsaux8.Open "SELECT *  FROM TB_ORACLE_COMPARACION_PEDIDO_AFECTACION where cantidad_afectada > 0 and CANTIDAD_LEIDA <> cantidad_afectada AND CONSECUTIVO = " + CStr(var_consecutivo) + " order by PEDIDO desc "
            If Not rsaux8.EOF Then
               var_cadena_pedidos_mal = ""
               While Not rsaux8.EOF
                     If var_cadena_pedidos_mal = "" Then
                        var_cadena_pedidos_mal = CStr(IIf(IsNull(rsaux8!pedido), 0, rsaux8!pedido))
                     Else
                        var_cadena_pedidos_mal = var_cadena_pedidos_mal + ", " + CStr(IIf(IsNull(rsaux8!pedido), 0, rsaux8!pedido))
                     End If
                     rsaux8.MoveNext
               Wend
               MsgBox "Los siguientes pedidos tienen errores entra la cantidad leida y la cantidad afectada: " + CStr(var_cadena_pedidos_mal), vbOKOnly, "ATENCION"
            Else
               clnt.MSSoapInit "http://intranet/WsEBS12Prod/wsInterfaceOM.asmx?wsdl"
               While Not rsaux.EOF
                     rsaux2.Open "select distinct delivery_id from wsh_deliverables_v where SOURCE_HEADER_NUMBER = " + CStr(rsaux!source_header_number) + " AND delivery_id IS NOT NULL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     While Not rsaux2.EOF
                           VAR_ENTREGA = rsaux2!delivery_id
                           rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           VAR_ESTATUS = 0
                           On Error GoTo salirc:
                           var_arreglo = clnt.ASIGNAR_embarque(VAR_ENTREGA, Val(VAR_X_TRIP_ID), "CONFIRM")
                           rsaux1.Open "insert into tb_oracle_pedidos_confirmados (pedido, fecha, maquina, error) values (" + CStr(rsaux!source_header_number) + ", getdate(), '" + fun_NombrePc + "'," + CStr(VAR_ESTATUS) + ")", cnn, adOpenDynamic, adLockOptimistic
                           rsaux2.MoveNext
                     Wend
                     rsaux2.Close
                     rsaux.MoveNext
               Wend
               Set clnt = Nothing
               MsgBox "Se termino de cerrar el embarque", vbOKOnly, "ATENCION"
            End If
            rsaux8.Close
         End If
         rsaux.Close
      Else
         If VAR_ESTATUS = "F" Then
            MsgBox "EL embarque ya fue facturado"
         Else
            MsgBox "El embarque NO a sido cerrado", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   rs.Close
            
            
            
            '--------------- fin de confirmar pedidos
            
         Else
            Me.frm_sellos.Visible = False
            MsgBox "El embarque esta vacio", vbOKOnly, "ATENCION"
            If rs.State = 1 Then
              rs.Close
            End If
         End If
      End If
   End If
   Else
      MsgBox "El embarque ya habia sido cerrado", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salirc:
   If Err.Number = -2147467259 Then
      MsgBox Err.Description
      Resume Next
      VAR_ESTATUS = 1
   End If

salir2:
   If Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      Resume
   End If
End Sub

Private Sub cmd_cerrar_embarque_Click()
   Me.lv_sellos.ListItems.Clear
   rs.Open "select * from tb_Sellos where inte_emb_embarque = " + Str(var_numero_embarque), cnn, adOpenDynamic, adLockOptimistic
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
   Me.frm_sellos.Visible = True
   Me.txt_sello = ""
   Me.txt_sello.SetFocus
End Sub

Private Sub cmd_mensaje_2_Click()
   Me.wmp2.Controls.play
End Sub

Private Sub cmd_mensaje_4_Click()
   Me.wmp4.Controls.play
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_agente = ""
   Me.txt_archivo = ""
   Me.txt_cliente = ""
   Me.txt_origen = ""
   Me.txt_codigo.Enabled = False
   Me.lv_salidas.ListItems.Clear
   Me.txt_archivo.Enabled = True
   Me.lbl_enviados = ""
   Me.lbl_recibidos = ""
   Me.txt_archivo.SetFocus
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
   frm_sellos.Visible = False
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


Private Sub txt_archivo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_archivo) Then
         Call ejecuta
      Else
         MsgBox "Orden de surtido incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad_eliminar) Then
         If CDbl(Me.lv_salidas.selectedItem.SubItems(3)) - CDbl(Me.txt_cantidad_eliminar) >= 0 Then
            Me.lv_salidas.selectedItem.SubItems(3) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(3)) - CDbl(Me.txt_cantidad_eliminar), "###,###,##0.00")
            Me.lv_salidas.selectedItem.SubItems(4) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(4)) + CDbl(Me.txt_cantidad_eliminar), "###,###,##0.00")
            Me.lbl_recibidos = Format(CDbl(Me.lbl_recibidos) - CDbl(Me.txt_cantidad_eliminar), "###,###,##0.00")
            rsaux.Open "update XXVIA_TB_SALIDAS set FLOA_SAL_CANTIDAD_LEIDA = FLOA_SAL_CANTIDAD_LEIDA - " + Me.txt_cantidad_eliminar + " where inte_emb_embarque = " + Me.txt_embarque + " and SOURCE_HEADER_NUMBER = " + CStr(CDbl(Me.txt_archivo)) + " and DELIVERY_DETAIL_ID = " + Me.lv_salidas.selectedItem.SubItems(6), cnnoracle_4, adOpenDynamic, adLockOptimistic
            rsaux5.Open "update TB_DETALLE_EQUIPOS_ORDEN_SURTIDO set FLOA_ORS_CANTIDAD_SURTIDA = isnull(FLOA_ORS_CANTIDAD_SURTIDA,0) - " + CStr(Me.txt_cantidad_eliminar) + " where INTE_ORS_ORDEN_SURTIDO = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
            
            Call cantidad_leida_por_persona(CDbl(txt_cantidad_eliminar), "-")
            rsaux.Open "INSERT INTO XXVIA_TB_BITACORA_LECTURA (PEDIDO, CODIGO, USUARIO, CANTIDAD, FECHA_HORA) VALUES (" + Me.txt_archivo + ", '" + Me.lv_salidas.selectedItem + "','" + var_clave_usuario_global + "',-" + CStr(txt_cantidad_eliminar) + ",SYSDATE)", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Me.txt_codigo.Enabled = True Then
               Me.txt_codigo.SetFocus
            End If
         Else
            MsgBox "Cantidad a eliminar incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Cantidad a eliminar incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   Me.frm_eliminar.Visible = False
End Sub


Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Mid(Me.txt_codigo, 1, 2) = "CA" Then
         rs.Open "SELECT * FROM XXVIA_TB_CAJAS_PROD WHERE vcha_caj_caja_id = '" + UCase(Me.txt_codigo) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + rs!VCHA_ART_ARTICULO_ID + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
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
            frmmensaje.lbl_mensaje = "La caja no existe"
            frmmensaje.Show 1
         End If
         rs.Close
      Else
         var_encontro = 0
         var_cantidad_leida = 1
         If Me.txt_codigo <> "" Then
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
               For var_j = 1 To Me.lv_salidas.ListItems.Count
                   lv_salidas.ListItems.Item(var_j).Selected = True
                   If Me.txt_codigo = lv_salidas.selectedItem And (CDbl(Me.lv_salidas.selectedItem.SubItems(4)) - var_cantidad_leida) >= 0 Then
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
               var_cadena = "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, nvl(a.description,'') as localizador, nvl(attribute1,1) as cantidad FROM mtl_cross_references_v A, xxvia_system_items_b B Where a.inventory_item_id = B.inventory_item_id AND B.organization_id = " + var_unidad_organizacional + " AND CROSS_REFERENCE LIKE '" + Me.txt_codigo + "'"
               rsaux9.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux9.EOF Then
                  Me.txt_codigo = IIf(IsNull(rsaux9!SEGMENT1), "", rsaux9!SEGMENT1)
                  rsaux10.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux10.EOF Then
                     var_salida_masiva = IIf(IsNull(rsaux10!attribute10), "N", rsaux10!attribute10)
                     If var_salida_masiva = "Y" Then
                        var_codigo_global = Me.txt_codigo
                        frmoracle_cantidad.Show 1
                        var_cantidad_leida = var_cantidad_global
                        Me.txt_codigo = var_codigo_global
                     Else
                        var_cantidad_leida = IIf(IsNull(rsaux9!Cantidad), 1, rsaux9!Cantidad)
                        'var_cantidad_leida = 1
                     End If
                     For var_j = 1 To Me.lv_salidas.ListItems.Count
                         lv_salidas.ListItems.Item(var_j).Selected = True
                         If Me.txt_codigo = lv_salidas.selectedItem And (CDbl(Me.lv_salidas.selectedItem.SubItems(4)) - var_cantidad_leida) >= 0 Then
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
                     frmmensaje.lbl_mensaje = "El artículo no existe"
                     frmmensaje.Show 1
                  End If
                  rsaux10.Close
               Else
                  Call cmd_mensaje_2_Click
                  txt_codigo = ""
                  frmmensaje.lbl_articulo = ""
                  frmmensaje.lbl_mensaje = "El artículo no existe"
                  frmmensaje.Show 1
               End If
               rsaux9.Close
            End If
            rsaux8.Close
         Else
            If var_localizador = 2 And Me.txt_codigo = "" Then
               txt_codigo = ""
               frmmensaje.lbl_mensaje = "El artículo necesita localizador"
               frmmensaje.Show
            Else
               txt_codigo = ""
               frmmensaje.lbl_mensaje = "El artículo no existe"
               frmmensaje.Show
            End If
         End If
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   If Trim(Me.txt_codigo) <> "" Then
      If var_encontro > 0 Then
         If rsaux1.State = 1 Then
            rsaux1.Close
         End If
         rsaux1.Open "SELECT * FROM TB_ORACLE_EMBARQUES_ORDENES WHERE source_header_number = " + CStr(CDbl(Me.txt_archivo)), cnn, adOpenDynamic, adLockOptimistic
         If rsaux1.EOF Then
            rs.Open "select * from tb_oracle_embarques_ordenes where  INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND source_header_number = " + CStr(CDbl(Me.txt_archivo)), cnn, adOpenDynamic, adLockOptimistic
            If rs.EOF Then
               rsaux.Open "INSERT INTO TB_ORACLE_EMBARQUES_ORDENES (INTE_EMB_EMBARQUE, source_header_number) VALUES (" + Me.txt_embarque + "," + CStr(CDbl(Me.txt_archivo)) + ")", cnn, adOpenDynamic, adLockOptimistic
            End If
            rs.Close
            rs.Open "SELECT * FROM XXVIA_TB_SALIDAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND source_header_number = " + CStr(CDbl(Me.txt_archivo)) + " AND DELIVERY_DETAIL_ID = " + lv_salidas.selectedItem.SubItems(6) + "", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If rs.EOF Then
               rsaux.Open "SELECT MAX(CONSECUTIVO) FROM XXVIA_TB_SALIDAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND source_header_number = " + CStr(CDbl(Me.txt_archivo)), cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
               Else
                  var_consecutivo = 1
               End If
               rsaux.Close
               var_cadena = "INSERT INTO XXVIA_TB_SALIDAS (INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER, SEGMENT1, FLOA_SAL_CANTIDAD_LEIDA, INVENTORY_ITEM_ID, DELIVERY_DETAIL_ID, SOURCE_LINE_NUMBER, DELIVERY_ID, CONSECUTIVO)"
               var_cadena = var_cadena + " values (" + Me.txt_embarque + "," + CStr(CDbl(Me.txt_archivo)) + ",'" + Me.txt_codigo + "'," + CStr(var_cantidad_leida) + "," + lv_salidas.selectedItem.SubItems(5) + "," + Me.lv_salidas.selectedItem.SubItems(6) + "," + Me.lv_salidas.selectedItem.SubItems(7) + "," + Me.lv_salidas.selectedItem.SubItems(8) + "," + CStr(var_consecutivo) + ") "
               rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               rsaux.Open "INSERT INTO XXVIA_TB_BITACORA_LECTURA (PEDIDO, CODIGO, USUARIO, CANTIDAD, FECHA_HORA) VALUES (" + Me.txt_archivo + ", '" + Me.txt_codigo + "','" + var_clave_usuario_global + "'," + CStr(var_cantidad_leida) + ",SYSDATE)", cnnoracle_4, adOpenDynamic, adLockOptimistic
            Else
               rsaux.Open "update XXVIA_TB_SALIDAS set FLOA_SAL_CANTIDAD_LEIDA = FLOA_SAL_CANTIDAD_LEIDA + " + CStr(var_cantidad_leida) + " where inte_emb_embarque = " + Me.txt_embarque + " and SOURCE_HEADER_NUMBER = " + CStr(CDbl(Me.txt_archivo)) + " and DELIVERY_DETAIL_ID = '" + Me.lv_salidas.selectedItem.SubItems(6) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               rsaux.Open "INSERT INTO XXVIA_TB_BITACORA_LECTURA (PEDIDO, CODIGO, USUARIO, CANTIDAD, FECHA_HORA) VALUES (" + Me.txt_archivo + ", '" + Me.txt_codigo + "','" + var_clave_usuario_global + "'," + CStr(var_cantidad_leida) + ",SYSDATE)", cnnoracle_4, adOpenDynamic, adLockOptimistic
            End If
            rsaux5.Open "update TB_DETALLE_EQUIPOS_ORDEN_SURTIDO set FLOA_ORS_CANTIDAD_SURTIDA = isnull(FLOA_ORS_CANTIDAD_SURTIDA,0) + " + CStr(var_cantidad_leida) + " where INTE_ORS_ORDEN_SURTIDO = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
            Call cantidad_leida_por_persona(var_cantidad_leida, "+")

            Me.lv_salidas.selectedItem.SubItems(3) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(3)) + var_cantidad_leida, "###,###,##0.00")
            Me.lv_salidas.selectedItem.SubItems(4) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(2)) - var_cantidad_leida, "###,###,##0.00")
            Me.lbl_recibidos = Format(CDbl(Me.lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
            Me.txt_codigo.SetFocus
            rs.Close
            Call cmd_mensaje_4_Click
         Else
            If rsaux1!inte_emb_embarque = CDbl(Me.txt_embarque) Then
               rs.Open "SELECT * FROM XXVIA_TB_SALIDAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND source_header_number = " + CStr(CDbl(Me.txt_archivo)) + " AND DELIVERY_DETAIL_ID = " + Me.lv_salidas.selectedItem.SubItems(6), cnnoracle_4, adOpenDynamic, adLockOptimistic
               If rs.EOF Then
                  rsaux.Open "SELECT MAX(CONSECUTIVO) FROM XXVIA_TB_SALIDAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND source_header_number = " + CStr(CDbl(Me.txt_archivo)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
                  Else
                     var_consecutivo = 1
                  End If
                  rsaux.Close
                  var_cadena = "INSERT INTO XXVIA_TB_SALIDAS (INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER, SEGMENT1, FLOA_SAL_CANTIDAD_LEIDA, INVENTORY_ITEM_ID, DELIVERY_DETAIL_ID, SOURCE_LINE_NUMBER, DELIVERY_ID, CONSECUTIVO)"
                  var_cadena = var_cadena + " values (" + Me.txt_embarque + "," + CStr(CDbl(Me.txt_archivo)) + ",'" + Me.txt_codigo + "'," + CStr(var_cantidad_leida) + "," + lv_salidas.selectedItem.SubItems(5) + "," + Me.lv_salidas.selectedItem.SubItems(6) + ",'" + Me.lv_salidas.selectedItem.SubItems(7) + "'," + Me.lv_salidas.selectedItem.SubItems(8) + "," + CStr(var_consecutivo) + ") "
                  rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  rsaux.Open "INSERT INTO XXVIA_TB_BITACORA_LECTURA (PEDIDO, CODIGO, USUARIO, CANTIDAD, FECHA_HORA) VALUES (" + Me.txt_archivo + ", '" + Me.txt_codigo + "','" + var_clave_usuario_global + "'," + CStr(var_cantidad_leida) + ",SYSDATE)", cnnoracle_4, adOpenDynamic, adLockOptimistic
               Else
                  rsaux.Open "update XXVIA_TB_SALIDAS set FLOA_SAL_CANTIDAD_LEIDA = FLOA_SAL_CANTIDAD_LEIDA + " + CStr(var_cantidad_leida) + " where inte_emb_embarque = " + Me.txt_embarque + " and SOURCE_HEADER_NUMBER = " + CStr(CDbl(Me.txt_archivo)) + " and DELIVERY_DETAIL_ID = " + Me.lv_salidas.selectedItem.SubItems(6), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  rsaux.Open "INSERT INTO XXVIA_TB_BITACORA_LECTURA (PEDIDO, CODIGO, USUARIO, CANTIDAD, FECHA_HORA) VALUES (" + Me.txt_archivo + ", '" + Me.txt_codigo + "','" + var_clave_usuario_global + "'," + CStr(var_cantidad_leida) + ",SYSDATE)", cnnoracle_4, adOpenDynamic, adLockOptimistic
               End If
               rs.Close
               rsaux5.Open "update TB_DETALLE_EQUIPOS_ORDEN_SURTIDO set FLOA_ORS_CANTIDAD_SURTIDA = isnull(FLOA_ORS_CANTIDAD_SURTIDA,0) + " + CStr(var_cantidad_leida) + " where INTE_ORS_ORDEN_SURTIDO = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
               Call cantidad_leida_por_persona(var_cantidad_leida, "+")
               Me.lv_salidas.selectedItem.SubItems(3) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(3)) + var_cantidad_leida, "###,###,##0.00")
               Me.lv_salidas.selectedItem.SubItems(4) = Format(CDbl(Me.lv_salidas.selectedItem.SubItems(4)) - var_cantidad_leida, "###,###,##0.00")
               Me.lbl_recibidos = Format(CDbl(Me.lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
               Call cmd_mensaje_4_Click
               Me.txt_codigo.SetFocus
            Else
               Call cmd_mensaje_2_Click
               txt_codigo = ""
               rsaux1.Open "SELECT dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_EMBARQUES.INTE_JAU_JAULA_ID, dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID, dbo.TB_USUARIOS.VCHA_USU_NOMBRE, dbo.TB_ENCABEZADO_EMBARQUES.VCHA_AUD_MAQUINA, dbo.Tb_usuarios.VCHA_USU_APELLIDOS FROM dbo.TB_ENCABEZADO_EMBARQUES INNER JOIN dbo.TB_USUARIOS ON dbo.TB_ENCABEZADO_EMBARQUES.VCHA_AUD_USUARIO = dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID Where (dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE = " + CStr(rsaux!inte_emb_embarque) + ")", cnn, adOpenDynamic, adLockOptimistic
               frmmensaje.lbl_articulo = "La orden de surtido se encuentra en el embarque " + CStr(rsaux1!inte_emb_embarque)
               frmmensaje.lbl_mensaje = " en la máquina " + IIf(IsNull(rsaux1!vcha_aud_maquina), "", rsaux1!vcha_aud_maquina) + " con el usuario " + IIf(IsNull(rsaux1!VCHA_USU_NOMBRE), "", rsaux1!VCHA_USU_NOMBRE) + " " + IIf(IsNull(rsaux1!VCHA_USU_APELLIDOS), "", rsaux1!VCHA_USU_APELLIDOS)
               rsaux1.Close
               Me.txt_codigo.Enabled = False
               frmmensaje.Show 1
            End If
            var_renglon = Me.lv_salidas.selectedItem.Index
            Call ilumina_grid
         End If
         rsaux1.Close
         Me.txt_codigo = ""
         var_encontro = 0
        
      End If
   End If
End Sub

Private Sub txt_foco_LostFocus()
   Me.txt_foco.Enabled = False
End Sub

Private Sub txt_sello_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_aceptar_sello.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.frm_sellos.Visible = False
   End If
End Sub
