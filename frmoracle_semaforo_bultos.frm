VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_semaforo_bultos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Semaforo bultos"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_embarque 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1890
      TabIndex        =   0
      Top             =   165
      Width           =   2010
   End
   Begin VB.Frame Frame11 
      Height          =   2280
      Left            =   120
      TabIndex        =   24
      Top             =   6600
      Width           =   15090
      Begin MSComctlLib.ListView lv_ultimos_pedidos 
         Height          =   1740
         Left            =   60
         TabIndex        =   27
         Top             =   480
         Width           =   14880
         _ExtentX        =   26247
         _ExtentY        =   3069
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   " Código"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pedido"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Hora Creación"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Hora aduana"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cliente"
            Object.Width           =   5733
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Cantidad"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "estatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Caja"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Tipo empaque"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Caja"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Sello"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   "  Pedidos pendientes"
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   2
         Left            =   30
         TabIndex        =   25
         Top             =   165
         Width           =   15000
      End
   End
   Begin VB.CommandButton cmd_silenciar 
      Caption         =   "Detener"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   165
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   2490
      Left            =   120
      TabIndex        =   12
      Top             =   700
      Width           =   15060
      Begin MSComctlLib.ListView lv_cajas 
         Height          =   1860
         Left            =   60
         TabIndex        =   13
         Top             =   495
         Width           =   14880
         _ExtentX        =   26247
         _ExtentY        =   3281
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   " Código"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pedido"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Hora Creación"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Hora aduana"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cliente"
            Object.Width           =   5733
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Cantidad"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "estatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Caja"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Tipo empaque"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Caja"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Sello"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   "  Pedido actual"
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   135
         Width           =   14970
      End
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   11220
      TabIndex        =   5
      Top             =   3100
      Width           =   3855
      Begin VB.TextBox txt_cantidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3930
         TabIndex        =   6
         Top             =   165
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Semáforo:"
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
         Left            =   105
         TabIndex        =   8
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label lbl_semaforo 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1185
         TabIndex        =   7
         Top             =   210
         Width           =   2520
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   8000
      Left            =   12255
      Top             =   0
   End
   Begin VB.CommandButton cmd_mensaje_2 
      Caption         =   "mensaje 2"
      Height          =   195
      Left            =   4725
      TabIndex        =   4
      Top             =   45
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton cmd_mensaje_4 
      Caption         =   "mensaje 4"
      Height          =   195
      Left            =   4890
      TabIndex        =   3
      Top             =   45
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Frame Frame8 
      Height          =   615
      Left            =   165
      TabIndex        =   2
      Top             =   3100
      Width           =   11025
      Begin VB.Label lbl_total_bultos_pedido_surtiendo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   30
         Top             =   120
         Width           =   9735
      End
   End
   Begin VB.Frame Frame9 
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   8880
      Width           =   15105
      Begin VB.Label lbl_bultos_3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   9735
      End
   End
   Begin VB.Frame Frame4 
      Height          =   60
      Left            =   0
      TabIndex        =   9
      Top             =   640
      Width           =   15105
   End
   Begin VB.Frame Frame3 
      Height          =   2520
      Left            =   120
      TabIndex        =   10
      Top             =   3675
      Width           =   15090
      Begin MSComctlLib.ListView lv_cajas_siguientes 
         Height          =   1860
         Left            =   60
         TabIndex        =   26
         Top             =   480
         Width           =   14880
         _ExtentX        =   26247
         _ExtentY        =   3281
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   " Código"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pedido"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Hora Creación"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Hora aduana"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cliente"
            Object.Width           =   5733
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Cantidad"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "estatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Caja"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Tipo empaque"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Caja"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Sello"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   "  Pedidos pendientes"
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   1
         Left            =   30
         TabIndex        =   11
         Top             =   165
         Width           =   15000
      End
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   6120
      Width           =   11150
      Begin VB.Label lbl_bultos_2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   10935
      End
   End
   Begin VB.Frame Frame5 
      Height          =   495
      Left            =   11280
      TabIndex        =   18
      Top             =   6120
      Width           =   3855
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3930
         TabIndex        =   19
         Top             =   165
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lbl_semaforo_2 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1185
         TabIndex        =   21
         Top             =   135
         Width           =   2520
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Semáforo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   165
         Width           =   1080
      End
   End
   Begin VB.Label lbl_tipo_bulto 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5160
      TabIndex        =   29
      Top             =   240
      Width           =   9735
   End
   Begin VB.Label Label3 
      Caption         =   "Embarque:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   28
      Top             =   210
      Width           =   1605
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp4 
      Height          =   375
      Left            =   5280
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   630
      URL             =   "C:\sistemas\desarrollo\INTEGRAL\Mec_Alarm_10.wma"
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
      _cx             =   1111
      _cy             =   661
   End
End
Attribute VB_Name = "frmoracle_semaforo_bultos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub ilumina_grid()
    var_n = lv_cajas.ListItems.Count
    For var_i = 1 To var_n
        lv_cajas.ListItems.Item(var_i).Selected = True
        If Trim(lv_cajas.selectedItem.SubItems(6)) = "L" Or Trim(lv_cajas.selectedItem.SubItems(6)) = "S" Then
           lv_cajas.ListItems.Item(var_i).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(1).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(2).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(3).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(4).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(5).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(6).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(7).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(8).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(9).Bold = False
           lv_cajas.ListItems.Item(var_i).ForeColor = &HC000&
           lv_cajas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HC000&
           lv_cajas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HC000&
           lv_cajas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HC000&
           lv_cajas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HC000&
           lv_cajas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HC000&
           lv_cajas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HC000&
           lv_cajas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HC000&
           lv_cajas.ListItems.Item(var_i).ListSubItems(8).ForeColor = &HC000&
           lv_cajas.ListItems.Item(var_i).ListSubItems(9).ForeColor = &HC000&
        Else
           lv_cajas.ListItems.Item(var_i).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(1).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(2).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(3).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(4).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(5).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(6).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(7).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(8).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(9).Bold = False
           lv_cajas.ListItems.Item(var_i).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(8).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(9).ForeColor = &HFF&
        End If
    Next var_i
    If var_renglon > 0 Then
       If var_renglon <= var_n Then
          var_i = var_renglon
          lv_cajas.ListItems.Item(var_i).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(3).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(4).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(5).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(6).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(7).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(8).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(9).Bold = True
          lv_cajas.ListItems.Item(var_i).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(8).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(9).ForeColor = &H8000&
       End If
    End If
    lv_cajas.Refresh
End Sub










Private Sub cmd_mensaje_4_Click()
   Me.wmp4.Controls.play
End Sub


Private Sub cmd_silenciar_Click()
   If Me.cmd_silenciar.Caption = "Detener" Then
      Me.Timer1.Enabled = False
      Me.cmd_silenciar.Caption = "Seguir"
   Else
      Me.Timer1.Enabled = True
      Me.cmd_silenciar.Caption = "Detener"
   End If
   
   
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub





Private Sub Command2_Click()
End Sub

Private Sub Command4_Click()
   If IsNumeric(Me.txt_embarque) Then
      Me.Timer1.Enabled = True
   End If
End Sub

Private Sub Form_Load()
   Me.cmd_silenciar.Caption = "Seguir"
   Me.Timer1.Enabled = True
   If rsaux.State = 1 Then
      rsaux.Close
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub Label4_Click()

End Sub

Private Sub lv_cajas_LostFocus()
   Me.Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
   If Me.cmd_silenciar.Caption = "Detener" Then
      If rs.State = 1 Then
         rs.Close
      End If
      If IsNumeric(Me.txt_embarque) Then
         var_Cadena_pedidos = ""
         var_total_bultos = 0
         var_total_leidos = 0
         rs.Open "SELECT count(*) FROM tb_oracle_cajas_aduana WHERE EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
           var_total_bultos = Format(IIf(IsNull(rs(0).Value), 0, rs(0).Value), "###,###,##0")
         rs.Close
         
         rs.Open "SELECT count(*) FROM tb_oracle_cajas_aduana WHERE ESTATUS in ('L','S') AND EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
            var_total_leidos = Format(IIf(IsNull(rs(0).Value), 0, rs(0).Value), "###,###,##0")
         rs.Close
         
         rs.Open "SELECT SUM(PIEZAS) FROM tb_oracle_cajas_aduana WHERE ESTATUS in ('L','S') AND EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
         Me.txt_cantidad = Format(IIf(IsNull(rs(0).Value), 0, rs(0).Value), "###,###,##0.00")
         rs.Close
         var_contador = 1
         var_contador_2 = 0
         var_contador_3 = 0
         rs.Open "select top 2 AGENTE, nombre_agente, pedido, cliente, orden_pedido, ESTATUS, estatus_pedido from tb_oracle_pedidos_asignados_embarques WHERE EMBARQUE = " + Me.txt_embarque + " AND ((ISNULL(ESTATUS,'') = '')) order by orden_pedido, consecutivo_tabla", cnn, adOpenDynamic, adLockOptimistic
         'rs.Open "select top 2 AGENTE, nombre_agente, pedido, cliente, orden_pedido, ESTATUS, estatus_pedido from tb_oracle_pedidos_asignados_embarques WHERE EMBARQUE = " + Me.txt_embarque + " order by orden_pedido", cnn, adOpenDynamic, adLockOptimistic
         Me.lv_cajas.ListItems.Clear
         Me.lv_cajas_siguientes.ListItems.Clear
         var_Cadena_pedidos = ""
         var_cadena_ultimos_pedidos = ""
         While Not rs.EOF
               If var_Cadena_pedidos = "" Then
                  var_Cadena_pedidos = CStr(rs!pedido)
               Else
                  var_Cadena_pedidos = var_Cadena_pedidos + "," + CStr(rs!pedido)
               End If
               If var_contador = 1 Then
                  var_pedido = rs!pedido
                  var_cadena_ultimos_pedidos = CStr(var_pedido)
                  'rsaux10.Open "select * from tb_oracle_cajas_aduana where embarque = " + Me.txt_embarque + " and pedido = " + CStr(var_pedido) + " and (ISNULL(estatus,'') <> 'S')", cnn, adOpenDynamic, adLockOptimistic
                  rsaux10.Open "select * from tb_oracle_cajas_aduana where embarque = " + Me.txt_embarque + " and pedido = " + CStr(var_pedido), cnn, adOpenDynamic, adLockOptimistic
                  var_contador_total_1 = 0
                  var_Contador_leido_1 = 0
                  While Not rsaux10.EOF
                        var_contador_total_1 = var_contador_total_1 + 1
                        If IIf(IsNull(rs!estatus_pedido), 0, rs!estatus_pedido) = 1 Then
                           Me.lbl_semaforo.BackColor = &HC000&
                        Else
                           Me.lbl_semaforo.BackColor = &HC0&
                        End If
                        Set list_item = Me.lv_cajas.ListItems.Add(, , rsaux10!Caja)
                        list_item.SubItems(1) = IIf(IsNull(rsaux10!pedido), "", rsaux10!pedido)
                        list_item.SubItems(2) = IIf(IsNull(rsaux10!fecha_Creacion), "", rsaux10!fecha_Creacion)
                        list_item.SubItems(3) = IIf(IsNull(rsaux10!fecha_aduana), "", rsaux10!fecha_aduana)
                        list_item.SubItems(4) = IIf(IsNull(rsaux10!Cliente), "", rsaux10!Cliente)
                        list_item.SubItems(5) = Format(rsaux10!PIEZAS, "###,###,##0.00")
                        list_item.SubItems(6) = IIf(IsNull(rsaux10!estatus), "", rsaux10!estatus)
                        If IIf(IsNull(rsaux10!estatus), "", rsaux10!estatus) = "L" Or IIf(IsNull(rsaux10!estatus), "", rsaux10!estatus) = "S" Then
                           var_Contador_leido_1 = var_Contador_leido_1 + 1
                        End If
                        list_item.SubItems(7) = IIf(IsNull(rsaux10!numero_caja), "", rsaux10!numero_caja)
                        list_item.SubItems(8) = IIf(IsNull(rsaux10!TIPO_EMPAQUE), "", rsaux10!TIPO_EMPAQUE)
                        list_item.SubItems(9) = IIf(IsNull(rsaux10!caja_pedido), "", rsaux10!caja_pedido)
                        list_item.SubItems(10) = IIf(IsNull(rsaux10!sello), "", rsaux10!sello)
                         
                        rsaux10.MoveNext
                        var_contador_2 = var_contador_2 + 1
                  Wend
                  rsaux10.Close
                  Me.lbl_total_bultos_pedido_surtiendo = "Bultos leidos: " + CStr(var_Contador_leido_1) + " de " + CStr(var_contador_total_1)
                   
                  var_total_aduana = 0
                  For var_j = 1 To Me.lv_cajas.ListItems.Count
                      Me.lv_cajas.ListItems.Item(var_j).Selected = True
                      If Me.lv_cajas.selectedItem.SubItems(3) <> "" Then
                         var_total_aduana = var_total_aduana + 1
                      End If
                  Next var_j
               
               
               Else
                  rsaux10.Open "select * from tb_oracle_cajas_aduana where embarque = " + Me.txt_embarque + " and pedido = " + CStr(rs!pedido) + " and pedido <>" + CStr(var_pedido), cnn, adOpenDynamic, adLockOptimistic
                  If var_cadena_ultimos_pedidos <> "" Then
                     var_cadena_ultimos_pedidos = var_cadena_ultimos_pedidos + "," + CStr(rs!pedido)
                  Else
                     var_cadena_ultimos_pedidos = CStr(rs!pedido)
                  End If
                  'rsaux10.Open "select char_paq_estatus, inte_paq_caja, sum(floa_sal_cantidad_leida) as cantidad from xxvia_tb_Salidas_cajas where source_header_number = " + CStr(rs!pedido) + " and char_paq_estatus = '' and source_header_number <> " + CStr(var_pedido) + "  and inte_emb_embarque = " + Me.txt_embarque + " group by char_paq_estatus, inte_paq_caja", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If IIf(IsNull(rs!estatus_pedido), 0, rs!estatus_pedido) = 1 Then
                     Me.lbl_semaforo_2.BackColor = &HC000&
                  Else
                     Me.lbl_semaforo_2.BackColor = &HC0&
                  End If
               
                  While Not rsaux10.EOF
                        Set list_item = Me.lv_cajas_siguientes.ListItems.Add(, , rsaux10!Caja)
                        list_item.SubItems(1) = IIf(IsNull(rsaux10!pedido), "", rsaux10!pedido)
                        list_item.SubItems(2) = IIf(IsNull(rsaux10!fecha_Creacion), "", rsaux10!fecha_Creacion)
                        list_item.SubItems(3) = IIf(IsNull(rsaux10!fecha_aduana), "", rsaux10!fecha_aduana)
                        list_item.SubItems(4) = IIf(IsNull(rsaux10!Cliente), "", rsaux10!Cliente)
                        list_item.SubItems(5) = Format(rsaux10!PIEZAS, "###,###,##0.00")
                        list_item.SubItems(6) = IIf(IsNull(rsaux10!estatus), "", rsaux10!estatus)
                        list_item.SubItems(7) = IIf(IsNull(rsaux10!numero_caja), "", rsaux10!numero_caja)
                        list_item.SubItems(8) = IIf(IsNull(rsaux10!TIPO_EMPAQUE), "", rsaux10!TIPO_EMPAQUE)
                        list_item.SubItems(9) = IIf(IsNull(rsaux10!caja_pedido), "", rsaux10!caja_pedido)
                        list_item.SubItems(10) = IIf(IsNull(rsaux10!sello), "", rsaux10!sello)
                        rsaux10.MoveNext
                        var_contador_3 = var_contador_3 + 1
                  Wend
                  rsaux10.Close
               End If
               var_contador = var_contador + 1
               rs.MoveNext
         Wend
         rs.Close
      
         var_total_restante = 0
         If var_cadena_ultimos_pedidos <> "" Then
         
            rs.Open "select  AGENTE, nombre_agente, pedido, cliente, orden_pedido, ESTATUS, estatus_pedido from tb_oracle_pedidos_asignados_embarques WHERE EMBARQUE = " + Me.txt_embarque + " and pedido not in (" + var_cadena_ultimos_pedidos + ") order by orden_pedido", cnn, adOpenDynamic, adLockOptimistic
            Me.lv_ultimos_pedidos.ListItems.Clear
            While Not rs.EOF
                  var_pedido = rs!pedido
                  rsaux10.Open "select * from tb_oracle_cajas_aduana where embarque = " + Me.txt_embarque + " and pedido = " + CStr(var_pedido), cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux10.EOF
                        var_total_restante = var_total_restante + 1
                        Set list_item = Me.lv_ultimos_pedidos.ListItems.Add(, , rsaux10!Caja)
                        list_item.SubItems(1) = IIf(IsNull(rsaux10!pedido), "", rsaux10!pedido)
                        list_item.SubItems(2) = IIf(IsNull(rsaux10!fecha_Creacion), "", rsaux10!fecha_Creacion)
                        list_item.SubItems(3) = IIf(IsNull(rsaux10!fecha_aduana), "", rsaux10!fecha_aduana)
                        list_item.SubItems(4) = IIf(IsNull(rsaux10!Cliente), "", rsaux10!Cliente)
                        list_item.SubItems(5) = Format(rsaux10!PIEZAS, "###,###,##0.00")
                        list_item.SubItems(6) = IIf(IsNull(rsaux10!estatus), "", rsaux10!estatus)
                        list_item.SubItems(7) = IIf(IsNull(rsaux10!numero_caja), "", rsaux10!numero_caja)
                        list_item.SubItems(8) = IIf(IsNull(rsaux10!TIPO_EMPAQUE), "", rsaux10!TIPO_EMPAQUE)
                        list_item.SubItems(9) = IIf(IsNull(rsaux10!caja_pedido), "", rsaux10!caja_pedido)
                        list_item.SubItems(10) = IIf(IsNull(rsaux10!sello), "", rsaux10!sello)
                           
                        rsaux10.MoveNext
                        var_contador_2 = var_contador_2 + 1
                  Wend
                  rsaux10.Close
                  rs.MoveNext
            Wend
            rs.Close
      
         End If
         
      
      
      
   
         'If var_contador_2 < 13 And Me.lbl_semaforo.BackColor = &HC0& Then
         '   If var_contador_2 > 0 Then
         '      Me.wmp4.URL = App.Path + "\Cerrar el lote.mp3"
         '      wmp4.Controls.play
         '   End If
         'End If
         Call ilumina_grid
         lbl_tipo_bulto = "Bultos total del embarque: " + CStr(var_total_leidos) + " de " + CStr(var_total_bultos)
         'If var_contador_3 > 0 Then
            Me.lbl_bultos_2 = "Bultos: " + CStr(var_contador_3)
         'End If
         'var_contador_4 = 0
         'If var_Cadena_pedidos <> "" Then
         '   rs.Open "select AGENTE, nombre_agente, pedido, cliente, orden_pedido, ESTATUS, estatus_pedido from tb_oracle_pedidos_asignados_embarques WHERE EMBARQUE = " + Me.txt_embarque + " AND ((ISNULL(ESTATUS,'') = '')) and pedido not in(" + var_Cadena_pedidos + ") order by orden_pedido", cnn, adOpenDynamic, adLockOptimistic
         '   While Not rs.EOF
         '         rsaux10.Open "select * from tb_oracle_cajas_aduana where embarque = " + Me.txt_embarque + " and pedido = " + CStr(rs!pedido) + " and pedido <>" + CStr(var_pedido), cnn, adOpenDynamic, adLockOptimistic
         '         While Not rsaux10.EOF
         '               var_contador_4 = var_contador_4 + 1
         '               rsaux10.MoveNext
         '         Wend
         '         rsaux10.Close
         '         rs.MoveNext
         '   Wend
         '   rs.Close
      
          '  rs.Open " select * from tb_oracle_pedidos_asignados_embarques where   EMBARQUE = " + Me.txt_embarque + " AND ((ISNULL(ESTATUS,'') = '')) and pedido not in(" + var_Cadena_pedidos + ") and isnull(estatus_pedido,0) = 0", cnn, adOpenDynamic, adLockOptimistic
           ' If Not rs.EOF Then
               'Me.lbl_semaforo_3.BackColor = &HC0&
            'Else
               'Me.lbl_semaforo_3.BackColor = &HC000&
'            End If
'            rs.Close
 '        End If
      
         Me.lbl_bultos_3 = "Bultos restantes: " + CStr(var_total_restante)
      End If
   End If
End Sub





Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Me.cmd_silenciar.SetFocus
   End If
End Sub
