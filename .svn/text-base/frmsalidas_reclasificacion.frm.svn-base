VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmsalidas_reclasificacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm_lista 
      Height          =   2325
      Left            =   2145
      TabIndex        =   3
      Top             =   750
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   4
         Top             =   405
         Width           =   5595
         _ExtentX        =   9869
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
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   5
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame4 
      Height          =   4935
      Left            =   5835
      TabIndex        =   27
      Top             =   2370
      Width           =   5685
      Begin MSComctlLib.ListView lv_destino 
         Height          =   4020
         Left            =   30
         TabIndex        =   28
         Top             =   420
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   7091
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
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl_total_destino 
         Alignment       =   1  'Right Justify
         Caption         =   "999,999,999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3135
         TabIndex        =   31
         Top             =   4470
         Width           =   2445
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   29
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   1110
      Width           =   7440
      Begin VB.TextBox txt_proveedor 
         Height          =   315
         Left            =   945
         TabIndex        =   18
         Top             =   810
         Width           =   1140
      End
      Begin VB.TextBox txt_nombre_proveedor 
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   825
         Width           =   5250
      End
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   450
         Width           =   5250
      End
      Begin VB.TextBox txt_almacen 
         Height          =   315
         Left            =   945
         TabIndex        =   15
         Top             =   450
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   21
         Top             =   885
         Width           =   585
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   20
         Top             =   120
         Width           =   7350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Origen:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   19
         Top             =   510
         Width           =   510
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11175
      Picture         =   "frmsalidas_reclasificacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Salir"
      Top             =   705
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   735
      Picture         =   "frmsalidas_reclasificacion.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmsalidas_reclasificacion.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Buscar Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmsalidas_reclasificacion.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   75
      TabIndex        =   9
      Top             =   570
      Width           =   11520
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Index           =   0
      Left            =   7590
      TabIndex        =   6
      Top             =   1110
      Width           =   3930
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
         TabIndex        =   7
         Top             =   540
         Width           =   3690
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   8
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   390
      TabIndex        =   0
      Top             =   1095
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         TabIndex        =   1
         Top             =   495
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   30
         TabIndex        =   2
         Top             =   120
         Width           =   3060
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   570
      Top             =   825
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
            Picture         =   "frmsalidas_reclasificacion.frx":0940
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reclasificacion.frx":121A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reclasificacion.frx":1AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reclasificacion.frx":2090
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reclasificacion.frx":296C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reclasificacion.frx":3246
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reclasificacion.frx":3B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reclasificacion.frx":3C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reclasificacion.frx":3D44
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reclasificacion.frx":3E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reclasificacion.frx":3F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reclasificacion.frx":407A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   615
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
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   75
      TabIndex        =   22
      Top             =   975
      Width           =   11520
   End
   Begin VB.Frame Frame2 
      Height          =   4935
      Left            =   120
      TabIndex        =   24
      Top             =   2370
      Width           =   5685
      Begin MSComctlLib.ListView lv_origen 
         Height          =   4035
         Left            =   30
         TabIndex        =   25
         Top             =   405
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   7117
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
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl_tota_origen 
         Alignment       =   1  'Right Justify
         Caption         =   "999,999,999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3105
         TabIndex        =   30
         Top             =   4515
         Width           =   2445
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   26
         Top             =   120
         Width           =   5610
      End
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
      TabIndex        =   23
      Top             =   75
      Width           =   11430
   End
End
Attribute VB_Name = "frmsalidas_reclasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   var_posible_kanban = 0
   Me.frm_lista.Visible = False
End Sub

