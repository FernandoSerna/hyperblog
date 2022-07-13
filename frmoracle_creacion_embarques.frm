VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_creacion_embarques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creación de embarques"
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
   Begin VB.Frame Frame1 
      Height          =   7350
      Left            =   45
      TabIndex        =   0
      Top             =   -45
      Width           =   11550
      Begin VB.CommandButton cmd_buscar 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   735
         Picture         =   "frmoracle_creacion_embarques.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Buscar pedidos"
         Top             =   1620
         Width           =   330
      End
      Begin VB.Frame Frame2 
         Height          =   45
         Left            =   15
         TabIndex        =   22
         Top             =   1950
         Width           =   11520
      End
      Begin VB.CommandButton com_guardar_equipo 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   75
         Picture         =   "frmoracle_creacion_embarques.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Agregar pedido"
         Top             =   1605
         Width           =   330
      End
      Begin VB.CommandButton com_eliminar_equipo 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   405
         Picture         =   "frmoracle_creacion_embarques.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Eliminar pedido"
         Top             =   1605
         Width           =   330
      End
      Begin VB.TextBox txt_total 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   9540
         TabIndex        =   19
         Top             =   6825
         Width           =   1920
      End
      Begin VB.TextBox txt_piezas 
         Height          =   390
         Left            =   9945
         TabIndex        =   17
         Top             =   2550
         Width           =   1485
      End
      Begin VB.TextBox txt_cliente 
         Height          =   390
         Left            =   1185
         TabIndex        =   15
         Top             =   2565
         Width           =   7860
      End
      Begin VB.TextBox txt_ruta 
         Height          =   390
         Left            =   3570
         TabIndex        =   13
         Top             =   2085
         Width           =   7860
      End
      Begin VB.TextBox txt_pedido 
         Height          =   390
         Left            =   1185
         TabIndex        =   11
         Top             =   2070
         Width           =   1485
      End
      Begin VB.TextBox txt_jaula 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6525
         TabIndex        =   9
         Text            =   "112"
         Top             =   585
         Width           =   1950
      End
      Begin VB.TextBox txt_embarque 
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
         Height          =   600
         Left            =   2655
         TabIndex        =   5
         Text            =   "12313"
         Top             =   525
         Width           =   3105
      End
      Begin VB.Frame Frame4 
         Height          =   60
         Left            =   0
         TabIndex        =   3
         Top             =   3015
         Width           =   11535
      End
      Begin VB.Frame Frame3 
         Height          =   45
         Left            =   30
         TabIndex        =   1
         Top             =   1245
         Width           =   11535
      End
      Begin MSComctlLib.ListView lv_pedidos 
         Height          =   3660
         Left            =   45
         TabIndex        =   6
         Top             =   3105
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   6456
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "   Número"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ruta"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Piezas"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Clave ruta"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Clave cliente"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8775
         TabIndex        =   18
         Top             =   6900
         Width           =   690
      End
      Begin VB.Label Label8 
         Caption         =   "Piezas:"
         Height          =   270
         Left            =   9225
         TabIndex        =   16
         Top             =   2610
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2655
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Left            =   2910
         TabIndex        =   12
         Top             =   2175
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   2175
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Jaula:"
         Height          =   300
         Left            =   6030
         TabIndex        =   8
         Top             =   675
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   1950
         TabIndex        =   7
         Top             =   728
         Width           =   600
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   "  Pedidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   45
         TabIndex        =   4
         Top             =   1305
         Width           =   11475
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   " Embarque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   2
         Top             =   120
         Width           =   11475
      End
   End
End
Attribute VB_Name = "frmoracle_creacion_embarques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Top = 0
  Left = 0
End Sub
