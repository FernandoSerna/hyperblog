VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_configuracion_articulos_devolucion_ANES 
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5280
      Picture         =   "frmoracle_configuracion_articulos_devolucion_ANES.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   3480
      Width           =   330
   End
   Begin VB.ComboBox cmb_filtro 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   5055
   End
   Begin MSComctlLib.ListView lv_articulos 
      Height          =   6000
      Left            =   5640
      TabIndex        =   0
      Top             =   1080
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   10583
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
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Marca"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lv_filtro 
      Height          =   6600
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   11642
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
         Text            =   "Filtro"
         Object.Width           =   9022
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Marca"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmoracle_configuracion_articulos_devolucion_ANES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
