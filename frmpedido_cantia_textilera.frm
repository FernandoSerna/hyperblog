VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmpedido_cantia_textilera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos Cantia - Textilera"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   2085
      TabIndex        =   38
      Top             =   4140
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   39
         Top             =   480
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
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7057
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   40
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_busqueda 
      Height          =   915
      Left            =   1125
      TabIndex        =   33
      Top             =   270
      Width           =   2760
      Begin VB.TextBox txt_busqueda 
         Height          =   360
         Left            =   135
         TabIndex        =   34
         Top             =   435
         Width           =   2505
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         Caption         =   " Número de Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   30
         TabIndex        =   35
         Top             =   120
         Width           =   2685
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   " Pedido "
      Height          =   1485
      Left            =   8715
      TabIndex        =   32
      Top             =   435
      Width           =   2100
      Begin VB.TextBox txt_numero_pedido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   60
         TabIndex        =   24
         Top             =   570
         Width           =   1950
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   750
      Picture         =   "frmpedido_cantia_textilera.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Movimiento Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10440
      Picture         =   "frmpedido_cantia_textilera.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmpedido_cantia_textilera.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Buscar Pedido Alt + B"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmpedido_cantia_textilera.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Pedido Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   90
      Left            =   60
      TabIndex        =   31
      Top             =   300
      Width           =   10785
   End
   Begin VB.Frame Frame3 
      Caption         =   " Artículo "
      Height          =   1845
      Left            =   90
      TabIndex        =   28
      Top             =   1965
      Width           =   10725
      Begin VB.CommandButton cmd_guardar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10245
         Picture         =   "frmpedido_cantia_textilera.frx":0940
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Guardar"
         Top             =   1425
         Width           =   330
      End
      Begin VB.TextBox txt_nota_2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1050
         TabIndex        =   12
         Top             =   1005
         Width           =   9555
      End
      Begin VB.TextBox txt_alto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6225
         TabIndex        =   16
         Top             =   1395
         Width           =   1365
      End
      Begin VB.TextBox txt_largo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4080
         TabIndex        =   15
         Top             =   1410
         Width           =   1350
      End
      Begin VB.TextBox txt_ancho 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1785
         TabIndex        =   14
         Top             =   1395
         Width           =   1410
      End
      Begin VB.TextBox txt_nota_1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1050
         TabIndex        =   11
         Top             =   615
         Width           =   9555
      End
      Begin VB.CheckBox chk_blackout 
         Caption         =   "Blackout"
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1455
         Width           =   975
      End
      Begin VB.TextBox txt_precio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8505
         TabIndex        =   17
         Top             =   1395
         Width           =   1215
      End
      Begin VB.TextBox txt_cantidad 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9180
         TabIndex        =   10
         Top             =   225
         Width           =   1410
      End
      Begin VB.TextBox txt_descripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2595
         TabIndex        =   9
         Top             =   225
         Width           =   5625
      End
      Begin VB.TextBox txt_codigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1050
         TabIndex        =   8
         Top             =   225
         Width           =   1515
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Alto:"
         Height          =   195
         Left            =   5775
         TabIndex        =   46
         Top             =   1485
         Width           =   315
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Largo:"
         Height          =   195
         Left            =   3585
         TabIndex        =   45
         Top             =   1485
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Ancho:"
         Height          =   195
         Left            =   1215
         TabIndex        =   44
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   135
         TabIndex        =   43
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Notas:"
         Height          =   195
         Left            =   135
         TabIndex        =   42
         Top             =   705
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Precio:"
         Height          =   195
         Left            =   7980
         TabIndex        =   30
         Top             =   1485
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Index           =   0
         Left            =   8490
         TabIndex        =   29
         Top             =   308
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Pedido"
      Height          =   1500
      Left            =   90
      TabIndex        =   27
      Top             =   435
      Width           =   8580
      Begin VB.TextBox txt_notas 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   765
         TabIndex        =   7
         Top             =   1035
         Width           =   7770
      End
      Begin VB.TextBox txt_anticipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   765
         TabIndex        =   6
         Top             =   645
         Width           =   1335
      End
      Begin VB.TextBox txt_nombre_cliente 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2115
         TabIndex        =   5
         Top             =   255
         Width           =   6390
      End
      Begin VB.TextBox txt_cliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   765
         TabIndex        =   4
         Top             =   255
         Width           =   1335
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Notas:"
         Height          =   195
         Left            =   135
         TabIndex        =   48
         Top             =   1118
         Width           =   465
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   135
         TabIndex        =   47
         Top             =   338
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Anticipo:"
         Height          =   195
         Left            =   135
         TabIndex        =   41
         Top             =   728
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   90
      TabIndex        =   25
      Top             =   3780
      Width           =   10725
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   6435
         TabIndex        =   49
         Top             =   810
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   75
            TabIndex        =   50
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
            TabIndex        =   51
            Top             =   15
            Width           =   2895
         End
      End
      Begin VB.TextBox txt_subimporte 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   8250
         TabIndex        =   20
         Top             =   2265
         Width           =   2355
      End
      Begin VB.TextBox txt_importe_descuento 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   8895
         TabIndex        =   22
         Top             =   2655
         Width           =   1710
      End
      Begin VB.TextBox txt_descuento 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8265
         TabIndex        =   21
         Top             =   2655
         Width           =   615
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   8265
         TabIndex        =   23
         Top             =   3045
         Width           =   2355
      End
      Begin MSComctlLib.ListView lv_pedidos 
         Height          =   2070
         Left            =   60
         TabIndex        =   19
         Top             =   210
         Width           =   10590
         _ExtentX        =   18680
         _ExtentY        =   3651
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Precio"
            Object.Width           =   2294
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Importe"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "NOTA 1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "NOTA 2"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Ancho"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Largo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Alto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "blackout"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Subimporte:"
         Height          =   195
         Left            =   7350
         TabIndex        =   37
         Top             =   2355
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descuento:"
         Height          =   195
         Left            =   7365
         TabIndex        =   36
         Top             =   2745
         Width           =   825
      End
      Begin VB.Label Label1 
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
         Left            =   7500
         TabIndex        =   26
         Top             =   3105
         Width           =   690
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   0
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
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "frmpedido_cantia_textilera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_renglon As Integer
Dim var_primera_vez As Integer
Dim var_estatus As String

Sub ilumina_grid()
   var_n = lv_pedidos.ListItems.Count
   For var_i = 1 To var_n
       If var_i = var_renglon Then
          lv_pedidos.ListItems.Item(var_i).Bold = True
          lv_pedidos.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_pedidos.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_pedidos.ListItems.Item(var_i).ListSubItems(3).Bold = True
          lv_pedidos.ListItems.Item(var_i).ListSubItems(4).Bold = True
          lv_pedidos.ListItems.Item(var_i).ListSubItems(5).Bold = True
          lv_pedidos.ListItems.Item(var_i).ListSubItems(6).Bold = True
          lv_pedidos.ListItems.Item(var_i).ListSubItems(7).Bold = True
          lv_pedidos.ListItems.Item(var_i).ListSubItems(8).Bold = True
          lv_pedidos.ListItems.Item(var_i).ForeColor = &H8000&
          lv_pedidos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_pedidos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
          lv_pedidos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H8000&
          lv_pedidos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H8000&
          lv_pedidos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H8000&
          lv_pedidos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H8000&
       Else
          lv_pedidos.ListItems.Item(var_i).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(1).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(2).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(3).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(4).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(5).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(6).Bold = False
          lv_pedidos.ListItems.Item(var_i).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H80000012
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_pedidos.ListItems.Item(var_renglon).Selected = True
      lv_pedidos.selectedItem.EnsureVisible
   End If
   lv_pedidos.Refresh
End Sub


Private Sub cmd_buscar_Click()
   Me.frm_busqueda.Visible = True
   Me.txt_busqueda = ""
   Me.txt_busqueda.SetFocus
End Sub

Private Sub cmd_guardar_Click()
      If Me.txt_cliente <> "" Then
         If Me.txt_codigo <> "" Then
            If IsNumeric(Me.txt_precio) Then
               If IsNumeric(Me.txt_cantidad) Then
                  If Trim(Me.txt_anticipo) = "" Then
                     Me.txt_anticipo = 0
                  End If
                  If IsNumeric(Me.txt_anticipo) Then
                     If Trim(Me.txt_alto) = "" Then
                        Me.txt_alto = 0
                     End If
                     If IsNumeric(Me.txt_alto) Then
                        If Trim(Me.txt_ancho) = "" Then
                           Me.txt_ancho = 0
                        End If
                        If IsNumeric(Me.txt_ancho) Then
                           If Trim(Me.txt_largo) = "" Then
                              Me.txt_largo = 0
                           End If
                           If IsNumeric(Me.txt_largo) Then
                              If Trim(Me.txt_nota_1) <> "" Then
                                 If var_primera_vez = 1 Then
                                    cnn.BeginTrans
                                    rs.Open "SELECT MAX(INTE_PED_NUMERO) FROM TB_ENCABEZADO_PEDIDOS", cnn_pedido_cantia_textilera, adOpenDynamic, adLockOptimistic
                                    var_numero_pedido = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                                    Me.txt_numero_pedido = CStr(var_numero_pedido)
                                    var_cadena = "INSERT INTO TB_ENCABEZADO_PEDIDOS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, CHAR_TPE_TIPO_PEDIDO_ID, INTE_PED_NUMERO, INTE_PED_REFERENCIA, DTIM_PED_FECHA, DTIM_PED_REFERENCIA, VCHA_AGE_AGENTE_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_PED_RESURTIBLE, INTE_PED_ESPECIALES, CHAR_PED_ESTATUS, FLOA_PED_DESCUENTO_1, FLOA_PED_DESCUENTO_2, FLOA_PED_DESCUENTO_3, INTE_PED_DIAS_CONDICIONES, INTE_PED_DIAS_CADUCIDAD, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, DTIM_AUD_FECHA, VCHA_MON_MONEDA_ID, INTE_PED_SUGERIDO, VCHA_PED_CLIENTE_CANTIA, INTE_PED_PEDIDO_CANTIA, FLOA_PED_ANTICIPO, VCHA_PED_NOTAS)"
                                    var_cadena = var_cadena + " VALUES ('18','16', 'PTTEX','M'," + CStr(var_numero_pedido) + ",0,GETDATE(),GETDATE(),'00100','T000001052','C000003687','E000003228',0,0,''," + Me.txt_descuento + ",0,0,60,0,'" + var_clave_usuario_global + "','" + fun_NombrePc + "',GETDATE(),'1',0,'" + Me.txt_cliente + "',1," + txt_anticipo + ", '" + Me.txt_notas + "')"
                                    If rsaux.State = 1 Then
                                       rsaux.Close
                                    End If
                                    rsaux.Open var_cadena, cnn_pedido_cantia_textilera, adOpenDynamic, adLockOptimistic
                                    var_primera_vez = 0
                                    rs.Close
                                    cnn.CommitTrans
                                    Me.txt_cliente.Enabled = False
                                    Me.txt_nombre_cliente.Enabled = False
                                    Me.txt_anticipo.Enabled = False
                                    Me.txt_notas.Enabled = False
                                 End If
                                 rs.Open "SELECT * FROM TB_DETALLE_PEDIDOS WHERE INTE_PED_NUMERO = " + Me.txt_numero_pedido + " AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn_pedido_cantia_textilera, adOpenDynamic, adLockOptimistic
                                 If Not rs.EOF Then
                                    
                                    rsaux.Open "UPDATE TB_dETALLE_PEDIDOS SET FLOA_PED_CANTIDAD = " + CStr(CDbl(Me.txt_cantidad)) + ", FLOA_PED_PRECIO = " + CStr(CDbl(Me.txt_precio) / 1.16) + ", INTE_PED_BLACKOUT = " + CStr(Me.chk_blackout.Value) + ", vcha_ped_nota_1 = '" + Me.txt_nota_1 + "', vcha_ped_nota_2 = '" + Me.txt_nota_2 + "', FLOA_PED_ANCHO = " + Me.txt_ancho + ", FLOA_PED_LARGO = " + Me.txt_largo + ", FLOA_PED_ALTO = " + Me.txt_alto + " WHERE  INTE_PED_NUMERO = " + CStr(Me.txt_numero_pedido) + " AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn_pedido_cantia_textilera, adOpenDynamic, adLockOptimistic
                                    var_n = lv_pedidos.ListItems.Count
                                    var_i = 1
                                    While (var_i <= var_n)
                                        lv_pedidos.ListItems.Item(var_i).Selected = True
                                        valor = Trim(lv_pedidos.selectedItem)
                                        If txt_codigo = valor Then
                                           var_encontro = 1
                                           var_i = var_n
                                        End If
                                        var_i = var_i + 1
                                    Wend
                                    bandera_suma = True
                                    Me.lv_pedidos.selectedItem.SubItems(2) = Format(CDbl(Me.txt_cantidad), "###,###,##0.00")
                                    Me.lv_pedidos.selectedItem.SubItems(3) = Format(CDbl(Me.txt_precio), "###,###,##0.00")
                                    Me.lv_pedidos.selectedItem.SubItems(4) = Format(CDbl(Me.lv_pedidos.selectedItem.SubItems(2)) * CDbl(Me.txt_precio), "###,###,##0.00")
                                    Me.lv_pedidos.selectedItem.SubItems(5) = Me.txt_nota_1
                                    Me.lv_pedidos.selectedItem.SubItems(6) = Me.txt_nota_2
                                    Me.lv_pedidos.selectedItem.SubItems(7) = Me.txt_ancho
                                    Me.lv_pedidos.selectedItem.SubItems(8) = Me.txt_largo
                                    Me.lv_pedidos.selectedItem.SubItems(9) = Me.txt_alto
                                    Me.lv_pedidos.selectedItem.SubItems(10) = Me.chk_blackout.Value
                                    var_renglon = lv_pedidos.selectedItem.Index
                                    Call ilumina_grid
                                 Else
                                    var_cadena = "insert into tb_detalle_pedidos (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, inte_ped_numero, vcha_Art_articulo_id, floa_ped_precio, floa_ped_cantidad, floa_ped_cantidad_surtida, floa_ped_Cantidad_depurada, floa_ped_promocion_1, floa_ped_promocion_2, floa_ped_negado_produccion, floa_ped_negado_ditribucion, floa_ped_negado_autorizacion, char_ped_tipo, INTE_PED_BLACKOUT,                    vcha_ped_nota_1,         vcha_ped_nota_2,  FLOA_PED_ANCHO, FLOA_PED_LARGO, FLOA_PED_ALTO)"
                                    var_cadena = var_cadena + "  values (                         '18',               '16',             'PTTEX', " + CStr(Me.txt_numero_pedido) + ",'" + Me.txt_codigo + "'," + CStr(CDbl(Me.txt_precio) / 1.16) + "," + Me.txt_cantidad + ",  0,                          0,                    0,                    0,                          0,0,                          0,             ''," + CStr(Me.chk_blackout.Value) + ",'" + Me.txt_nota_1 + "', '" + Me.txt_nota_2 + "', " + Me.txt_ancho + ", " + Me.txt_largo + "," + Me.txt_alto + ")  "
                                    
                                    rsaux.Open var_cadena, cnn_pedido_cantia_textilera, adOpenDynamic, adLockOptimistic
                                    Set list_item = lv_pedidos.ListItems.Add(, , txt_codigo)
                                    list_item.SubItems(1) = Me.txt_descripcion
                                    list_item.SubItems(2) = Format(Me.txt_cantidad, "###,###,##0.00")
                                    list_item.SubItems(3) = Format(Me.txt_precio, "###,###,##0.00")
                                    list_item.SubItems(4) = Format(CDbl(Me.txt_cantidad) * CDbl(Me.txt_precio), "###,###,##0.00")
                                    list_item.SubItems(5) = Me.txt_nota_1
                                    list_item.SubItems(6) = Me.txt_nota_2
                                    list_item.SubItems(7) = Me.txt_ancho
                                    list_item.SubItems(8) = Me.txt_largo
                                    list_item.SubItems(9) = Me.txt_alto
                                    list_item.SubItems(10) = Me.chk_blackout.Value
                                    var_renglon = lv_pedidos.ListItems.Count
                                    Call ilumina_grid
                                 End If
                                 rs.Close
                                 var_importe = 0
                                 For var_i = 1 To lv_pedidos.ListItems.Count
                                     lv_pedidos.ListItems.Item(var_i).Selected = True
                                     var_importe = var_importe + CDbl(Me.lv_pedidos.selectedItem.SubItems(4))
                                 Next var_i
                                 Me.txt_subimporte = Format(var_importe, "###,###,##0.00")
                                 Me.txt_importe_descuento = Format(CDbl(Me.txt_subimporte) - (CDbl(Me.txt_subimporte) * (1 - (CDbl(Me.txt_descuento) / 100))), "###,###,##0.00")
                                 Me.txt_importe = Format(CDbl(Me.txt_subimporte) - CDbl(Me.txt_importe_descuento), "###,###,##0.00")
                                 Me.txt_codigo.SetFocus
                              Else
                                 MsgBox "No se a indicado la tela", vbOKOnly, "ATENCION"
                              End If
                           Else
                              MsgBox "Medida de largo incorrecta", vbOKOnly, "ATENCION"
                           End If
                        Else
                           MsgBox "Medida de ancho incorrecta", vbOKOnly, "ATENCION"
                        End If
                     Else
                        MsgBox "Medida de alto incorrecto", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "Anticipo incorrecto", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Precio incorrecto", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Código incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Cliente incorrecto", vbOKOnly, "ATENCION"
      End If
End Sub

Private Sub cmd_imprimir_Click()
   Dim dl As Long                                 ' Valor devuelto por la función API
   Dim sAttributes As String                  ' Aributos
   Dim sDriver As String                       ' Nombre del controlador
   Dim sDescription As String                ' Descripción del DSN
   Dim sDsnName As String                  ' Nombre del DSN
   Const ODBC_ADD_SYS_DSN As Long = 4         ' Se creará un DSN de sistema
   Const vbAPINull As Long = 0&                         ' Puntero NULL
   ' se elimina
   Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema
   Me.frm_busqueda.Visible = False
   If Trim(Me.txt_numero_pedido) <> "" Then
      If Trim(var_estatus) = "" Then
         var_si = MsgBox("¿Se va a cerrar el pedido?", vbOKCancel, "ATENCION")
         If var_si = 1 Then
            var_estatus = "I"
            rsaux2.Open "update tb_encabezado_pedidos set INTE_PED_AUTORIZO = 0, char_ped_estatus = 'I' where inte_ped_numero = " + txt_numero_pedido, cnn_pedido_cantia_textilera, adOpenDynamic, adLockOptimistic
            
            sDsnName = "DSN=sqlsistema"
            sDriver = "SQL Server"
            dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)

            'se crea
            sDsnName = "sqlsistema"
            sDescription = "sqlsistema"
            sDriver = "SQL Server"
            sAttributes = "DSN=" & sDsnName & Chr(0)
            sAttributes = sAttributes & "Server=" + "sqlquezada2" & Chr$(0)
            sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
            sAttributes = sAttributes & "Database=" + "sidtextilera" & Chr(0)
            strAttributes = strAttributes & "UID=sa" & Chr$(0)
            strAttributes = strAttributes & "PWD=elia" & Chr$(0)
            dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
            
            
            Set reporte = appl.OpenReport(App.Path + "\REP_PEDIDO_cANTIA_TEXTILERA_2.rpt")
            reporte.RecordSelectionFormula = "{VW_PEDIDOS_CANTIA_TEXITLERA.INTE_PED_NUMERO} = " + Me.txt_numero_pedido
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Pedidos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
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
            sAttributes = sAttributes & "Database=" + var_bd_reportes & Chr(0)
            strAttributes = strAttributes & "UID=sa" & Chr$(0)
            strAttributes = strAttributes & "PWD=elia" & Chr$(0)
            dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
            txt_codigo.Enabled = False
            txt_cantidad.Enabled = False
            txt_descripcion.Enabled = False
            Me.txt_precio.Enabled = False
            Me.txt_nota_1.Enabled = False
            Me.txt_nota_2.Enabled = False
            Me.txt_ancho.Enabled = False
            Me.txt_alto.Enabled = False
            Me.txt_largo.Enabled = False
            Me.chk_blackout.Enabled = False
            Me.txt_anticipo.Enabled = False
            Me.txt_notas.Enabled = False
         
            If MAPISession1.SessionID = 0 Then
               MAPISession1.SignOn
            End If
            MAPIMessages1.SessionID = MAPISession1.SessionID
            MAPIMessages1.Compose
            MAPIMessages1.RecipDisplayName = "pablo.medina@textilera.com.mx"
            MAPIMessages1.RecipAddress = "pablo.medina@textilera.com.mx"
            MAPIMessages1.AddressResolveUI = True
            MAPIMessages1.ResolveName
            MAPIMessages1.MsgSubject = "Pedido CANTIA - Textilera número " + Me.txt_numero_pedido
            MAPIMessages1.MsgNoteText = "Se a generado el pedido CANTIA - Textilera número " + Me.txt_numero_pedido
            MAPIMessages1.Send True
            If MAPISession1.SessionID > 0 Then
               MAPISession1.SignOff
            End If
         
         
         
         
         End If
      Else
         sDsnName = "DSN=sqlsistema"
         sDriver = "SQL Server"
         dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)
        'se crea
         sDsnName = "sqlsistema"
         sDescription = "sqlsistema"
         sDriver = "SQL Server"
         sAttributes = "DSN=" & sDsnName & Chr(0)
         sAttributes = sAttributes & "Server=" + "sqlquezada2" & Chr$(0)
         sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
         sAttributes = sAttributes & "Database=" + "sidtextilera" & Chr(0)
         strAttributes = strAttributes & "UID=sa" & Chr$(0)
         strAttributes = strAttributes & "PWD=elia" & Chr$(0)
         dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
         Set reporte = appl.OpenReport(App.Path + "\REP_PEDIDO_cANTIA_TEXTILERA_2.rpt")
         reporte.RecordSelectionFormula = "{VW_PEDIDOS_CANTIA_TEXITLERA.INTE_PED_NUMERO} = " + txt_numero_pedido
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Pedidos"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         
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
         sAttributes = sAttributes & "Database=" + var_bd_reportes & Chr(0)
         strAttributes = strAttributes & "UID=sa" & Chr$(0)
         strAttributes = strAttributes & "PWD=elia" & Chr$(0)
         dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
         
         txt_codigo.Enabled = False
         txt_cantidad.Enabled = False
         txt_descripcion.Enabled = False
         Me.txt_precio.Enabled = False
         Me.txt_nota_1.Enabled = False
         Me.txt_nota_2.Enabled = False
         Me.txt_ancho.Enabled = False
         Me.txt_alto.Enabled = False
         Me.txt_largo.Enabled = False
         Me.chk_blackout.Enabled = False
         Me.txt_anticipo.Enabled = False
         Me.txt_notas.Enabled = False
     
     
         If MAPISession1.SessionID = 0 Then
            MAPISession1.SignOn
         End If
         MAPIMessages1.SessionID = MAPISession1.SessionID
         MAPIMessages1.Compose
         MAPIMessages1.RecipDisplayName = "pablo.medina@textilera.com.mx"
         MAPIMessages1.RecipAddress = "pablo.medina@textilera.com.mx"
         MAPIMessages1.AddressResolveUI = True
         MAPIMessages1.ResolveName
         MAPIMessages1.MsgSubject = "Pedido CANTIA - Textilera número " + Me.txt_numero_pedido
         MAPIMessages1.MsgNoteText = "Se a generado el pedido CANTIA - Textilera número " + Me.txt_numero_pedido
         MAPIMessages1.Send True
         If MAPISession1.SessionID > 0 Then
            MAPISession1.SignOff
         End If
     
     
     
     
      End If
   End If
End Sub

Private Sub cmd_nuevo_Click()
   rs.Open "select * from vw_clientes where vcha_cli_clave_id = 'C000003687'", cnn_pedido_cantia_textilera, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Me.txt_descuento = IIf(IsNull(rs!floa_gac_Descuento_1), 0, rs!floa_gac_Descuento_1)
   End If
   rs.Close
   var_primera_vez = 1
   var_estatus = ""
   Me.txt_cliente = ""
   Me.txt_nombre_cliente = ""
   Me.txt_codigo = ""
   Me.txt_descripcion = ""
   Me.txt_cantidad = ""
   Me.txt_precio = ""
   Me.txt_numero_pedido = ""
   Me.txt_importe_descuento = ""
   Me.txt_subimporte = ""
   Me.lv_pedidos.ListItems.Clear
   Me.txt_anticipo = ""
   Me.txt_notas = ""
   Me.txt_nota_1 = ""
   Me.txt_nota_2 = ""
   Me.txt_alto = ""
   Me.txt_ancho = ""
   Me.txt_largo = ""
   Me.chk_blackout.Value = 0
   Me.txt_cliente.Enabled = True
   Me.txt_nombre_cliente.Enabled = True
   Me.txt_codigo.Enabled = True
   Me.txt_descripcion.Enabled = True
   Me.txt_precio.Enabled = True
   Me.txt_cantidad.Enabled = True
   Me.txt_anticipo.Enabled = True
   Me.txt_notas.Enabled = True
   Me.txt_nota_1.Enabled = True
   Me.txt_nota_2.Enabled = True
   Me.txt_alto.Enabled = True
   Me.txt_ancho.Enabled = True
   Me.txt_largo.Enabled = True
   Me.chk_blackout.Enabled = True
   
   Me.txt_cliente.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.frm_lista.Visible = False
   frm_busqueda.Visible = False
   Me.frm_eliminar.Visible = False
   var_primera_vez = 1
   Top = 0
   Left = 350
   rs.Open "select * from vw_clientes where vcha_cli_clave_id = 'C000003687'", cnn_pedido_cantia_textilera, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Me.txt_descuento = IIf(IsNull(rs!floa_gac_Descuento_1), 0, rs!floa_gac_Descuento_1)
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_lista.ListItems.Count > 0 Then
         Me.txt_cliente = lv_lista.selectedItem
         Me.txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
         Me.txt_cliente.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.txt_cliente.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub lv_pedidos_GotFocus()
      Me.txt_codigo = Me.lv_pedidos.selectedItem
      Me.txt_descripcion = Me.lv_pedidos.selectedItem.SubItems(1)
      Me.txt_cantidad = Format(Me.lv_pedidos.selectedItem.SubItems(2), "###,###,##0.00")
      Me.txt_precio = Format(Me.lv_pedidos.selectedItem.SubItems(3), "###,###,##0.00")
      'Me.lv_pedidos.selectedItem.SubItems(4) = Format(CDbl(Me.lv_pedidos.selectedItem.SubItems(2)) * CDbl(Me.txt_precio), "###,###,##0.00")
      Me.txt_nota_1 = Me.lv_pedidos.selectedItem.SubItems(5)
      Me.txt_nota_2 = Me.lv_pedidos.selectedItem.SubItems(6)
      Me.txt_ancho = Me.lv_pedidos.selectedItem.SubItems(7)
      Me.txt_largo = Me.lv_pedidos.selectedItem.SubItems(8)
      Me.txt_alto = Me.lv_pedidos.selectedItem.SubItems(9)
      Me.chk_blackout.Value = Me.lv_pedidos.selectedItem.SubItems(10)
End Sub

Private Sub lv_pedidos_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo salir:
   If Me.lv_pedidos.ListItems.Count > 0 Then
      Me.txt_codigo = Me.lv_pedidos.selectedItem
      Me.txt_descripcion = Me.lv_pedidos.selectedItem.SubItems(1)
      Me.txt_cantidad = Format(Me.lv_pedidos.selectedItem.SubItems(2), "###,###,##0.00")
      Me.txt_precio = Format(Me.lv_pedidos.selectedItem.SubItems(3), "###,###,##0.00")
      Me.txt_nota_1 = Me.lv_pedidos.selectedItem.SubItems(5)
      Me.txt_nota_2 = Me.lv_pedidos.selectedItem.SubItems(6)
      Me.txt_ancho = Me.lv_pedidos.selectedItem.SubItems(7)
      Me.txt_largo = Me.lv_pedidos.selectedItem.SubItems(8)
      Me.txt_alto = Me.lv_pedidos.selectedItem.SubItems(9)
      Me.chk_blackout = Me.lv_pedidos.selectedItem.SubItems(10)
   End If
   Exit Sub
salir:
   Exit Sub
End Sub

Private Sub lv_pedidos_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If Me.txt_precio.Enabled = True Then
         Me.frm_eliminar.Visible = True
         Me.txt_cantidad_eliminar = ""
         Me.txt_cantidad_eliminar.SetFocus
      End If
   End If
End Sub

Private Sub txt_alto_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_alto_LostFocus()
   If Not IsNumeric(Me.txt_alto) Then
      MsgBox "Medida de alto incorrecta", vbOKOnly, "ATENCION"
      Me.txt_alto = ""
   End If
End Sub

Private Sub txt_ancho_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ancho_LostFocus()
   If Not IsNumeric(Me.txt_ancho) Then
      MsgBox "Medida de ancho incorrecta", vbOKOnly, "ATENCION"
      Me.txt_ancho = ""
   End If
End Sub

Private Sub txt_anticipo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_anticipo_LostFocus()
   If Not IsNumeric(Me.txt_anticipo) Then
      MsgBox "Importe de anticipo incorrecto", vbOKOnly, "ATENCION"
      Me.txt_anticipo = ""
   End If
End Sub

Private Sub txt_busqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      If IsNumeric(Me.txt_busqueda) Then
         var_cadena = "SELECT INTE_PED_BLACKOUT, dbo.TB_DETALLE_PEDIDOS.vcha_ped_nota_1, vcha_ped_nota_2,  FLOA_PED_ANCHO, FLOA_PED_LARGO, FLOA_PED_ALTO, FLOA_PED_ANTICIPO, VCHA_PED_NOTAS, TB_ENCABEZADO_PEDIDOS.floa_ped_descuento_1, isnull(dbo.TB_ENCABEZADO_PEDIDOS.char_ped_estatus,'') as char_ped_Esatus, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_CLIENTE_CANTIA, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_PEDIDO_CANTIA, dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_DETALLE_PEDIDOS.FLOA_PED_PRECIO, dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD FROM dbo.TB_ENCABEZADO_PEDIDOS INNER JOIN dbo.TB_DETALLE_PEDIDOS ON dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO = dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_DETALLE_PEDIDOS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID Where (dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_PEDIDO_CANTIA = 1) And (dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO = " + Me.txt_busqueda + ")"
         rs.Open var_cadena, cnn_pedido_cantia_textilera, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + IIf(IsNull(rs!VCHA_PED_CLIENTE_CANTIA), "", rs!VCHA_PED_CLIENTE_CANTIA) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_cliente = IIf(IsNull(rsaux!vcha_cli_clave_id), "", rsaux!vcha_cli_clave_id)
               Me.txt_nombre_cliente = IIf(IsNull(rsaux!VCHA_CLI_NOMBRE), "", rsaux!VCHA_CLI_NOMBRE)
            Else
               Me.txt_cliente = ""
               Me.txt_nombre_cliente = ""
            End If
            rsaux.Close
            Me.txt_anticipo = IIf(IsNull(rs!floa_ped_anticipo), 0, rs!floa_ped_anticipo)
            Me.txt_notas = IIf(IsNull(rs!VCHA_PED_NOTAS), "", rs!VCHA_PED_NOTAS)
            var_importe = 0
            Me.txt_numero_pedido = Me.txt_busqueda
            Me.txt_descuento = IIf(IsNull(rs!floa_ped_descuento_1), 0, rs!floa_ped_descuento_1)
            Me.lv_pedidos.ListItems.Clear
            var_estatus = Trim(IIf(IsNull(rs!char_ped_Esatus), "", rs!char_ped_Esatus))
            var_primera_vez = 0
            Me.chk_blackout.Value = IIf(IsNull(rs!INTE_PED_BLACKOUT), 0, rs!INTE_PED_BLACKOUT)
            Me.txt_cliente.Enabled = False
            Me.txt_nombre_cliente.Enabled = False
            Me.txt_anticipo.Enabled = False
            Me.txt_notas.Enabled = False
            If var_estatus <> "" Then
               Me.txt_codigo.Enabled = False
               Me.txt_descripcion.Enabled = False
               Me.txt_precio.Enabled = False
               Me.txt_cantidad.Enabled = False
               Me.txt_nota_1.Enabled = False
               Me.txt_nota_2.Enabled = False
               Me.txt_ancho.Enabled = False
               Me.txt_alto.Enabled = False
               Me.txt_largo.Enabled = False
               Me.chk_blackout.Enabled = False
            Else
               Me.txt_codigo.Enabled = True
               Me.txt_descripcion.Enabled = True
               Me.txt_precio.Enabled = True
               Me.txt_cantidad.Enabled = True
               Me.txt_nota_1.Enabled = True
               Me.txt_nota_2.Enabled = True
               Me.txt_ancho.Enabled = True
               Me.txt_alto.Enabled = True
               Me.txt_largo.Enabled = True
               Me.chk_blackout.Enabled = True
            End If
            While Not rs.EOF
                  Set list_item = lv_pedidos.ListItems.Add(, , rs!vcha_Art_Articulo_id)
                  list_item.SubItems(1) = rs!vcha_Art_nombre_español
                  list_item.SubItems(2) = Format(rs!FLOA_PED_CANTIDAD, "###,###,##0.00")
                  list_item.SubItems(3) = Format(rs!FLOA_PED_PRECIO * 1.16, "###,###,##0.00")
                  list_item.SubItems(4) = Format(CDbl(rs!FLOA_PED_CANTIDAD) * CDbl(rs!FLOA_PED_PRECIO) * 1.16, "###,###,##0.00")
                  list_item.SubItems(5) = IIf(IsNull(rs!vcha_ped_nota_1), "", rs!vcha_ped_nota_1)
                  list_item.SubItems(6) = IIf(IsNull(rs!vcha_ped_nota_2), "", rs!vcha_ped_nota_2)
                  list_item.SubItems(7) = IIf(IsNull(rs!floa_ped_ancho), 0, rs!floa_ped_ancho)
                  list_item.SubItems(8) = IIf(IsNull(rs!floa_ped_largo), 0, rs!floa_ped_largo)
                  list_item.SubItems(9) = IIf(IsNull(rs!floa_ped_alto), 0, rs!floa_ped_alto)
                  list_item.SubItems(10) = IIf(IsNull(rs!INTE_PED_BLACKOUT), 0, rs!INTE_PED_BLACKOUT)
                  var_importe = var_importe + (CDbl(rs!FLOA_PED_CANTIDAD) * CDbl(rs!FLOA_PED_PRECIO) * 1.16)
                  rs.MoveNext
            Wend
            Me.txt_subimporte = Format(var_importe, "###,###,##0.00")
            Me.txt_importe_descuento = Format(CDbl(Me.txt_subimporte) - (CDbl(Me.txt_subimporte) * (1 - (CDbl(Me.txt_descuento) / 100))), "###,###,##0.00")
            Me.txt_importe = Format(CDbl(Me.txt_subimporte) - CDbl(Me.txt_importe_descuento), "###,###,##0.00")
            If Me.txt_codigo.Enabled = True Then
               Me.txt_codigo.SetFocus
            Else
               If Me.lv_pedidos.ListItems.Count > 0 Then
                  Me.lv_pedidos.SetFocus
               Else
                  Me.frm_busqueda.Visible = False
               End If
            End If
         Else
            MsgBox "El pedido no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Número de pedido incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_busqueda_LostFocus()
   Me.frm_busqueda.Visible = False
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad_eliminar) Then
         If CDbl(Me.txt_cantidad_eliminar) <= CDbl(Me.lv_pedidos.selectedItem.SubItems(2)) Then
            rsaux.Open "UPDATE TB_dETALLE_PEDIDOS SET FLOA_PED_CANTIDAD = ISNULL(FLOA_PED_CANTIDAD,0) - " + Me.txt_cantidad_eliminar + " WHERE  INTE_PED_NUMERO = " + CStr(Me.txt_numero_pedido) + " AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn_pedido_cantia_textilera, adOpenDynamic, adLockOptimistic
            Me.lv_pedidos.selectedItem.SubItems(2) = Format(CDbl(Me.lv_pedidos.selectedItem.SubItems(2)) - CDbl(Me.txt_cantidad_eliminar), "###,###,##0.00")
            Me.lv_pedidos.selectedItem.SubItems(4) = Format(CDbl(Me.lv_pedidos.selectedItem.SubItems(2)) - CDbl(Me.lv_pedidos.selectedItem.SubItems(3)), "###,###,##0.00")
            
            var_importe = 0
            var_j = Me.lv_pedidos.selectedItem.Index
            For var_i = 1 To lv_pedidos.ListItems.Count
                lv_pedidos.ListItems.Item(var_i).Selected = True
                var_importe = var_importe + CDbl(Me.lv_pedidos.selectedItem.SubItems(4))
            Next var_i
            Me.txt_subimporte = Format(var_importe, "###,###,##0.00")
            Me.txt_importe_descuento = Format(CDbl(Me.txt_subimporte) - (CDbl(Me.txt_subimporte) * (1 - (CDbl(Me.txt_descuento) / 100))), "###,###,##0.00")
            Me.txt_importe = Format(CDbl(Me.txt_subimporte) - CDbl(Me.txt_importe_descuento), "###,###,##0.00")
            Me.lv_pedidos.ListItems.Item(var_j).Selected = True
            Me.lv_pedidos.SetFocus
         Else
            MsgBox "La cantidad a eliminar no puede ser mayor a la cantidad en el pedido", vbOKOnly, "ATENCION"
            Me.lv_pedidos.SetFocus
         End If
      Else
         MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
         Me.lv_pedidos.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_eliminar.Visible = False
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   Me.frm_eliminar.Visible = False
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      If KeyAscii = 27 Then
       
      End If
   End If
End Sub

Private Sub txt_cantidad_LostFocus()
   If Not IsNumeric(Me.txt_cantidad) Then
      MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub txt_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_cli_clave_id,vcha_cli_nombre from vw_clientes where vcha_age_agente_id = '00231' order by vcha_cli_nombre ", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      var_tipo_lista = 5
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_cliente_LostFocus()
   If Me.txt_cliente <> "" Then
      rs.Open "select * from vw_clientes where vcha_Cli_clave_id = '" + Me.txt_cliente + "' and vcha_Emp_Empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
      Else
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
         Me.txt_cliente = ""
         Me.txt_nombre_cliente = ""
      End If
      rs.Close
   Else
      'MsgBox "Se debe de seleccionar un cliente", vbOKOnly, "ATENCION"
      Me.txt_cliente = ""
      Me.txt_nombre_cliente = ""
   End If
End Sub

Private Sub txt_codigo_GotFocus()
   Me.txt_descripcion = ""
   Me.txt_codigo = ""
   Me.txt_precio = ""
   Me.txt_cantidad = ""
   Me.txt_alto = ""
   Me.txt_ancho = ""
   Me.txt_largo = ""
   Me.chk_blackout.Value = 0
   Me.txt_nota_1 = ""
   Me.txt_nota_2 = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   If Me.txt_codigo <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      If Len(Me.txt_codigo) = 11 Then
         rs.Open "SELECT * FROM TB_aRTICULOS WHERE substring(VCHA_aRT_ARTICULO_ID,1,11) = '" + Me.txt_codigo + "'", cnn_pedido_cantia_textilera, adOpenDynamic, adLockOptimistic
      Else
         rs.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn_pedido_cantia_textilera, adOpenDynamic, adLockOptimistic
      End If
      If Not rs.EOF Then
         Me.txt_codigo = rs!vcha_Art_Articulo_id
         Me.txt_descripcion = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
      Else
         rsaux.Open "SELECT * FROM TB_EQUIVALENCIAS WHERE VCHA_EQU_CODIGO_EQUIVALENTE = '" + Me.txt_codigo + "'", cnn_pedido_cantia_textilera, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            Me.txt_codigo = IIf(IsNull(rs!vcha_Art_Articulo_id), "", rs!vcha_Art_Articulo_id)
            If Me.txt_codigo <> "" Then
               rsaux1.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn_pedido_cantia_textilera, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  Me.txt_codigo = rsaux1!vcha_Art_Articulo_id
                  Me.txt_descripcion = IIf(IsNull(rsaux1!vcha_Art_nombre_español), "", rsaux1!vcha_Art_nombre_español)
                  Me.txt_cantidad = ""
                  Me.txt_precio = ""
               Else
                  MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                  Me.txt_descripcion = ""
                  Me.txt_cantidad = ""
                  Me.txt_precio = ""
                  Me.txt_codigo = ""
               End If
               rsaux1.Close
            Else
               MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
               Me.txt_descripcion = ""
               Me.txt_cantidad = ""
               Me.txt_precio = ""
               Me.txt_codigo = ""
            End If
         Else
            MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
            Me.txt_descripcion = ""
            Me.txt_cantidad = ""
            Me.txt_precio = ""
         End If
         rsaux.Close
      End If
      rs.Close
   Else
      Me.txt_descripcion = ""
      Me.txt_cantidad = ""
      Me.txt_precio = ""
   End If
End Sub

Private Sub txt_nota_1_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      If KeyAscii = 27 Then
       
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_descuento_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Me.txt_importe_descuento.SetFocus
   End If
End Sub

Private Sub txt_importe_descuento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_importe.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_importe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_largo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_largo_LostFocus()
   If Not IsNumeric(Me.txt_largo) Then
      MsgBox "Medida de largo incorrecta"
      Me.txt_largo = ""
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      If KeyAscii = 27 Then
       
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_nota_2_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_notas_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_codigo.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_precio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_subimporte_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_importe_descuento.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub
