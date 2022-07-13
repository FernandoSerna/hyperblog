VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmgenera_pedidos_multibondeados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generación de pedidos"
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
   Begin VB.TextBox txt_foco 
      Enabled         =   0   'False
      Height          =   315
      Left            =   12645
      TabIndex        =   19
      Top             =   2160
      Width           =   765
   End
   Begin VB.Frame frm_busqueda 
      Height          =   915
      Left            =   600
      TabIndex        =   16
      Top             =   270
      Width           =   2760
      Begin VB.TextBox txt_busqueda 
         Height          =   360
         Left            =   135
         TabIndex        =   17
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
         TabIndex        =   18
         Top             =   120
         Width           =   2685
      End
   End
   Begin VB.Frame frm_articulos 
      Height          =   3015
      Left            =   3360
      TabIndex        =   14
      Top             =   4020
      Width           =   5550
      Begin VB.ListBox lst_articulos 
         Height          =   2790
         Left            =   75
         TabIndex        =   15
         Top             =   150
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11145
      Picture         =   "frmgenera_pedidos_multibondeados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   765
      Picture         =   "frmgenera_pedidos_multibondeados.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Imprimir Movimiento Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmgenera_pedidos_multibondeados.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Buscar Pedido Alt + B"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmgenera_pedidos_multibondeados.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Nuevo Pedido Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   2070
      TabIndex        =   7
      Top             =   405
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   8
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
         TabIndex        =   9
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cargar pedido"
      Height          =   315
      Left            =   3045
      TabIndex        =   6
      Top             =   30
      Width           =   1875
   End
   Begin VB.Frame frm_disponibles 
      Height          =   3720
      Left            =   2730
      TabIndex        =   2
      Top             =   4185
      Width           =   6990
      Begin VB.TextBox txt_nombre_articulo 
         Height          =   345
         Left            =   90
         TabIndex        =   3
         Top             =   330
         Width           =   6720
      End
      Begin MSComctlLib.ListView lv_disponibles 
         Height          =   2910
         Left            =   15
         TabIndex        =   4
         Top             =   645
         Width           =   6825
         _ExtentX        =   12039
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
            Text            =   "Código"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre del Artículo"
            Object.Width           =   7057
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Disponible"
            Object.Width           =   2470
         EndProperty
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000D&
         Caption         =   " Artículos Disponibles"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   6990
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "refacturacion"
      Height          =   315
      Left            =   5295
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CommandButton cmd_cargar_pedido 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1095
      Picture         =   "frmgenera_pedidos_multibondeados.frx":0940
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Genera orden de surtido"
      Top             =   15
      Width           =   330
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   15
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
            Picture         =   "frmgenera_pedidos_multibondeados.frx":0A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgenera_pedidos_multibondeados.frx":131C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgenera_pedidos_multibondeados.frx":1BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgenera_pedidos_multibondeados.frx":2192
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgenera_pedidos_multibondeados.frx":2A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgenera_pedidos_multibondeados.frx":3348
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgenera_pedidos_multibondeados.frx":3C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgenera_pedidos_multibondeados.frx":3D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgenera_pedidos_multibondeados.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgenera_pedidos_multibondeados.frx":3F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgenera_pedidos_multibondeados.frx":406A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgenera_pedidos_multibondeados.frx":417C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   15
      TabIndex        =   65
      Top             =   285
      Width           =   11520
   End
   Begin VB.Frame Frame1 
      Height          =   2100
      Left            =   120
      TabIndex        =   37
      Top             =   405
      Width           =   11430
      Begin VB.TextBox txt_establecimiento 
         Height          =   315
         Left            =   6915
         TabIndex        =   53
         Top             =   1365
         Width           =   1125
      End
      Begin VB.TextBox txt_clave_cliente 
         Height          =   315
         Left            =   1245
         TabIndex        =   52
         Top             =   1695
         Width           =   1125
      End
      Begin VB.TextBox txt_tipo_pedido 
         Height          =   315
         Left            =   1245
         TabIndex        =   51
         Top             =   1035
         Width           =   1125
      End
      Begin VB.TextBox txt_numero 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   450
         Width           =   1680
      End
      Begin VB.TextBox txt_fecha 
         Enabled         =   0   'False
         Height          =   315
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   450
         Width           =   1005
      End
      Begin VB.TextBox txt_titular 
         Height          =   315
         Left            =   1245
         TabIndex        =   48
         Top             =   1365
         Width           =   1125
      End
      Begin VB.TextBox txt_descuento1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   6915
         TabIndex        =   47
         Top             =   1695
         Width           =   810
      End
      Begin VB.TextBox txt_plazo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   10455
         TabIndex        =   46
         Top             =   1695
         Width           =   810
      End
      Begin VB.TextBox txt_descuento2 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   8955
         TabIndex        =   45
         Top             =   1695
         Width           =   810
      End
      Begin VB.Frame Frame6 
         Height          =   90
         Left            =   45
         TabIndex        =   44
         Top             =   780
         Width           =   11355
      End
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   6915
         TabIndex        =   43
         Top             =   1035
         Width           =   1125
      End
      Begin VB.TextBox txt_nombre_tipo_pedido 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2385
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1035
         Width           =   3210
      End
      Begin VB.TextBox txt_nombre_titular 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2385
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1365
         Width           =   3210
      End
      Begin VB.TextBox txt_nombre_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2385
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1695
         Width           =   3210
      End
      Begin VB.TextBox txt_nombre_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8055
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1035
         Width           =   3210
      End
      Begin VB.TextBox txt_nombre_establecimiento 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8055
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1365
         Width           =   3210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Pedido:"
         Height          =   195
         Left            =   90
         TabIndex        =   64
         Top             =   1095
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   90
         TabIndex        =   63
         Top             =   1755
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Left            =   5730
         TabIndex        =   62
         Top             =   1425
         Width           =   1155
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   " Datos Generales "
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
         TabIndex        =   61
         Top             =   120
         Width           =   11355
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   2040
         TabIndex        =   60
         Top             =   510
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   285
         TabIndex        =   59
         Top             =   510
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Left            =   90
         TabIndex        =   58
         Top             =   1425
         Width           =   480
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Descuento 1:"
         Height          =   195
         Left            =   5730
         TabIndex        =   57
         Top             =   1755
         Width           =   960
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Plazo:"
         Height          =   195
         Left            =   10005
         TabIndex        =   56
         Top             =   1755
         Width           =   435
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Descuento 2:"
         Height          =   195
         Left            =   7920
         TabIndex        =   55
         Top             =   1770
         Width           =   960
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   5730
         TabIndex        =   54
         Top             =   1095
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Left            =   120
      TabIndex        =   31
      Top             =   2505
      Width           =   11430
      Begin VB.TextBox txt_disponible 
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
         ForeColor       =   &H000000C0&
         Height          =   480
         Left            =   8490
         TabIndex        =   71
         Top             =   690
         Width           =   2205
      End
      Begin VB.TextBox txt_apartados 
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
         ForeColor       =   &H000000C0&
         Height          =   480
         Left            =   4830
         TabIndex        =   70
         Top             =   690
         Width           =   2205
      End
      Begin VB.TextBox txt_existen 
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
         Left            =   1170
         TabIndex        =   67
         Top             =   690
         Width           =   2205
      End
      Begin VB.TextBox txt_codigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   690
         TabIndex        =   34
         Top             =   180
         Width           =   2550
      End
      Begin VB.TextBox txt_cantidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   9600
         TabIndex        =   33
         Top             =   165
         Width           =   1635
      End
      Begin VB.TextBox txt_Articulo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   480
         Left            =   3270
         TabIndex        =   32
         Top             =   180
         Width           =   5310
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Disponibles:"
         Height          =   195
         Left            =   7575
         TabIndex        =   69
         Top             =   810
         Width           =   855
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Apartados:"
         Height          =   195
         Left            =   3960
         TabIndex        =   68
         Top             =   810
         Width           =   765
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Existen:"
         Height          =   195
         Left            =   435
         TabIndex        =   66
         Top             =   810
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   90
         TabIndex        =   36
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   8910
         TabIndex        =   35
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3525
      Left            =   120
      TabIndex        =   20
      Top             =   3705
      Width           =   11430
      Begin VB.TextBox txt_suma_cantidad 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   9060
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   3930
         Width           =   1245
      End
      Begin VB.TextBox txt_suma_importe 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   3930
         Width           =   1215
      End
      Begin VB.Frame frm_eliminar 
         Height          =   975
         Left            =   4950
         TabIndex        =   21
         Top             =   1560
         Width           =   2205
         Begin VB.TextBox txt_eliminar 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   90
            TabIndex        =   22
            Top             =   390
            Width           =   1995
         End
         Begin VB.Label Label11 
            BackColor       =   &H8000000D&
            Caption         =   " Cantidad a Eliminar"
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
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   15
            Width           =   2685
         End
      End
      Begin MSComctlLib.ListView lv_pedidos 
         Height          =   2565
         Left            =   30
         TabIndex        =   26
         Top             =   390
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   4524
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
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   10936
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Precio "
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cantidad "
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Importe    "
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Tipo"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Totales:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8220
         TabIndex        =   30
         Top             =   3990
         Width           =   705
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         Caption         =   " Detalle del Pedido"
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
         TabIndex        =   29
         Top             =   120
         Width           =   11340
      End
      Begin VB.Label lbl_total 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999999999999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8970
         TabIndex        =   28
         Top             =   3030
         Width           =   2070
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Piezas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7080
         TabIndex        =   27
         Top             =   3030
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmgenera_pedidos_multibondeados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_primera_vez As Boolean
Dim var_cantidad_pedida As Variant
Dim var_precio_pedido As Variant
Dim var_nombre_articulo As String
Dim var_tipo_cliente As String
Dim var_suma_cantidad As Variant
Dim var_suma_importe As Variant
Dim var_descuento_1 As Variant
Dim var_descuento_2 As Variant
Dim var_descuento_3 As Variant
Dim var_dias_condiciones As Integer
Dim var_dias_caducidad As Integer
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_origen_codigo As Integer
Dim var_almacen As String
Dim var_lista_precios As String
Dim var_canal_venta As String
Dim var_clave_moneda As String
Dim var_resurtible As Integer
Dim var_tipo_lista As Integer
Dim var_renglon As Double
Dim var_estatus As String
Dim canal_venta As String



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
          lv_pedidos.ListItems.Item(var_i).ForeColor = &H8000&
          lv_pedidos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_pedidos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
          lv_pedidos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H8000&
          lv_pedidos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H8000&
          lv_pedidos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H8000&
       Else
          lv_pedidos.ListItems.Item(var_i).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(1).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(2).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(3).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(4).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(5).Bold = False
          lv_pedidos.ListItems.Item(var_i).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H80000012
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_pedidos.ListItems.Item(var_renglon).Selected = True
      lv_pedidos.selectedItem.EnsureVisible
   End If
   lv_pedidos.Refresh
End Sub

Private Sub cmd_buscar_Click()
         frm_busqueda.Visible = True
         txt_busqueda.SetFocus
End Sub


Private Sub cmd_cargar_pedido_Click()
   var_activa_forma_ordensurtido = "frmgenera_pedidos_multibondeados"
   Me.Enabled = False
   frmordensurtido.Show
End Sub

Private Sub cmd_imprimir_Click()
   Set TB_ENC_PEDIDOS_M = New TB_ENC_PEDIDOS_M
   If Trim(txt_numero) <> "" Then
      If Trim(var_estatus) = "" Then
         var_si = MsgBox("¿Se va a cerrar el pedido?", vbOKCancel, "ATENCION")
         If var_si = 1 Then
            'cnn.BeginTrans
            var_estatus = "I"
            ok = TB_ENC_PEDIDOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_numero, "I")
            If Trim(Me.txt_tipo_pedido) = "T" Or Trim(Me.txt_tipo_pedido) = "V" Then
               rsaux2.Open "update tb_encabezado_pedidos set INTE_PED_AUTORIZO = 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'  and vcha_alm_almacen_id = '" + var_almacen + "' and inte_ped_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
            End If
            If var_empresa = "16" Then
               rsaux.Open "SELECT a.vcha_tit_titular_id,b.vcha_tit_nombre,sum(Case when dias > 0 then floa_sal_importe else 0 end) as Vencido, sum(Case when dias < 1 then floa_sal_importe else 0 end) as Por_Vencer from vw_facturas_venciadas a inner join vw_jv_clientes b on a.vcha_cli_clave_id = b.vcha_cli_clave_id where (b.vcha_tit_titular_id = '" + Me.txt_titular + "') GROUP BY a.vcha_tit_titular_id,b.vcha_tit_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  If rsaux!vencido <= 10 Then
                     rsaux2.Open "Select tb_titulares.vcha_tit_titular_id,tb_titulares.vcha_tit_nombre,tb_titulares.floa_tit_limite_credito From TB_TITULARES where tb_titulares.vcha_tit_titular_id = '" + Me.txt_titular + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        rsaux5.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_cLI_CLAVE_ID = '" + Me.txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_plazo = IIf(IsNull(rsaux5!VCHA_PLA_PLAZO_ID), "", rsaux5!VCHA_PLA_PLAZO_ID)
                        If var_plazo = "4" Then
                           rsaux4.Open "UPDATE TB_ENCABEZADO_PEDIDOS SET INTE_PED_AUTORIZO = 1, VCHA_PED_AUTORIZO = '8', DTIM_PED_AUTORIZO =  GETDATE(), VCHA_PED_AUTORIZACION = 'CLIENTE DE CONTADO' WHERE INTE_PED_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                           MsgBox "El pedido a sido autorizado", vbOKOnly, "ATENCION"
                        Else
                           var_cadena = "SELECT SUM(dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD * dbo.TB_DETALLE_PEDIDOS.FLOA_PED_PRECIO) AS IMPORTE_NETO, dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2, SUM((dbo.TB_DETALLE_PEDIDOS.FLOA_PED_PRECIO * dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD) * (1 - dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1 / 100)) * (1 - dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2 / 100) AS IMPORTE_PEDIDO FROM dbo.TB_ENCABEZADO_PEDIDOS INNER JOIN dbo.TB_DETALLE_PEDIDOS ON dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO = dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO GROUP BY dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2 Having (dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO = " + Me.txt_numero + ")"
                           rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                           If rsaux!Por_Vencer + (rsaux3!importe_pedido * 1.16) < IIf(IsNull(rsaux2!floa_tit_limite_credito), 0, rsaux2!floa_tit_limite_credito) Then
                              rsaux4.Open "UPDATE TB_ENCABEZADO_PEDIDOS SET INTE_PED_AUTORIZO = 1, VCHA_PED_AUTORIZO = '8', DTIM_PED_AUTORIZO =  GETDATE(), VCHA_PED_AUTORIZACION = 'IMPORTE VENCIDO: " + CStr(rsaux!vencido) + "  SALDO POR VENCER: " + CStr(rsaux!Por_Vencer) + " IMPORTE PEDIDO: " + CStr(rsaux3!importe_pedido * 1.16) + " LÍNEA DE CRÉDITO: " + CStr(IIf(IsNull(rsaux2!floa_tit_limite_credito), 0, rsaux2!floa_tit_limite_credito)) + "' WHERE INTE_PED_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                              MsgBox "El pedido a sido autorizado", vbOKOnly, "ATENCION"
                           Else
                              var_cadena = "SELECT SUM(dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD * dbo.TB_DETALLE_PEDIDOS.FLOA_PED_PRECIO) AS IMPORTE_NETO, dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2, SUM((dbo.TB_DETALLE_PEDIDOS.FLOA_PED_PRECIO * dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD) * (1 - dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1 / 100)) * (1 - dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2 / 100) AS IMPORTE_PEDIDO FROM dbo.TB_ENCABEZADO_PEDIDOS INNER JOIN dbo.TB_DETALLE_PEDIDOS ON dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO = dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO GROUP BY dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2 Having (dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO = " + Me.txt_numero + ")"
                              rsaux4.Open "update tb_encabezado_pedidos set VCHA_PED_AUTORIZACION = 'IMPORTE VENCIDO: " + CStr(rsaux!vencido) + "  SALDO POR VENCER: " + CStr(rsaux!Por_Vencer) + " IMPORTE PEDIDO: " + CStr(rsaux3!importe_pedido * 1.16) + " LÍNEA DE CRÉDITO: " + CStr(IIf(IsNull(rsaux2!floa_tit_limite_credito), 0, rsaux2!floa_tit_limite_credito)) + "' WHERE INTE_PED_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                              MsgBox "El pedido requiere autorización de crédito y cobranza IMPORTE VENCIDO: " + Format(CStr(rsaux!vencido), "###,###,##0.00") + "  SALDO POR VENCER: " + Format(CStr(rsaux!Por_Vencer), "###,###,##0.00") + " IMPORTE PEDIDO: " + Format(CStr(rsaux3!importe_pedido * 1.16), "###,###,##0.00") + " LÍNEA DE CRÉDITO: " + Format(CStr(IIf(IsNull(rsaux2!floa_tit_limite_credito), 0, rsaux2!floa_tit_limite_credito)), "###,###,##0.00"), vbOKOnly, "ATENCION"
                           End If
                           rsaux3.Close
                        End If
                        rsaux5.Close
                     End If
                     rsaux2.Close
                  Else
                     var_cadena = "SELECT SUM(dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD * dbo.TB_DETALLE_PEDIDOS.FLOA_PED_PRECIO) AS IMPORTE_NETO, dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2, SUM((dbo.TB_DETALLE_PEDIDOS.FLOA_PED_PRECIO * dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD) * (1 - dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1 / 100)) * (1 - dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2 / 100) AS IMPORTE_PEDIDO FROM dbo.TB_ENCABEZADO_PEDIDOS INNER JOIN dbo.TB_DETALLE_PEDIDOS ON dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO = dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO GROUP BY dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2 Having (dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO = " + Me.txt_numero + ")"
                     rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rsaux2.Open "Select tb_titulares.vcha_tit_titular_id,tb_titulares.vcha_tit_nombre,tb_titulares.floa_tit_limite_credito From TB_TITULARES where tb_titulares.vcha_tit_titular_id = '" + Me.txt_titular + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                     rsaux5.Open "update tb_encabezado_pedidos set VCHA_PED_AUTORIZACION = 'IMPORTE VENCIDO: " + CStr(rsaux!vencido) + "  SALDO POR VENCER: " + CStr(rsaux!Por_Vencer) + " IMPORTE PEDIDO: " + CStr(rsaux3!importe_pedido * 1.16) + " LÍNEA DE CRÉDITO: " + CStr(IIf(IsNull(rsaux2!floa_tit_limite_credito), 0, rsaux2!floa_tit_limite_credito)) + "' WHERE INTE_PED_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                     MsgBox "El pedido requiere autorización de crédito y cobranza 'IMPORTE VENCIDO: " + CStr(rsaux!vencido) + "  SALDO POR VENCER: " + CStr(rsaux!Por_Vencer) + " IMPORTE PEDIDO: " + CStr(rsaux3!importe_pedido * 1.16) + " LÍNEA DE CRÉDITO: " + CStr(IIf(IsNull(rsaux2!floa_tit_limite_credito), 0, rsaux2!floa_tit_limite_credito)), vbOKOnly, "ATENCION"
                     rsaux2.Close
                     rsaux3.Close
                  End If
               Else
                  rsaux2.Open "Select tb_titulares.vcha_tit_titular_id,tb_titulares.vcha_tit_nombre,tb_titulares.floa_tit_limite_credito From TB_TITULARES where tb_titulares.vcha_tit_titular_id = '" + Me.txt_titular + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                  var_cadena = "SELECT SUM(dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD * dbo.TB_DETALLE_PEDIDOS.FLOA_PED_PRECIO) AS IMPORTE_NETO, dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2, SUM((dbo.TB_DETALLE_PEDIDOS.FLOA_PED_PRECIO * dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD) * (1 - dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1 / 100)) * (1 - dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2 / 100) AS IMPORTE_PEDIDO FROM dbo.TB_ENCABEZADO_PEDIDOS INNER JOIN dbo.TB_DETALLE_PEDIDOS ON dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO = dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO GROUP BY dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2 Having (dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO = " + Me.txt_numero + ")"
                  rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rsaux4.Open "UPDATE TB_ENCABEZADO_PEDIDOS SET INTE_PED_AUTORIZO = 1, VCHA_PED_AUTORIZO = '8', DTIM_PED_AUTORIZO =  GETDATE(), VCHA_PED_AUTORIZACION = 'IMPORTE VENCIDO: 0  SALDO POR VENCER: 0 IMPORTE PEDIDO: " + CStr(rsaux3!importe_pedido * 1.16) + " LÍNEA DE CRÉDITO: " + CStr(IIf(IsNull(rsaux2!floa_tit_limite_credito), 0, rsaux2!floa_tit_limite_credito)) + "' WHERE INTE_PED_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                  rsaux3.Close
                  rsaux2.Close
                  MsgBox "El pedido a sido autorizado", vbOKOnly, "ATENCION"
               End If
               rsaux.Close
               'rsaux2.Open "update tb_encabezado_pedidos set INTE_PED_AUTORIZO = 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'  and vcha_alm_almacen_id = '" + var_almacen + "' and inte_ped_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
            End If
            'cnn.CommitTrans
            Set reporte = appl.OpenReport(App.Path + "\rep_PEDIDos_1.rpt")
            reporte.RecordSelectionFormula = "{VW_PEDIDOS.INTE_PED_NUMERO} = " + txt_numero
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Pedidos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            txt_codigo.Enabled = False
            txt_cantidad.Enabled = False
            txt_foco.Enabled = False
            If Me.lv_disponibles.ListItems.Count > 0 Then
               Me.lv_disponibles.SetFocus
            End If
         End If
      Else
         Set reporte = appl.OpenReport(App.Path + "\rep_PEDIDos_1.rpt")
         reporte.RecordSelectionFormula = "{VW_PEDIDOS.INTE_PED_NUMERO} = " + txt_numero
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Pedidos"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         txt_codigo.Enabled = False
         txt_cantidad.Enabled = False
         txt_foco.Enabled = False
         If Me.lv_disponibles.ListItems.Count > 0 Then
            Me.lv_disponibles.SetFocus
         End If
      End If
   End If
End Sub

Private Sub cmd_nuevo_Click()
   lbl_total = "0"
   var_estatus = ""
   txt_tipo_pedido = ""
   txt_titular = ""
   txt_establecimiento = ""
   txt_clave_cliente = ""
   txt_nombre_tipo_pedido = ""
   txt_nombre_agente = ""
   txt_nombre_titular = ""
   txt_nombre_establecimiento = ""
   txt_nombre_cliente = ""
   var_cantidad_pedida = 0
   var_precio_pedido = 0
   var_primera_vez = True
   frm_articulos.Visible = False
   lv_pedidos.ListItems.Clear
   txt_fecha = Date
   txt_numero = ""
   txt_codigo = ""
   txt_descuento1 = ""
   txt_descuento2 = ""
   txt_plazo = ""
   var_suma_cantidad = 0
   var_suma_importe = 0
   txt_suma_cantidad = Format(0, "###,###,##0.00")
   txt_suma_importe = Format(0, "###,###,##0.00")
   txt_tipo_pedido.Enabled = True
   txt_codigo.Enabled = False
   txt_cantidad.Enabled = False
   txt_foco.Enabled = False
   txt_agente = ""
   If var_empresa <> "16" Then
   rsaux5.Open "select * from TB_USUARIOS_PEDIDOS_VISTAS where vcha_usu_usuario_id = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux5.EOF Then
      Me.txt_codigo.Enabled = True
      Me.txt_tipo_pedido = "V"
      Frmmenu2.StatusBar1.Panels(1) = ""
      If Trim(txt_tipo_pedido) <> "" Then
         txt_tipo_pedido = UCase(txt_tipo_pedido)
         rs.Open "select * from tb_tipopedidos where char_tpe_tipo_pedido_id = '" + txt_tipo_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_tipo_cliente = rs!VCHA_TCL_TIPO_CLIENTE_ID
            txt_nombre_tipo_pedido = rs!VCHA_tpe_NOMBRE
            rs.Close
            txt_agente.Enabled = True
            txt_tipo_pedido.Enabled = False
         Else
            rs.Close
            MsgBox "Tipo de pedido incorrecto", vbOKOnly, "ATENCION"
            txt_tipo_pedido = ""
            txt_nombre_tipo_pedido = ""
            txt_agente.Enabled = False
         End If
      End If
      
      Me.txt_agente = IIf(IsNull(rsaux5!VCHA_AGE_AGENTE_ID), "", rsaux5!VCHA_AGE_AGENTE_ID)
      Frmmenu2.StatusBar1.Panels(1) = ""
      If Trim(txt_agente) <> "" Then
         txt_agente = UCase(txt_agente)
         rs.Open "select * from vw_pedidos_2 where char_tpe_tipo_pedido_id = '" + txt_tipo_pedido + "' and vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_nombre_agente = rs!VCHA_AGE_NOMBRE
            var_agente = rs!VCHA_AGE_AGENTE_ID
            canal_venta = rs!vcha_can_canal_venta_id
            rs.Close
            txt_titular.Enabled = True
            txt_agente.Enabled = False
         Else
            rs.Close
            MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
            txt_agente = ""
            txt_nombre_agente = ""
            canal_venta = ""
            txt_titular.Enabled = False
         End If
      End If
      Me.txt_titular = IIf(IsNull(rsaux5!vcha_tit_titular_id), "", rsaux5!vcha_tit_titular_id)
      Frmmenu2.StatusBar1.Panels(1) = ""
      If Trim(txt_titular) <> "" Then
         txt_titular = UCase(txt_titular)
         rs.Open "select * from vw_pedidos_2 where vcha_tit_titular_id = '" + txt_titular + "' and VCHA_TCL_TIPO_CLIENTE_ID = '" + var_tipo_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_nombre_titular = rs!VCHA_TIT_NOMBRE
            If txt_tipo_pedido = "V" Then
               rs.Close
               rs.Open "select distinct floa_gac_descuento_1, floa_gac_descuento_2,inte_pla_dias,inte_tpe_dias_caducidad,floa_gac_descuento_3,vcha_esb_establecimiento_id,vcha_esb_nombre,vcha_cli_clave_id,vcha_cli_nombre,vcha_lis_lista_id, vcha_can_canal_venta_id, inte_tpe_resurtible, vcha_mon_moneda_id from vw_pedidos_2 where VCHA_TCL_TIPO_CLIENTE_ID = '" + var_tipo_cliente + "' and vcha_tit_titular_id = '" + txt_titular + "' order by vcha_esb_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
               If Not rs.EOF Then
                  var_lista_precios = IIf(IsNull(rs!vcha_LIS_LISTA_iD), "", rs!vcha_LIS_LISTA_iD)
                  txt_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
                  Me.txt_nombre_establecimiento = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
                  txt_clave_cliente = rs!vcha_cli_clave_id
                  txt_nombre_cliente = rs!VCHA_CLI_NOMBRE
                  txt_establecimiento.Enabled = False
                  txt_titular.Enabled = False
                  txt_clave_cliente.Enabled = False
                  
                  var_lista_precios = IIf(IsNull(rs!vcha_LIS_LISTA_iD), "", rs!vcha_LIS_LISTA_iD)
                  var_canal_venta = IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id)
                  var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                  var_canal_venta = IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id)
                  var_resurtible = IIf(IsNull(rs!inte_tpe_resurtible), 0, rs!inte_tpe_resurtible)
                  
                  If IsNull(rs(0).Value) Then
                     var_descuento_1 = 0
                     txt_descuento1 = Format(var_descuento_1, "##0.000")
                  Else
                     var_descuento_1 = rs(0).Value
                     txt_descuento1 = Format(rs(0).Value, "##0.000")
                  End If
                  If IsNull(rs(1).Value) Then
                     var_descuento_2 = 0
                     txt_descuento2 = Format(var_descuento_2, "##0.000")
                  Else
                     var_descuento_2 = rs(1).Value
                     txt_descuento2 = Format(var_descuento_2, "##0.000")
                  End If
                  If IsNull(rs(2).Value) Then
                     txt_plazo = 0
                     var_dias_condiciones = 0
                  Else
                     txt_plazo = rs(2).Value
                     var_dias_condiciones = rs(2).Value
                  End If
                  If IsNull(rs(3).Value) Then
                     var_dias_caducidad = 0
                  Else
                     var_dias_caducidad = rs(3).Value
                  End If
                  txt_codigo.Enabled = True
                  'txt_codigo.SetFocus
               Else
                  MsgBox "El titular no tiene relacionado algun establecimiento o un cliente", vbOKOnly, "ATENCION"
               End If
               rs.Close
            Else
               txt_establecimiento.Enabled = True
               rs.Close
               txt_titular.Enabled = False
            End If
         Else
            rs.Close
            txt_titular = ""
            txt_nombre_titular = ""
            txt_establecimiento.Enabled = False
            MsgBox "Titular Incorrecto", vbOKOnly, "ATENCION"
         End If
      End If
      'Me.txt_establecimiento = IIf(IsNull(rsaux5!vcha_Esb_establecimiento_id), "", rsaux5!vcha_Esb_establecimiento_id)
      'txt_establecimiento_LostFocus
      'Me.txt_clave_cliente = IIf(IsNull(rsaux5!vcha_cli_clave_id), "", rsaux5!vcha_cli_clave_id)
      'txt_clave_cliente_LostFocus
      Me.txt_codigo.Enabled = True
      Me.txt_codigo.SetFocus
   Else
      txt_tipo_pedido.SetFocus
   End If
   rsaux5.Close
   End If
   If var_empresa = "16" Then
      Me.txt_tipo_pedido = "M"
      var_tipo_cliente = "M"
      Me.txt_nombre_tipo_pedido = "MAYOREO"
      Me.txt_tipo_pedido.Enabled = False
      Me.txt_nombre_tipo_pedido.Enabled = False
      Me.txt_agente.Enabled = True
      Me.txt_agente.SetFocus
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   rsaux5.Open "SELECT * FROM PEDIDO_PRICE", cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux5.EOF
         txt_codigo.Text = rsaux5!VCHA_ART_ARTICULO_ID
         var_cantidad_pedida = rsaux5!Cantidad
   Set TB_ENC_PEDIDOS_I = New TB_ENC_PEDIDOS_I
   Set TB_DETALLE_PEDIDOS_I = New TB_DETALLE_PEDIDOS_I
   Set TB_DETALLE_PEDIDOS_M = New TB_DETALLE_PEDIDOS_M
   Dim var_i As Integer
   Dim var_n As Integer
   Dim var_precio_anterior As Variant
   Dim list_item As ListItem
   Dim var_catalogo As String
   Dim var_numero_dias As Double
   Dim var_otorga_oferta As Boolean
   Dim var_posible As Boolean
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   Dim agrupador_catalogo As String
   Dim var_precio_externo As Double
   Dim var_catalogo_EFASA As String
   var_origen_codigo = 0
   If var_lista_precios <> "" Then
      If Trim(var_clave_moneda) <> "" Then
         If Trim(txt_codigo) <> "" Then
            If rsaux4.State = 1 Then
               rsaux4.Close
            End If
            rsaux4.Open "select * from tb_detalle_lista_precios where VCHA_LIS_LISTA_PRECIOS_ID = '" + var_lista_precios + "' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux4.EOF Then
               var_precio_pedido = IIf(IsNull(rsaux4!floa_dli_precio), 0, rsaux4!floa_dli_precio)
               If var_precio_pedido > 0 Then
                  'cnn.BeginTrans
                  var_promocion_1 = 0
                  var_promocion_2 = 0
                  var_precio_pedido = rsaux4!floa_dli_precio
                  
                  var_promociones_ya_no = 0
                  If var_promociones_ya_no = 1 Then
                     rs.Open "select * from vw_lista_precios_clientes where vcha_cli_clave_id = '" + txt_clave_cliente + "' and vcha_lis_lista_id = '" + var_lista_precios + "' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_promocion_1 = 0
                        var_promocion_2 = 0
                        var_precio_pedido = rs!floa_dli_precio
                        var_catalogo = rs!vcha_cat_catalogo_id
                        var_otorga_oferta = False
                        If Not IsNull(rs!dtim_vig_fecha_fin) Then
                           var_numero_dias = Date - rs!dtim_vig_fecha_fin
                           var_otorga_oferta = True
                        Else
                           var_otorga_oferta = False
                        End If
                     End If
                     rs.Close
                  
                  
                     rs.Open "select * from vw_descuentos_promociones_vigentes where vcha_can_canal_venta_id = '" + var_canal_venta + "' and vcha_art_Articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_promocion_1 = IIf(IsNull(rs!floa_dpr_desCuento), 0, rs!floa_dpr_desCuento)
                        var_precio_pedido = var_precio_pedido - (var_precio_pedido * (IIf(IsNull(rs!floa_dpr_desCuento), 0, rs!floa_dpr_desCuento) / 100))
                        rs.Close
                     Else
                        rs.Close
                        If var_otorga_oferta = True Then
                           rs.Open "select * from tb_descuentos_catalogos where vcha_can_canal_venta_id = '" + var_canal_venta + "' and inte_des_limite_inferior <= " + Str(var_numero_dias) + " and inte_des_limite_superior >= " + Str(var_numero_dias), cnn, adOpenDynamic, adLockOptimistic
                           If Not rs.EOF Then
                              var_promocion_2 = IIf(IsNull(rs!FLOA_DES_DESCUENTO), 0, rs!FLOA_DES_DESCUENTO)
                              var_precio_pedido = var_precio_pedido - (var_precio_pedido * (rs!FLOA_DES_DESCUENTO / 100))
                           End If
                           rs.Close
                        End If
                     End If
                  End If
                  
                  
                  If var_primera_vez = True Then
                     txt_numero = maximo_pedido
                     var_primera_vez = False
                     ok = TB_ENC_PEDIDOS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_tipo_pedido, maximo_pedido, 0, Date, Date, txt_agente, txt_titular, txt_clave_cliente, txt_establecimiento, var_resurtible, 0, "", var_descuento_1, var_descuento_2, var_descuento_3, var_dias_condiciones, var_dias_caducidad, var_clave_usuario_global, fun_NombrePc, Date, var_clave_moneda, 0)
                     txt_numero = maximo_pedido
                     rs.Open "select * from VW_PEDIDOS where inte_ped_NUMERO = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                     var_suma_cantidad = 0
                     var_suma_importe = 0
                     While Not rs.EOF
                           Set list_item = lv_pedidos.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
                           list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
                           list_item.SubItems(2) = IIf(IsNull(rs!FLOA_PED_PRECIO), Format(0, "###,###,##0.00"), Format(rs!FLOA_PED_PRECIO, "###,###,##0.00"))
                           list_item.SubItems(3) = IIf(IsNull(rs!FLOA_PED_CANTIDAD), Format(0, "###,###,##0.00"), Format(rs!FLOA_PED_CANTIDAD, "###,###,##0.00"))
                           list_item.SubItems(4) = Format(list_item.SubItems(2) * list_item.SubItems(3), "###,###,##0.00")
                           list_item.SubItems(5) = IIf(IsNull(rs!char_ped_tipo), "P", rs!char_ped_tipo)
                           var_renglon = lv_pedidos.ListItems.Count
                           Call ilumina_grid
                           var_suma_cantidad = var_suma_cantidad + list_item.SubItems(3)
                           var_suma_importe = var_suma_importe + list_item.SubItems(4)
                           rs.MoveNext
                     Wend
                     rs.Close
                     txt_suma_cantidad = Format(var_suma_cantidad, "###,###,##0.00")
                     txt_suma_importe = Format(var_suma_importe, "###,###,##0.00")
                     txt_tipo_pedido.Enabled = False
                     txt_titular.Enabled = False
                     txt_establecimiento.Enabled = False
                     txt_clave_cliente.Enabled = False
                     txt_agente.Enabled = False
                  End If
                  rsaux.Open "select * from tb_detalle_pedidos where INTE_PED_NUMERO = " + txt_numero + " and VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND CHAR_PED_TIPO = 'P'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     rsaux.Close
                     valor = txt_codigo
                     rs.Open "update tb_detalle_pedidos set floa_ped_cantidad = floa_ped_cantidad + " + CStr(var_cantidad_pedida) + " where inte_ped_numero = " + txt_numero + " and vcha_art_articulo_id = '" + txt_codigo + "' AND CHAR_PED_TIPO = 'P'", cnn, adOpenDynamic, adLockOptimistic
                     lbl_total = CDbl(lbl_total) + var_cantidad_pedida
                     var_n = lv_pedidos.ListItems.Count
                     var_encontro = 0
                     var_i = 1
                     While (var_i <= var_n)
                         lv_pedidos.ListItems.Item(var_i).Selected = True
                         valor = Trim(lv_pedidos.selectedItem)
                         If txt_codigo = valor Then
                            'If lv_pedidos.selectedItem.SubItems(5) = "P" Then
                            '   var_precio_anterior = (lv_pedidos.selectedItem.SubItems(2) * 1)
                            '   If var_precio_anterior <> var_precio_pedido Then
                            '      var_encontro = 0
                            '   Else
                            var_encontro = 1
                            var_i = var_n
                            '   End If
                            'End If
                         End If
                         var_i = var_i + 1
                     Wend
                     bandera_suma = True
                     convierte_numero (lv_pedidos.selectedItem.SubItems(3))
                     var_cantidad_anterior = Val(numero_devuelto)
                     lv_pedidos.selectedItem.SubItems(3) = Format(var_cantidad_anterior + var_cantidad_pedida, "###,###,##0.00")
                     lv_pedidos.selectedItem.SubItems(4) = Format((var_cantidad_anterior + var_cantidad_pedida) * var_precio_pedido, "###,###,##0.00")
                     var_renglon = lv_pedidos.selectedItem.Index
                     Call ilumina_grid
                     var_suma_cantidad = var_suma_cantidad + var_cantidad_pedida
                     var_suma_importe = var_suma_importe + (var_cantidad_pedida * var_precio_pedido)
                  Else
                     rsaux.Close
                     x = 0
                     If x = 0 Then
                        var_empresa_cliente = ""
                        cnn.CommandTimeout = 360
                        rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_empresa_cliente = IIf(IsNull(rs!VCHA_EMP_EMPRESA_ID), "", rs!VCHA_EMP_EMPRESA_ID)
                        End If
                        rs.Close
                        If var_empresa_cliente = "03" Then
                           rs.Open "SELECT * FROM VW_catalogos_efasa where vcha_Art_articulo_id = '" + txt_codigo + "'"
                           If Not rs.EOF Then
                              var_precio_pedido = var_precio_pedido / (1 - (var_descuento_1 / 100))
                              var_precio_pedido = var_precio_pedido / (1 - (var_descuento_2 / 100))
                           End If
                           rs.Close
                        End If
                     End If

                     
                     
                     ok = TB_DETALLE_PEDIDOS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_numero, txt_codigo, var_precio_pedido, var_cantidad_pedida, 0, var_promocion_1, var_promocion_2, "P")
                     Set list_item = lv_pedidos.ListItems.Add(, , txt_codigo)
                     list_item.SubItems(1) = Trim(txt_Articulo)
                     list_item.SubItems(2) = Format(var_precio_pedido, "###,###,##0.00")
                     list_item.SubItems(3) = Format(var_cantidad_pedida, "###,###,##0.00")
                     list_item.SubItems(4) = Format(var_precio_pedido * var_cantidad_pedida, "###,###,##0.00")
                     list_item.SubItems(5) = "P"
                     var_renglon = lv_pedidos.ListItems.Count
                     Call ilumina_grid
                     var_suma_cantidad = var_suma_cantidad + var_cantidad_pedida
                     var_suma_importe = var_suma_importe + (var_cantidad_pedida * var_precio_pedido)
                     lbl_total = CDbl(lbl_total) + var_cantidad_pedida
                  End If
                  txt_suma_importe = Format(var_suma_importe, "###,###,##0.00")
                  txt_suma_cantidad = Format(var_suma_cantidad, "###,###,##0.00")
                  'cnn.CommitTrans
               Else
                  MsgBox "El precio del artículo esta en ceros", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Este artículo no se encuentra en la lista de precios asignada al cliente", vbOKOnly, "ATENCION"
            End If
            rsaux4.Close
         Else
            MsgBox "Código Incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El cliente no tiene una moneda asociada", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "El cliente no tiene una lista de precios asociada", vbOKOnly, "ATENCION"
   End If
         
         
         rsaux5.MoveNext
   Wend
   rsaux5.Close
End Sub

Private Sub Command2_Click()
   Set TB_ENC_PEDIDOS_I = New TB_ENC_PEDIDOS_I
   Set TB_DETALLE_PEDIDOS_I = New TB_DETALLE_PEDIDOS_I
   Set TB_DETALLE_PEDIDOS_M = New TB_DETALLE_PEDIDOS_M
   Dim var_i As Integer
   Dim var_n As Integer
   Dim var_precio_anterior As Variant
   Dim list_item As ListItem
   Dim var_catalogo As String
   Dim var_numero_dias As Double
   Dim var_otorga_oferta As Boolean
   Dim var_posible As Boolean
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   Dim agrupador_catalogo As String
   Dim var_precio_externo As Double
   Dim var_catalogo_EFASA As String
   rsaux6.Open "select * from pedidos_refacturar", cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux6.EOF
         var_primera_vez = True
         rsaux5.Open "SELECT * FROM PEDIDO_PRICE where inte_ped_numero = " + CStr(rsaux6!pedido), cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux5.EOF
               txt_codigo.Text = rsaux5!VCHA_ART_ARTICULO_ID
               var_cantidad_pedida = rsaux5!Cantidad
               var_origen_codigo = 0
               If var_lista_precios <> "" Then
                  If Trim(var_clave_moneda) <> "" Then
                     If Trim(txt_codigo) <> "" Then
                        If rsaux4.State = 1 Then
                           rsaux4.Close
                        End If
                        rsaux4.Open "select * from tb_detalle_lista_precios where VCHA_LIS_LISTA_PRECIOS_ID = '" + var_lista_precios + "' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux4.EOF Then
                           var_precio_pedido = IIf(IsNull(rsaux4!floa_dli_precio), 0, rsaux4!floa_dli_precio)
                           If var_precio_pedido > 0 Then
                              'cnn.BeginTrans
                              var_promocion_1 = 0
                              var_promocion_2 = 0
                              var_precio_pedido = rsaux4!floa_dli_precio
                  
                              var_promociones_ya_no = 0
                              If var_promociones_ya_no = 1 Then
                                 rs.Open "select * from vw_lista_precios_clientes where vcha_cli_clave_id = '" + txt_clave_cliente + "' and vcha_lis_lista_id = '" + var_lista_precios + "' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rs.EOF Then
                                    var_promocion_1 = 0
                                    var_promocion_2 = 0
                                    var_precio_pedido = rs!floa_dli_precio
                                    var_catalogo = rs!vcha_cat_catalogo_id
                                    var_otorga_oferta = False
                                    If Not IsNull(rs!dtim_vig_fecha_fin) Then
                                       var_numero_dias = Date - rs!dtim_vig_fecha_fin
                                       var_otorga_oferta = True
                                    Else
                                       var_otorga_oferta = False
                                    End If
                                 End If
                                 rs.Close
                   
                  
                                 rs.Open "select * from vw_descuentos_promociones_vigentes where vcha_can_canal_venta_id = '" + var_canal_venta + "' and vcha_art_Articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rs.EOF Then
                                    var_promocion_1 = IIf(IsNull(rs!floa_dpr_desCuento), 0, rs!floa_dpr_desCuento)
                                    var_precio_pedido = var_precio_pedido - (var_precio_pedido * (IIf(IsNull(rs!floa_dpr_desCuento), 0, rs!floa_dpr_desCuento) / 100))
                                    rs.Close
                                 Else
                                    rs.Close
                                    If var_otorga_oferta = True Then
                                       rs.Open "select * from tb_descuentos_catalogos where vcha_can_canal_venta_id = '" + var_canal_venta + "' and inte_des_limite_inferior <= " + Str(var_numero_dias) + " and inte_des_limite_superior >= " + Str(var_numero_dias), cnn, adOpenDynamic, adLockOptimistic
                                       If Not rs.EOF Then
                                          var_promocion_2 = IIf(IsNull(rs!FLOA_DES_DESCUENTO), 0, rs!FLOA_DES_DESCUENTO)
                                          var_precio_pedido = var_precio_pedido - (var_precio_pedido * (rs!FLOA_DES_DESCUENTO / 100))
                                       End If
                                       rs.Close
                                    End If
                                 End If
                              End If
                  
                  
                              If var_primera_vez = True Then
                                 txt_numero = maximo_pedido
                                 var_primera_vez = False
                                 ok = TB_ENC_PEDIDOS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_tipo_pedido, maximo_pedido, 0, Date, Date, txt_agente, txt_titular, txt_clave_cliente, txt_establecimiento, var_resurtible, 0, "", var_descuento_1, var_descuento_2, var_descuento_3, var_dias_condiciones, var_dias_caducidad, var_clave_usuario_global, fun_NombrePc, Date, var_clave_moneda, 0)
                                 txt_numero = maximo_pedido
                                 rs.Open "select * from VW_PEDIDOS where inte_ped_NUMERO = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                                 var_suma_cantidad = 0
                                 var_suma_importe = 0
                                 While Not rs.EOF
                                       Set list_item = lv_pedidos.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
                                       list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
                                       list_item.SubItems(2) = IIf(IsNull(rs!FLOA_PED_PRECIO), Format(0, "###,###,##0.00"), Format(rs!FLOA_PED_PRECIO, "###,###,##0.00"))
                                       list_item.SubItems(3) = IIf(IsNull(rs!FLOA_PED_CANTIDAD), Format(0, "###,###,##0.00"), Format(rs!FLOA_PED_CANTIDAD, "###,###,##0.00"))
                                       list_item.SubItems(4) = Format(list_item.SubItems(2) * list_item.SubItems(3), "###,###,##0.00")
                                       list_item.SubItems(5) = IIf(IsNull(rs!char_ped_tipo), "P", rs!char_ped_tipo)
                                       var_renglon = lv_pedidos.ListItems.Count
                                       Call ilumina_grid
                                       var_suma_cantidad = var_suma_cantidad + list_item.SubItems(3)
                                       var_suma_importe = var_suma_importe + list_item.SubItems(4)
                                       rs.MoveNext
                                 Wend
                                 rs.Close
                                 txt_suma_cantidad = Format(var_suma_cantidad, "###,###,##0.00")
                                 txt_suma_importe = Format(var_suma_importe, "###,###,##0.00")
                                 txt_tipo_pedido.Enabled = False
                                 txt_titular.Enabled = False
                                 txt_establecimiento.Enabled = False
                                 txt_clave_cliente.Enabled = False
                                 txt_agente.Enabled = False
                              End If
                              rsaux.Open "select * from tb_detalle_pedidos where INTE_PED_NUMERO = " + txt_numero + " and VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND CHAR_PED_TIPO = 'P'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux.EOF Then
                                 rsaux.Close
                                 valor = txt_codigo
                                 rs.Open "update tb_detalle_pedidos set floa_ped_cantidad = floa_ped_cantidad + " + CStr(var_cantidad_pedida) + " where inte_ped_numero = " + txt_numero + " and vcha_art_articulo_id = '" + txt_codigo + "' AND CHAR_PED_TIPO = 'P'", cnn, adOpenDynamic, adLockOptimistic
                                 lbl_total = CDbl(lbl_total) + var_cantidad_pedida
                                 var_n = lv_pedidos.ListItems.Count
                                 var_encontro = 0
                                 var_i = 1
                                 While (var_i <= var_n)
                                       lv_pedidos.ListItems.Item(var_i).Selected = True
                                       valor = Trim(lv_pedidos.selectedItem)
                                       If txt_codigo = valor Then
                                          'If lv_pedidos.selectedItem.SubItems(5) = "P" Then
                                          '   var_precio_anterior = (lv_pedidos.selectedItem.SubItems(2) * 1)
                                          '   If var_precio_anterior <> var_precio_pedido Then
                                          '      var_encontro = 0
                                          '   Else
                                                 var_encontro = 1
                                                 var_i = var_n
                                          '   End If
                                          'End If
                                       End If
                                       var_i = var_i + 1
                                 Wend
                                 bandera_suma = True
                                 convierte_numero (lv_pedidos.selectedItem.SubItems(3))
                                 var_cantidad_anterior = Val(numero_devuelto)
                                 lv_pedidos.selectedItem.SubItems(3) = Format(var_cantidad_anterior + var_cantidad_pedida, "###,###,##0.00")
                                 lv_pedidos.selectedItem.SubItems(4) = Format((var_cantidad_anterior + var_cantidad_pedida) * var_precio_pedido, "###,###,##0.00")
                                 var_renglon = lv_pedidos.selectedItem.Index
                                 Call ilumina_grid
                                 var_suma_cantidad = var_suma_cantidad + var_cantidad_pedida
                                 var_suma_importe = var_suma_importe + (var_cantidad_pedida * var_precio_pedido)
                              Else
                                 rsaux.Close
                                 x = 0
                                 If x = 0 Then
                                    var_empresa_cliente = ""
                                    rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rs.EOF Then
                                       var_empresa_cliente = IIf(IsNull(rs!VCHA_EMP_EMPRESA_ID), "", rs!VCHA_EMP_EMPRESA_ID)
                                    End If
                                    rs.Close
                                    If var_empresa_cliente = "03" Then
                                       rs.Open "SELECT * FROM VW_catalogos_efasa where vcha_Art_articulo_id = '" + txt_codigo + "'"
                                       If Not rs.EOF Then
                                          var_precio_pedido = var_precio_pedido / (1 - (var_descuento_1 / 100))
                                          var_precio_pedido = var_precio_pedido / (1 - (var_descuento_2 / 100))
                                       End If
                                       rs.Close
                                    End If
                                 End If
 
                     
                     
                                 ok = TB_DETALLE_PEDIDOS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_numero, txt_codigo, var_precio_pedido, var_cantidad_pedida, 0, var_promocion_1, var_promocion_2, "P")
                                 Set list_item = lv_pedidos.ListItems.Add(, , txt_codigo)
                                 list_item.SubItems(1) = Trim(txt_Articulo)
                                 list_item.SubItems(2) = Format(var_precio_pedido, "###,###,##0.00")
                                 list_item.SubItems(3) = Format(var_cantidad_pedida, "###,###,##0.00")
                                 list_item.SubItems(4) = Format(var_precio_pedido * var_cantidad_pedida, "###,###,##0.00")
                                 list_item.SubItems(5) = "P"
                                 var_renglon = lv_pedidos.ListItems.Count
                                 Call ilumina_grid
                                 var_suma_cantidad = var_suma_cantidad + var_cantidad_pedida
                                 var_suma_importe = var_suma_importe + (var_cantidad_pedida * var_precio_pedido)
                                 lbl_total = CDbl(lbl_total) + var_cantidad_pedida
                              End If
                              txt_suma_importe = Format(var_suma_importe, "###,###,##0.00")
                              txt_suma_cantidad = Format(var_suma_cantidad, "###,###,##0.00")
                              'cnn.CommitTrans
                           Else
                              MsgBox "El precio del artículo esta en ceros", vbOKOnly, "ATENCION"
                           End If
                        Else
                           MsgBox "Este artículo no se encuentra en la lista de precios asignada al cliente", vbOKOnly, "ATENCION"
                        End If
                        rsaux4.Close
                     Else
                        MsgBox "Código Incorrecto", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "El cliente no tiene una moneda asociada", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "El cliente no tiene una lista de precios asociada", vbOKOnly, "ATENCION"
               End If
               rsaux5.MoveNext
         Wend
         rsaux5.Close
         rsaux6.MoveNext
   Wend
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
   
End Sub

Private Sub Form_Load()
   If var_empresa = "16" Or var_empresa = "18" Then
      var_posible_limite_credito = 1
   Else
      var_posible_limite_credito = 0
   End If
   If var_empresa = "02" Or var_empresa = "03" Then
      Command1.Visible = False
   Else
      Command1.Visible = False
   End If
   frm_disponibles.Visible = False
   lbl_total = "0"
   var_cadena_seguridad = ""
   var_tipo_lista = 0
   frm_lista.Visible = False
   Top = 0
   Left = 0
   var_resurtible = 0
   var_lista_precios = ""
   var_cantidad_pedida = 0
   var_precio_pedido = 0
   var_primera_vez = True
   frm_articulos.Visible = False
   txt_fecha = Date
   var_origen_codigo = 0
   var_descuento_1 = 0
   var_descuento_2 = 0
   var_descuento_3 = 0
   txt_titular.Enabled = False
   txt_establecimiento.Enabled = False
   txt_clave_cliente.Enabled = False
   txt_codigo.Enabled = False
   txt_cantidad.Enabled = False
   frm_busqueda.Visible = False
   frm_eliminar.Visible = False
   rs.Open "select * from tb_almacenes where inte_alm_surtir = 1", cnn, adOpenDynamic, adLockOptimistic
   If rs.EOF Then
      MsgBox "No se a configurado algún almacen para surtir mercancia", vbOKOnly, "ATENCION"
      rs.Close
      Unload Me
   Else
      If var_unidad_organizacional = "21" Then
         var_almacen = "AV00013"
      Else
         If var_empresa = "02" And (var_clave_usuario_global = "U0000000170" Or var_clave_usuario_global = "U0000000171") Then
            var_almacen = "AG"
         Else
            If var_empresa = "16" Then
               var_almacen = "28"
            Else
               var_almacen = rs!VCHA_ALM_ALMACEN_ID
            End If
         End If
      End If
      rs.Close
   End If
   If rsaux5.State = 1 Then
      rsaux5.Close
   End If
   rsaux5.Open "select * from TB_USUARIOS_PEDIDOS_VISTAS where vcha_usu_usuario_id = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux5.EOF Then
      Me.txt_codigo.Enabled = True
      Me.txt_tipo_pedido = "V"
      Frmmenu2.StatusBar1.Panels(1) = ""
      If Trim(txt_tipo_pedido) <> "" Then
         txt_tipo_pedido = UCase(txt_tipo_pedido)
         rs.Open "select * from tb_tipopedidos where char_tpe_tipo_pedido_id = '" + txt_tipo_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_tipo_cliente = rs!VCHA_TCL_TIPO_CLIENTE_ID
            txt_nombre_tipo_pedido = rs!VCHA_tpe_NOMBRE
            rs.Close
            txt_agente.Enabled = True
            txt_tipo_pedido.Enabled = False
         Else
            rs.Close
            MsgBox "Tipo de pedido incorrecto", vbOKOnly, "ATENCION"
            txt_tipo_pedido = ""
            txt_nombre_tipo_pedido = ""
            txt_agente.Enabled = False
         End If
      End If
      
      Me.txt_agente = IIf(IsNull(rsaux5!VCHA_AGE_AGENTE_ID), "", rsaux5!VCHA_AGE_AGENTE_ID)
      Frmmenu2.StatusBar1.Panels(1) = ""
      If Trim(txt_agente) <> "" Then
         txt_agente = UCase(txt_agente)
         rs.Open "select * from vw_pedidos_2 where char_tpe_tipo_pedido_id = '" + txt_tipo_pedido + "' and vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_nombre_agente = rs!VCHA_AGE_NOMBRE
            var_agente = rs!VCHA_AGE_AGENTE_ID
            canal_venta = rs!vcha_can_canal_venta_id
            rs.Close
            txt_titular.Enabled = True
            txt_agente.Enabled = False
         Else
            rs.Close
            MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
            txt_agente = ""
            txt_nombre_agente = ""
            canal_venta = ""
            txt_titular.Enabled = False
         End If
      End If
      Me.txt_titular = IIf(IsNull(rsaux5!vcha_tit_titular_id), "", rsaux5!vcha_tit_titular_id)
      Frmmenu2.StatusBar1.Panels(1) = ""
      If Trim(txt_titular) <> "" Then
         txt_titular = UCase(txt_titular)
         rs.Open "select * from vw_pedidos_2 where vcha_tit_titular_id = '" + txt_titular + "' and VCHA_TCL_TIPO_CLIENTE_ID = '" + var_tipo_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_nombre_titular = rs!VCHA_TIT_NOMBRE
            If txt_tipo_pedido = "V" Then
               rs.Close
               rs.Open "select distinct floa_gac_descuento_1, floa_gac_descuento_2,inte_pla_dias,inte_tpe_dias_caducidad,floa_gac_descuento_3,vcha_esb_establecimiento_id,vcha_esb_nombre,vcha_cli_clave_id,vcha_cli_nombre,vcha_lis_lista_id, vcha_can_canal_venta_id, inte_tpe_resurtible, vcha_mon_moneda_id from vw_pedidos_2 where VCHA_TCL_TIPO_CLIENTE_ID = '" + var_tipo_cliente + "' and vcha_tit_titular_id = '" + txt_titular + "' order by vcha_esb_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
               If Not rs.EOF Then
                  var_lista_precios = IIf(IsNull(rs!vcha_LIS_LISTA_iD), "", rs!vcha_LIS_LISTA_iD)
                  txt_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
                  Me.txt_nombre_establecimiento = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
                  txt_clave_cliente = rs!vcha_cli_clave_id
                  txt_nombre_cliente = rs!VCHA_CLI_NOMBRE
                  txt_establecimiento.Enabled = False
                  txt_titular.Enabled = False
                  txt_clave_cliente.Enabled = False
                  
                  var_lista_precios = IIf(IsNull(rs!vcha_LIS_LISTA_iD), "", rs!vcha_LIS_LISTA_iD)
                  var_canal_venta = IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id)
                  var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                  var_canal_venta = IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id)
                  var_resurtible = IIf(IsNull(rs!inte_tpe_resurtible), 0, rs!inte_tpe_resurtible)
                  
                  If IsNull(rs(0).Value) Then
                     var_descuento_1 = 0
                     txt_descuento1 = Format(var_descuento_1, "##0.000")
                  Else
                     var_descuento_1 = rs(0).Value
                     txt_descuento1 = Format(rs(0).Value, "##0.000")
                  End If
                  If IsNull(rs(1).Value) Then
                     var_descuento_2 = 0
                     txt_descuento2 = Format(var_descuento_2, "##0.000")
                  Else
                     var_descuento_2 = rs(1).Value
                     txt_descuento2 = Format(var_descuento_2, "##0.000")
                  End If
                  If IsNull(rs(2).Value) Then
                     txt_plazo = 0
                     var_dias_condiciones = 0
                  Else
                     txt_plazo = rs(2).Value
                     var_dias_condiciones = rs(2).Value
                  End If
                  If IsNull(rs(3).Value) Then
                     var_dias_caducidad = 0
                  Else
                     var_dias_caducidad = rs(3).Value
                  End If
                  txt_codigo.Enabled = True
                  'txt_codigo.SetFocus
               Else
                  MsgBox "El titular no tiene relacionado algun establecimiento o un cliente", vbOKOnly, "ATENCION"
               End If
               rs.Close
            Else
               txt_establecimiento.Enabled = True
               rs.Close
               txt_titular.Enabled = False
            End If
         Else
            rs.Close
            txt_titular = ""
            txt_nombre_titular = ""
            txt_establecimiento.Enabled = False
            MsgBox "Titular Incorrecto", vbOKOnly, "ATENCION"
         End If
      End If
      'Me.txt_establecimiento = IIf(IsNull(rsaux5!vcha_Esb_establecimiento_id), "", rsaux5!vcha_Esb_establecimiento_id)
      'txt_establecimiento_LostFocus
      'Me.txt_clave_cliente = IIf(IsNull(rsaux5!vcha_cli_clave_id), "", rsaux5!vcha_cli_clave_id)
      'txt_clave_cliente_LostFocus
   End If
   rsaux5.Close
   If var_empresa = "18" Then
      Me.txt_tipo_pedido = "M"
      Me.txt_nombre_tipo_pedido = "MAYOREO"
      Me.txt_tipo_pedido.Enabled = False
      Me.txt_nombre_tipo_pedido.Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_generapedido)
End Sub

Private Sub lst_articulos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_articulos.Visible = False
   End If
   If KeyAscii = 13 Then
      txt_codigo = Obtener_llave(cnn, rs, "TB_ARTICULOS", "VCHA_ART_NOMBRE_ESPAÑOL", lst_articulos.Text, 0, "T")
      frm_articulos.Visible = False
      txt_codigo.SetFocus
   End If
End Sub

Private Sub lst_articulos_LostFocus()
   frm_articulos.Visible = False
End Sub

Private Sub lv_disponibles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_disponibles, ColumnHeader)
End Sub

Private Sub lv_disponibles_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If Me.txt_codigo.Enabled = True Then
         Me.txt_codigo.SetFocus
      End If
      Me.frm_disponibles.Visible = False
   End If
   If KeyAscii = 13 Then
      If Me.lv_disponibles.ListItems.Count > 0 Then
         Me.txt_codigo = lv_disponibles.selectedItem
         Me.txt_codigo.SetFocus
      End If
   End If
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   Dim var_n As Integer
   If KeyAscii = 13 Then
      var_n = lv_lista.ListItems.Count
      If var_n > 0 Then
         If var_tipo_lista = 1 Then
            If lv_lista.ListItems.Count > 0 Then
               txt_tipo_pedido = lv_lista.selectedItem
               txt_nombre_tipo_pedido = lv_lista.selectedItem.SubItems(1)
            Else
               txt_tipo_pedido = ""
               txt_nombre_tipo_pedido = ""
            End If
            If txt_tipo_pedido.Enabled = False Then
               txt_tipo_pedido.Enabled = True
            End If
            txt_tipo_pedido.SetFocus
         End If
         If var_tipo_lista = 2 Then
            If lv_lista.ListItems.Count > 0 Then
               txt_agente = lv_lista.selectedItem
               txt_nombre_agente = lv_lista.selectedItem.SubItems(1)
            Else
               txt_agente = ""
               txt_nombre_agente = ""
            End If
            txt_agente.Enabled = True
            txt_agente.SetFocus
         End If
         If var_tipo_lista = 3 Then
            If lv_lista.ListItems.Count > 0 Then
               txt_titular = lv_lista.selectedItem
               txt_nombre_titular = lv_lista.selectedItem.SubItems(1)
            Else
               txt_titular = ""
               txt_nombre_titular = ""
            End If
            txt_titular.Enabled = True
            txt_titular.SetFocus
         End If
         If var_tipo_lista = 4 Then
            If lv_lista.ListItems.Count > 0 Then
               txt_establecimiento = lv_lista.selectedItem
               txt_nombre_establecimiento = lv_lista.selectedItem.SubItems(1)
            Else
               txt_establecimiento = ""
               txt_nombre_establecimiento = ""
            End If
            txt_establecimiento.Enabled = True
            txt_establecimiento.SetFocus
         End If
         If var_tipo_lista = 5 Then
            If lv_lista.ListItems.Count > 0 Then
               txt_clave_cliente = lv_lista.selectedItem
               txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
            Else
               txt_clave_cliente = ""
               txt_nombre_cliente = ""
            End If
            txt_clave_cliente.Enabled = True
            txt_clave_cliente.SetFocus
         End If
         frm_lista.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_pedidos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Me.frm_disponibles.Visible = False
   Call pro_ordena_listas(lv_pedidos, ColumnHeader)
End Sub


Private Sub lv_pedidos_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      frm_eliminar.Visible = True
      txt_eliminar.SetFocus
   End If
End Sub

Private Sub lv_pedidos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub


Private Sub txt_agente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_age_agente_id, vcha_age_nombre from vw_pedidos_2 where VCHA_TCL_TIPO_CLIENTE_ID = '" + var_tipo_cliente + "' order by vcha_age_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Agentes"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      txt_titular.Enabled = True
      txt_titular.SetFocus
   End If
End Sub

Private Sub txt_agente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_agente) <> "" Then
      txt_agente = UCase(txt_agente)
      rs.Open "select * from vw_pedidos_2 where char_tpe_tipo_pedido_id = '" + txt_tipo_pedido + "' and vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_agente = rs!VCHA_AGE_NOMBRE
         var_agente = rs!VCHA_AGE_AGENTE_ID
         canal_venta = rs!vcha_can_canal_venta_id
         rs.Close
         txt_titular.Enabled = True
         txt_agente.Enabled = False
      Else
         rs.Close
         MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
         txt_agente = ""
         txt_nombre_agente = ""
         canal_venta = ""
         txt_titular.Enabled = False
      End If
   End If
End Sub



Private Sub txt_articulo_GotFocus()
   Me.frm_disponibles.Visible = False
End Sub

Private Sub txt_busqueda_KeyPress(KeyAscii As Integer)
Dim list_item As ListItem
   If KeyAscii = 13 Then
      rs.Open "select * from VW_encabezado_pedidos where inte_ped_numero = " + txt_busqueda, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_tipo_pedido = ""
         txt_nombre_agente = ""
         txt_nombre_titular = ""
         txt_nombre_establecimiento = ""
         txt_nombre_cliente = ""
         var_lista_precios = ""
         var_canal_venta = ""
         txt_tipo_pedido = ""
         txt_tipo_pedido.Enabled = False
         txt_agente = ""
         txt_agente.Enabled = False
         txt_titular = ""
         txt_titular.Enabled = False
         txt_codigo = ""
         txt_Articulo = ""
         txt_cantidad = ""
         var_primera_vez = False
         lv_pedidos.ListItems.Clear
         txt_nombre_tipo_pedido = IIf(IsNull(rs!VCHA_tpe_NOMBRE), "", rs!VCHA_tpe_NOMBRE)
         txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
         txt_nombre_titular = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
         txt_nombre_establecimiento = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
         txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         
         var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
         If IsNull(rs!char_tpe_tipo_pedido_id) Then
            txt_tipo_pedido = ""
         Else
            txt_tipo_pedido = rs!char_tpe_tipo_pedido_id
         End If
         If IsNull(rs!vcha_cli_clave_id) Then
            txt_clave_cliente = ""
         Else
            txt_clave_cliente = rs!vcha_cli_clave_id
         End If
         txt_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
         If IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id) Then
            txt_estbalecimiento = ""
         Else
            txt_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
         End If
         If IsNull(rs!vcha_tit_titular_id) Then
            txt_titular = ""
         Else
            txt_titular = rs!vcha_tit_titular_id
         End If
         If IsNull(rs!floa_ped_descuento_1) Then
            var_descuento_1 = 0
            txt_descuento1 = Format(0, "##0.000")
         Else
            var_descuento_1 = rs!floa_ped_descuento_1
            txt_descuento1 = Format(var_descuento_1, "##0.000")
         End If
         If IsNull(rs!FLOA_PED_DESCUENTO_2) Then
            var_descuento_2 = 0
            txt_descuento2 = Format(0, "##0.000")
         Else
            var_descuento_2 = rs!FLOA_PED_DESCUENTO_2
            txt_descuento2 = Format(var_descuento_2, "##0.000")
         End If
         txt_numero = rs!inte_ped_numero
         txt_fecha = rs!dtim_ped_fecha
         var_estatus = Trim(rs!char_ped_estatus)
         rs.Close
         rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
         var_lista_precios = IIf(IsNull(rs!vcha_LIS_LISTA_iD), "", rs!vcha_LIS_LISTA_iD)
         var_canal_venta = IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id)
         rs.Close
         frm_busqueda.Visible = False
         txt_codigo.Enabled = True
         rs.Open "select * from VW_PEDIDOS where inte_ped_NUMERO = " + txt_busqueda, cnn, adOpenDynamic, adLockOptimistic
         var_suma_cantidad = 0
         var_suma_importe = 0
         lbl_total = "0"
         While Not rs.EOF
            Set list_item = lv_pedidos.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
            list_item.SubItems(2) = IIf(IsNull(rs!FLOA_PED_PRECIO), Format(0, "###,###,##0.00"), Format(rs!FLOA_PED_PRECIO, "###,###,##0.00"))
            list_item.SubItems(3) = IIf(IsNull(rs!FLOA_PED_CANTIDAD), Format(0, "###,###,##0.00"), Format(rs!FLOA_PED_CANTIDAD, "###,###,##0.00"))
            list_item.SubItems(4) = Format(list_item.SubItems(2) * list_item.SubItems(3), "###,###,##0.00")
            list_item.SubItems(5) = IIf(IsNull(rs!char_ped_tipo), "P", rs!char_ped_tipo)
            var_suma_cantidad = var_suma_cantidad + list_item.SubItems(3)
            var_suma_importe = var_suma_importe + list_item.SubItems(4)
            lbl_total = CDbl(lbl_total) + IIf(IsNull(rs!FLOA_PED_CANTIDAD), 0, rs!FLOA_PED_CANTIDAD)
            rs.MoveNext
         Wend
         rs.Close
         txt_suma_cantidad = Format(var_suma_cantidad, "###,###,##0.00")
         txt_suma_importe = Format(var_suma_importe, "###,###,##0.00")
         txt_tipo_pedido.Enabled = False
         txt_titular.Enabled = False
         txt_establecimiento.Enabled = False
         txt_clave_cliente.Enabled = False
         txt_agente.Enabled = False
         If var_estatus <> "" Then
            txt_codigo.Enabled = False
            txt_cantidad.Enabled = False
            txt_foco.Enabled = False
            lv_pedidos.SetFocus
         Else
            txt_codigo.Enabled = True
            txt_codigo.SetFocus
         End If
      Else
         rs.Close
         MsgBox "El número de pedido no existe", vbOKOnly, "ATENCION"
      End If
   End If
   If Me.lv_pedidos.ListItems.Count > 12 Then
      lv_pedidos.ColumnHeaders(2).Width = 5900
   Else
      lv_pedidos.ColumnHeaders(2).Width = 6199.93
   End If
   
   If KeyAscii = 27 Then
      frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_cantidad_GotFocus()
   Me.frm_disponibles.Visible = False
   txt_cantidad = ""
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      var_cantidad_pedida = Val(txt_cantidad)
      If var_cantidad_pedida > 0 Then
         txt_foco.Enabled = True
         txt_foco.SetFocus
      Else
         MsgBox "Cantidad Incorrecta", vbOKOnly, "ATENCION"
         txt_cantidad.SetFocus
      End If
   End If
End Sub

Private Sub txt_clave_cliente_GotFocus()
   Me.frm_disponibles.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clave_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_cli_clave_id,vcha_cli_nombre from vw_pedidos_2 where VCHA_TCL_TIPO_CLIENTE_ID = '" + var_tipo_cliente + "' and vcha_tit_titular_id = '" + txt_titular + "' and vcha_esb_establecimiento_id = '" + txt_establecimiento + "' and vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_age_agente_id = '" + Me.txt_agente + "' order by vcha_cli_nombre ", cnn, adOpenDynamic, adLockBatchOptimistic
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

Private Sub txt_clave_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      txt_codigo.Enabled = True
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_clave_cliente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_clave_cliente) <> "" Then
      txt_clave_cliente = UCase(txt_clave_cliente)
      rs.Open "select floa_gac_descuento_1, floa_gac_descuento_2, inte_pla_dias, inte_tpe_dias_caducidad, floa_gac_descuento_3, vcha_cli_nombre, vcha_lis_lista_id, vcha_mon_moneda_id, vcha_can_canal_venta_id, inte_tpe_resurtible from vw_pedidos_2 where VCHA_TCL_TIPO_CLIENTE_ID = '" + var_tipo_cliente + "' and vcha_tit_titular_id = '" + txt_titular + "' and vcha_esb_establecimiento_id = '" + txt_establecimiento + "'  and vcha_cli_clave_id = '" + txt_clave_cliente + "' order by vcha_cli_nombre ", cnn, adOpenDynamic, adLockBatchOptimistic
      If Not rs.EOF Then
         txt_nombre_cliente = rs!VCHA_CLI_NOMBRE
         var_lista_precios = IIf(IsNull(rs!vcha_LIS_LISTA_iD), "", rs!vcha_LIS_LISTA_iD)
         var_canal_venta = IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id)
         var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
         var_canal_venta = IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id)
         var_resurtible = IIf(IsNull(rs!inte_tpe_resurtible), 0, rs!inte_tpe_resurtible)
         If IsNull(rs(0).Value) Then
            var_descuento_1 = 0
            txt_descuento1 = Format(var_descuento_1, "##0.000")
         Else
            var_descuento_1 = rs(0).Value
            txt_descuento1 = Format(rs(0).Value, "##0.000")
         End If
         If IsNull(rs(1).Value) Then
            var_descuento_2 = 0
            txt_descuento2 = Format(var_descuento_2, "##0.000")
         Else
            var_descuento_2 = rs(1).Value
            txt_descuento2 = Format(var_descuento_2, "##0.000")
         End If
         If IsNull(rs(2).Value) Then
            txt_plazo = 0
            var_dias_condiciones = 0
         Else
            txt_plazo = rs(2).Value
            var_dias_condiciones = rs(2).Value
         End If
         If IsNull(rs(3).Value) Then
            var_dias_caducidad = 0
         Else
            var_dias_caducidad = rs(3).Value
         End If
         rs.Close
         txt_clave_cliente.Enabled = False
         txt_codigo.Enabled = True
      Else
         rs.Close
         txt_clave_cliente = ""
         txt_nombre_cliente = ""
         MsgBox "Cliente Incorrecto", vbOKOnly, "ATENCION"
      End If
      If var_lista_precios = "" Then
         MsgBox "El cliente no tiene una lista de precios asociada", vbOKOnly, "ATENCION"
         txt_codigo.Enabled = False
         txt_clave_cliente.Enabled = True
      End If
   End If
End Sub

Private Sub txt_codigo_Change()
   Me.txt_existen = ""
   Me.txt_apartados = ""
   Me.txt_disponible = ""
End Sub

Private Sub txt_codigo_GotFocus()
   Me.txt_existen = ""
   Me.txt_apartados = ""
   Me.txt_disponible = ""
   Me.frm_disponibles.Visible = False
   If var_origen_codigo <> 1 Then
      txt_codigo = ""
      txt_Articulo = ""
      txt_cantidad = ""
   End If
End Sub

Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      'rs.Open "select vcha_art_nombre_español from tb_articulos order by vcha_art_nombre_español", cnn, adOpenDynamic, adLockOptimistic
      'While Not rs.EOF
      '   lst_articulos.AddItem rs(0).Value
      '   rs.MoveNext
      'Wend
      'rs.Close
      'frm_articulos.Visible = True
      Me.txt_codigo = ""
      var_origen_codigo = 1
      'lst_articulos.SetFocus
      Me.txt_nombre_articulo = ""
      Me.lv_disponibles.ListItems.Clear
      frm_disponibles.Visible = True
      Me.txt_nombre_articulo.SetFocus
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_codigo) <> "" Then
         If var_empresa = "16" Then
            
         End If
         txt_cantidad.Enabled = True
         txt_cantidad.SetFocus
      End If
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   Dim var_posible As Boolean
   
   If Trim(txt_codigo) <> "" Then
      cnn.CommandTimeout = 360
      var_posible = False
      rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = True
         txt_Articulo = rs!vcha_Art_nombre_español
         rs.Close
      Else
         rs.Close
         rs.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_posible = True
               txt_codigo = rs!VCHA_ART_ARTICULO_ID
               txt_Articulo = rsaux!vcha_Art_nombre_español
               rsaux.Close
               rs.Close
            Else
               var_posible = False
               rsaux.Close
               rs.Close
            End If
         Else
            rs.Close
         End If
      End If
      If var_posible = True Then
         If var_origen_codigo = 0 Then
            txt_cantidad = Format(0, "###0.00")
         Else
            var_origen_codigo = 0
            txt_cantidad = Format(0, "###0.00")
         End If
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         txt_codigo.SetFocus
      End If
      rs.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + Me.txt_codigo + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_existen = Format(IIf(IsNull(rs!floa_Exi_Cantidad), 0, rs!floa_Exi_Cantidad), "###,###,##0.0000")
         Me.txt_apartados = Format(IIf(IsNull(rs!FLOA_EXI_CANTIDAD_APARTADA), 0, rs!FLOA_EXI_CANTIDAD_APARTADA), "###,###,##0.0000")
         Me.txt_disponible = Format(IIf(IsNull(rs!floa_Exi_Cantidad_disponible), 0, rs!floa_Exi_Cantidad_disponible), "###,###,##0.0000")
      Else
         Me.txt_existen = ""
         Me.txt_apartados = ""
         Me.txt_disponible = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_eliminar_KeyPress(KeyAscii As Integer)
   Set TB_DETALLE_PEDIDOS_M = New TB_DETALLE_PEDIDOS_M
   Dim var_cantidad_eliminar As Variant
   Dim var_precio As Variant
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(lv_pedidos.selectedItem.SubItems(5)) = "P" Then
         var_cantidad_eliminar = Val(txt_eliminar)
         convierte_numero (lv_pedidos.selectedItem.SubItems(3))
         var_precio = lv_pedidos.selectedItem.SubItems(2)
         var_cantidad_total = Val(numero_devuelto)
         If var_cantidad_eliminar <= var_cantidad_total Then
            lbl_total = CDbl(lbl_total) - var_cantidad_eliminar
            convierte_numero (lv_pedidos.selectedItem.SubItems(3))
            var_anterior_cantidad = Val(numero_devuelto)
            convierte_numero (lv_pedidos.selectedItem.SubItems(2))
            var_anterior_importe = Val(numero_devuelto) * var_anterior_cantidad
            rs.Open "update tb_detalle_pedidos set floa_ped_cantidad = floa_ped_cantidad - " + CStr(var_cantidad_eliminar) + " where inte_ped_numero = " + txt_numero + " and vcha_art_articulo_id = '" + Trim(lv_pedidos.selectedItem) + "' and char_ped_tipo = 'P'", cnn, adOpenDynamic, adLockOptimistic
            'ok = TB_DETALLE_PEDIDOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_numero, lv_pedidos.SelectedItem, 0, 0 - var_cantidad_eliminar, 0)
            lv_pedidos.selectedItem.SubItems(3) = Format(var_anterior_cantidad - var_cantidad_eliminar, "###,###,##0.00")
            var_nueva_cantidad = var_anterior_cantidad - var_cantidad_eliminar
            convierte_numero (lv_pedidos.selectedItem.SubItems(2))
            var_nuevo_importe = Val(numero_devuelto) * var_nueva_cantidad
            convierte_numero (lv_pedidos.selectedItem.SubItems(2))
            var_precio = Val(numero_devuelto)
            lv_pedidos.selectedItem.SubItems(4) = Format(var_precio * var_nueva_cantidad, "###,###,##0.00")
            var_renglon = lv_pedidos.selectedItem.Index
            Call ilumina_grid
            var_anterior_cantidad = var_anterior_cantidad - var_nueva_cantidad
            var_anterior_importe = var_anterior_importe - var_nuevo_importe
            var_suma_cantidad = var_suma_cantidad - var_anterior_cantidad
            var_suma_importe = var_suma_importe - var_anterior_importe
            txt_suma_cantidad = Format(var_suma_cantidad, "###,###,##0.00")
            txt_suma_importe = Format(var_suma_importe, "###,###,##0.00")
         Else
            MsgBox "Imposible eliminar esta cantidad", vbOKOnly, "ATENCION"
         End If
         frm_eliminar.Visible = False
      Else
         MsgBox "No puede eliminar artículos asignacos por el sistema", vbOKOnly, "ATENCION"
         frm_eliminar.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      frm_eliminar.Visible = False
   End If
End Sub

Private Sub txt_eliminar_LostFocus()
   frm_eliminar.Visible = False
End Sub

Private Sub txt_establecimiento_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_establecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_esb_establecimiento_id,vcha_esb_nombre from vw_pedidos_2 where VCHA_TCL_TIPO_CLIENTE_ID = '" + var_tipo_cliente + "' and vcha_tit_titular_id = '" + txt_titular + "' and vcha_age_agente_id = '" + Me.txt_agente + "' order by vcha_esb_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ESB_ESTABLECIMIENTO_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Establecimientos"
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      var_tipo_lista = 4
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      txt_clave_cliente.Enabled = True
      txt_clave_cliente.SetFocus
   End If
End Sub

Private Sub txt_establecimiento_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_establecimiento) <> "" Then
      txt_establecimiento = UCase(txt_establecimiento)
      rs.Open "select * from vw_pedidos_2 where vcha_tit_titular_id = '" + txt_titular + "' and VCHA_TCL_TIPO_CLIENTE_ID = '" + var_tipo_cliente + "' AND VCHA_ESB_ESTABLECIMIENTO_ID = '" + txt_establecimiento + "' and vcha_age_agente_id = '" + Me.txt_agente + "' ", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_clave_cliente.Enabled = True
         txt_nombre_establecimiento = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
         rs.Close
         txt_establecimiento.Enabled = False
      Else
         rs.Close
         txt_nombre_establecimiento = ""
         txt_establecimiento = ""
         txt_clave_cliente.Enabled = False
         MsgBox "Establecimiento Incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   Set TB_ENC_PEDIDOS_I = New TB_ENC_PEDIDOS_I
   Set TB_DETALLE_PEDIDOS_I = New TB_DETALLE_PEDIDOS_I
   Set TB_DETALLE_PEDIDOS_M = New TB_DETALLE_PEDIDOS_M
   Dim var_i As Integer
   Dim var_n As Integer
   Dim var_precio_anterior As Variant
   Dim list_item As ListItem
   Dim var_catalogo As String
   Dim var_numero_dias As Double
   Dim var_otorga_oferta As Boolean
   Dim var_posible As Boolean
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   Dim agrupador_catalogo As String
   Dim var_precio_externo As Double
   Dim var_catalogo_EFASA As String
   var_origen_codigo = 0
   If var_lista_precios <> "" Then
      cnn.CommandTimeout = 360
      If Trim(var_clave_moneda) <> "" Then
         If Trim(txt_codigo) <> "" Then
            If rsaux4.State = 1 Then
               rsaux4.Close
            End If
            rsaux4.Open "select * from tb_detalle_lista_precios where VCHA_LIS_LISTA_PRECIOS_ID = '" + var_lista_precios + "' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux4.EOF Then
               var_precio_pedido = IIf(IsNull(rsaux4!floa_dli_precio), 0, rsaux4!floa_dli_precio)
               If var_precio_pedido > 0 Then
                  'cnn.BeginTrans
                  var_promocion_1 = 0
                  var_promocion_2 = 0
                  var_precio_pedido = rsaux4!floa_dli_precio
                  
                  var_promociones_ya_no = 0
                  If var_promociones_ya_no = 1 Then
                     rs.Open "select * from vw_lista_precios_clientes where vcha_cli_clave_id = '" + txt_clave_cliente + "' and vcha_lis_lista_id = '" + var_lista_precios + "' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_promocion_1 = 0
                        var_promocion_2 = 0
                        var_precio_pedido = rs!floa_dli_precio
                        var_catalogo = rs!vcha_cat_catalogo_id
                        var_otorga_oferta = False
                        If Not IsNull(rs!dtim_vig_fecha_fin) Then
                           var_numero_dias = Date - rs!dtim_vig_fecha_fin
                           var_otorga_oferta = True
                        Else
                           var_otorga_oferta = False
                        End If
                     End If
                     rs.Close
                  
                  
                     rs.Open "select * from vw_descuentos_promociones_vigentes where vcha_can_canal_venta_id = '" + var_canal_venta + "' and vcha_art_Articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_promocion_1 = IIf(IsNull(rs!floa_dpr_desCuento), 0, rs!floa_dpr_desCuento)
                        var_precio_pedido = var_precio_pedido - (var_precio_pedido * (IIf(IsNull(rs!floa_dpr_desCuento), 0, rs!floa_dpr_desCuento) / 100))
                        rs.Close
                     Else
                        rs.Close
                        If var_otorga_oferta = True Then
                           rs.Open "select * from tb_descuentos_catalogos where vcha_can_canal_venta_id = '" + var_canal_venta + "' and inte_des_limite_inferior <= " + Str(var_numero_dias) + " and inte_des_limite_superior >= " + Str(var_numero_dias), cnn, adOpenDynamic, adLockOptimistic
                           If Not rs.EOF Then
                              var_promocion_2 = IIf(IsNull(rs!FLOA_DES_DESCUENTO), 0, rs!FLOA_DES_DESCUENTO)
                              var_precio_pedido = var_precio_pedido - (var_precio_pedido * (rs!FLOA_DES_DESCUENTO / 100))
                           End If
                           rs.Close
                        End If
                     End If
                  End If
                  
                  
                  If var_primera_vez = True Then
                     txt_numero = maximo_pedido
                     var_primera_vez = False
                     ok = TB_ENC_PEDIDOS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_tipo_pedido, maximo_pedido, 0, Date, Date, txt_agente, txt_titular, txt_clave_cliente, txt_establecimiento, var_resurtible, 0, "", var_descuento_1, var_descuento_2, var_descuento_3, var_dias_condiciones, var_dias_caducidad, var_clave_usuario_global, fun_NombrePc, Date, var_clave_moneda, 0)
                     txt_numero = maximo_pedido
                     rs.Open "select * from VW_PEDIDOS where inte_ped_NUMERO = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                     var_suma_cantidad = 0
                     var_suma_importe = 0
                     While Not rs.EOF
                           Set list_item = lv_pedidos.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
                           list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
                           list_item.SubItems(2) = IIf(IsNull(rs!FLOA_PED_PRECIO), Format(0, "###,###,##0.00"), Format(rs!FLOA_PED_PRECIO, "###,###,##0.00"))
                           list_item.SubItems(3) = IIf(IsNull(rs!FLOA_PED_CANTIDAD), Format(0, "###,###,##0.00"), Format(rs!FLOA_PED_CANTIDAD, "###,###,##0.00"))
                           list_item.SubItems(4) = Format(list_item.SubItems(2) * list_item.SubItems(3), "###,###,##0.00")
                           list_item.SubItems(5) = IIf(IsNull(rs!char_ped_tipo), "P", rs!char_ped_tipo)
                           var_renglon = lv_pedidos.ListItems.Count
                           Call ilumina_grid
                           var_suma_cantidad = var_suma_cantidad + list_item.SubItems(3)
                           var_suma_importe = var_suma_importe + list_item.SubItems(4)
                           rs.MoveNext
                     Wend
                     rs.Close
                     txt_suma_cantidad = Format(var_suma_cantidad, "###,###,##0.00")
                     txt_suma_importe = Format(var_suma_importe, "###,###,##0.00")
                     txt_tipo_pedido.Enabled = False
                     txt_titular.Enabled = False
                     txt_establecimiento.Enabled = False
                     txt_clave_cliente.Enabled = False
                     txt_agente.Enabled = False
                  End If
                  rsaux.Open "select * from tb_detalle_pedidos where INTE_PED_NUMERO = " + txt_numero + " and VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND CHAR_PED_TIPO = 'P'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     rsaux.Close
                     valor = txt_codigo
                     rs.Open "update tb_detalle_pedidos set floa_ped_cantidad = floa_ped_cantidad + " + CStr(var_cantidad_pedida) + " where inte_ped_numero = " + txt_numero + " and vcha_art_articulo_id = '" + txt_codigo + "' AND CHAR_PED_TIPO = 'P'", cnn, adOpenDynamic, adLockOptimistic
                     lbl_total = CDbl(lbl_total) + var_cantidad_pedida
                     var_n = lv_pedidos.ListItems.Count
                     var_encontro = 0
                     var_i = 1
                     While (var_i <= var_n)
                         lv_pedidos.ListItems.Item(var_i).Selected = True
                         valor = Trim(lv_pedidos.selectedItem)
                         If txt_codigo = valor Then
                            'If lv_pedidos.selectedItem.SubItems(5) = "P" Then
                            '   var_precio_anterior = (lv_pedidos.selectedItem.SubItems(2) * 1)
                            '   If var_precio_anterior <> var_precio_pedido Then
                            '      var_encontro = 0
                            '   Else
                            var_encontro = 1
                            var_i = var_n
                            '   End If
                            'End If
                         End If
                         var_i = var_i + 1
                     Wend
                     bandera_suma = True
                     convierte_numero (lv_pedidos.selectedItem.SubItems(3))
                     var_cantidad_anterior = Val(numero_devuelto)
                     lv_pedidos.selectedItem.SubItems(3) = Format(var_cantidad_anterior + var_cantidad_pedida, "###,###,##0.00")
                     lv_pedidos.selectedItem.SubItems(4) = Format((var_cantidad_anterior + var_cantidad_pedida) * var_precio_pedido, "###,###,##0.00")
                     var_renglon = lv_pedidos.selectedItem.Index
                     Call ilumina_grid
                     var_suma_cantidad = var_suma_cantidad + var_cantidad_pedida
                     var_suma_importe = var_suma_importe + (var_cantidad_pedida * var_precio_pedido)
                  Else
                     rsaux.Close
                     x = 0
                     If x = 0 Then
                        var_empresa_cliente = ""
                        rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_empresa_cliente = IIf(IsNull(rs!VCHA_EMP_EMPRESA_ID), "", rs!VCHA_EMP_EMPRESA_ID)
                        End If
                        rs.Close
                        If var_empresa_cliente = "03" Then
                           rs.Open "SELECT * FROM VW_catalogos_efasa where vcha_Art_articulo_id = '" + txt_codigo + "'"
                           If Not rs.EOF Then
                              var_precio_pedido = var_precio_pedido / (1 - (var_descuento_1 / 100))
                              var_precio_pedido = var_precio_pedido / (1 - (var_descuento_2 / 100))
                           End If
                           rs.Close
                        End If
                     End If

                     
                     
                     ok = TB_DETALLE_PEDIDOS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_numero, txt_codigo, var_precio_pedido, var_cantidad_pedida, 0, var_promocion_1, var_promocion_2, "P")
                     Set list_item = lv_pedidos.ListItems.Add(, , txt_codigo)
                     list_item.SubItems(1) = Trim(txt_Articulo)
                     list_item.SubItems(2) = Format(var_precio_pedido, "###,###,##0.00")
                     list_item.SubItems(3) = Format(var_cantidad_pedida, "###,###,##0.00")
                     list_item.SubItems(4) = Format(var_precio_pedido * var_cantidad_pedida, "###,###,##0.00")
                     list_item.SubItems(5) = "P"
                     var_renglon = lv_pedidos.ListItems.Count
                     Call ilumina_grid
                     var_suma_cantidad = var_suma_cantidad + var_cantidad_pedida
                     var_suma_importe = var_suma_importe + (var_cantidad_pedida * var_precio_pedido)
                     lbl_total = CDbl(lbl_total) + var_cantidad_pedida
                  End If
                  txt_suma_importe = Format(var_suma_importe, "###,###,##0.00")
                  txt_suma_cantidad = Format(var_suma_cantidad, "###,###,##0.00")
                  'cnn.CommitTrans
               Else
                  MsgBox "El precio del artículo esta en ceros", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Este artículo no se encuentra en la lista de precios asignada al cliente", vbOKOnly, "ATENCION"
            End If
            rsaux4.Close
         Else
            MsgBox "Código Incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El cliente no tiene una moneda asociada", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "El cliente no tiene una lista de precios asociada", vbOKOnly, "ATENCION"
   End If
   txt_codigo.SetFocus
   txt_foco.Enabled = False
End Sub



Private Sub txt_nombre_agente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_age_agente_id, vcha_age_nombre from vw_pedidos_2 where VCHA_TCL_TIPO_CLIENTE_ID = '" + var_tipo_cliente + "' order by vcha_age_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Agentes"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_agente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_articulo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.txt_codigo.SetFocus
      Me.frm_disponibles.Visible = False
   End If
   If KeyAscii = 39 Or KeyAscii = 61 Or KeyAscii = 44 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(Me.txt_nombre_articulo) <> "" Then
         var_cadena = ""
         var_like_1 = ""
         var_like_2 = ""
         var_like_3 = ""
         var_like_4 = ""
         var_like_5 = ""
         var_like_6 = ""
         var_like_7 = ""
         var_j = 1
         For var_i = 1 To Len(Me.txt_nombre_articulo)
             If Mid(Me.txt_nombre_articulo, var_i, 1) <> " " Then
                If var_j = 1 Then
                   var_like_1 = var_like_1 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 2 Then
                   var_like_2 = var_like_2 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 3 Then
                   var_like_3 = var_like_3 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 4 Then
                   var_like_4 = var_like_4 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 5 Then
                   var_like_5 = var_like_5 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 6 Then
                   var_like_6 = var_like_6 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j >= 7 Then
                   var_like_7 = var_like_7 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
             Else
                var_j = var_j + 1
             End If
         Next var_i
      End If
      If Trim(var_like_1) <> "" Then
         var_cadena = " vcha_art_nombre_Español like '%" + var_like_1 + "%'"
      End If
      If Trim(var_like_2) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_2 + "%'"
      End If
      If Trim(var_like_3) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_3 + "%'"
      End If
      If Trim(var_like_4) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_4 + "%'"
      End If
      If Trim(var_like_5) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_5 + "%'"
      End If
      If Trim(var_like_6) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_6 + "%'"
      End If
      If Trim(var_like_7) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_7 + "%'"
      End If
      Me.lv_disponibles.ListItems.Clear
      If Trim(var_cadena) <> "" Then
         var_cadena = "SELECT * FROM VW_DISPONIBLE WHERE " + var_cadena
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
            Set list_item = lv_disponibles.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
            list_item.SubItems(2) = Format(Round(IIf(IsNull(rs!floa_Exi_Cantidad_disponible), 0, rs!floa_Exi_Cantidad_disponible), 4), "###,###,##0.0000")
            If Mid(rs!VCHA_ART_ARTICULO_ID, 11, 1) Then
               list_item.ForeColor = &HFF&
               list_item.ListSubItems(1).ForeColor = &HFF&
               list_item.ListSubItems(2).ForeColor = &HFF&
            End If
            rs.MoveNext
         Wend
         rs.Close
         If Me.lv_disponibles.ListItems.Count > 0 Then
            Me.lv_disponibles.SetFocus
         End If
         If lv_disponibles.ListItems.Count > 11 Then
            lv_disponibles.ColumnHeaders(3).Width = 1200.18
         Else
            lv_disponibles.ColumnHeaders(3).Width = 1400.18
         End If
      End If
   End If
End Sub

Private Sub txt_nombre_cliente_GotFocus()
   Me.frm_disponibles.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_cli_clave_id,vcha_cli_nombre from vw_pedidos_2 where VCHA_TCL_TIPO_CLIENTE_ID = '" + var_tipo_cliente + "' and vcha_tit_titular_id = '" + txt_titular + "' and vcha_esb_establecimiento_id = '" + txt_establecimiento + "'  and vcha_age_agente_id = '" + Me.txt_agente + "' order by vcha_cli_nombre ", cnn, adOpenDynamic, adLockBatchOptimistic
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

Private Sub txt_nombre_cliente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_establecimiento_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_establecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_esb_establecimiento_id,vcha_esb_nombre from vw_pedidos_2 where VCHA_TCL_TIPO_CLIENTE_ID = '" + var_tipo_cliente + "' and vcha_tit_titular_id = '" + txt_titular + "' and vcha_age_agente_id = '" + Me.txt_agente + "' order by vcha_esb_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ESB_ESTABLECIMIENTO_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Establecimientos"
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      var_tipo_lista = 4
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_establecimiento_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_tipo_pedido_GotFocus()
   Me.frm_disponibles.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_tipo_pedido_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct char_tpe_tipo_pedido_id,vcha_tpe_nombre from vw_pedidos_2 order by VCHA_TPE_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!char_tpe_tipo_pedido_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_tpe_NOMBRE), "", rs!VCHA_tpe_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Tipo Pedidos"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_tipo_pedido_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_titular_GotFocus()
   Me.frm_disponibles.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_titular_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_tit_titular_id,vcha_tit_nombre from vw_pedidos_2 where VCHA_TCL_TIPO_CLIENTE_ID = '" + var_tipo_cliente + "' and vcha_age_agente_id = '" + txt_agente + "' order by vcha_tit_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tit_titular_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Titulares"
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      var_tipo_lista = 3
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_titular_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_tipo_pedido_GotFocus()
   Me.frm_disponibles.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_tipo_pedido_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct char_tpe_tipo_pedido_id,vcha_tpe_nombre from vw_pedidos_2 where char_tpe_tipo_pedido_id is not null order by VCHA_TPE_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!char_tpe_tipo_pedido_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_tpe_NOMBRE), "", rs!VCHA_tpe_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TIPO PEDIDOS"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_tipo_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      txt_agente.Enabled = True
      txt_agente.SetFocus
   End If
End Sub

Private Sub txt_tipo_pedido_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_tipo_pedido) <> "" Then
      txt_tipo_pedido = UCase(txt_tipo_pedido)
      rs.Open "select * from vw_clientes where char_tpe_tipo_pedido_id = '" + txt_tipo_pedido + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_tipo_cliente = rs!VCHA_TCL_TIPO_CLIENTE_ID
         txt_nombre_tipo_pedido = rs!VCHA_tpe_NOMBRE
         rs.Close
         txt_agente.Enabled = True
         txt_tipo_pedido.Enabled = False
      Else
         rs.Close
         MsgBox "Tipo de pedido incorrecto", vbOKOnly, "ATENCION"
         txt_tipo_pedido = ""
         txt_nombre_tipo_pedido = ""
         txt_agente.Enabled = False
      End If
   End If
End Sub

Private Sub txt_titular_GotFocus()
   Me.frm_disponibles.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_titular_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_tit_titular_id,vcha_tit_nombre from vw_pedidos_2 where VCHA_TCL_TIPO_CLIENTE_ID = '" + var_tipo_cliente + "' and vcha_age_agente_id = '" + txt_agente + "' order by vcha_tit_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tit_titular_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Titulares"
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      var_tipo_lista = 3
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_titular_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If txt_tipo_pedido = "V" Then
         txt_codigo.Enabled = True
         txt_codigo.SetFocus
      Else
         txt_establecimiento.Enabled = True
         txt_establecimiento.SetFocus
      End If
   End If
End Sub

Private Sub txt_titular_LostFocus()
   Dim var_posible_venta As Boolean
   Dim var_negado As Integer
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_titular) <> "" Then
      txt_titular = UCase(txt_titular)
      rs.Open "select * from vw_pedidos_2 where vcha_tit_titular_id = '" + txt_titular + "' and VCHA_TCL_TIPO_CLIENTE_ID = '" + var_tipo_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_negado = 0
         var_posible_venta = True
         If var_posible_limite_credito = 1 Then
            var_cadena = "SELECT     dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_TITULARES.VCHA_TIT_TITULAR_ID, dbo.TB_TITULARES.FLOA_TIT_LIMITE_CREDITO FROM dbo.TB_CLIENTES INNER JOIN dbo.TB_TITULARES ON dbo.TB_CLIENTES.VCHA_TIT_TITULAR_ID = dbo.TB_TITULARES.VCHA_TIT_TITULAR_ID WHERE     (dbo.TB_CLIENTES.vcha_tit_titular_id = '" + Me.txt_titular + "')"
            rsaux10.Open var_cadena, cnn_distribucion, adOpenDynamic, adLockOptimistic
            If Not rsaux10.EOF Then
               If rsaux10!vcha_tit_titular_id = "T000001038" Then
                  var_posible_venta = True
               Else
                  var_limite_credito = IIf(IsNull(rsaux10!floa_tit_limite_credito), 0, rsaux10!floa_tit_limite_credito)
                  
                  'var_cadena = "SELECT     SUM(dbo.TB_SALDOS.FLOA_SAL_IMPORTE) AS importe, dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO, dbo.TB_SALDOS.FLOA_SAL_IMPORTE , dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID FROM dbo.TB_ENCABEZADO_CARTERA INNER JOIN dbo.TB_SALDOS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = dbo.TB_SALDOS.VCHA_SER_SERIE_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO AND dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO = dbo.TB_SALDOS.INTE_CAR_NUMERO WHERE (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C' OR dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL) AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = 'FA') and (dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID = '" + Me.txt_titular + "') AND "
                  'var_cadena = var_cadena + "(dbo.TB_SALDOS.FLOA_SAL_IMPORTE > .50) GROUP BY dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO, dbo.TB_SALDOS.FLOA_SAL_IMPORTE, dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID"
                  
                  var_cadena = "SELECT SUM(dbo.TB_SALDOS.FLOA_SAL_IMPORTE) AS importe  FROM dbo.TB_ENCABEZADO_CARTERA INNER JOIN dbo.TB_SALDOS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = dbo.TB_SALDOS.VCHA_SER_SERIE_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO AND dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO = dbo.TB_SALDOS.INTE_CAR_NUMERO WHERE (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C' OR dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL) AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = 'FA') and (dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID = '" + Me.txt_titular + "') AND (dbo.TB_SALDOS.FLOA_SAL_IMPORTE > .50) "
                  If rsaux9.State = 1 Then
                     rsaux9.Close
                  End If
                  'MsgBox cnn_distribucion.ConnectionString
                  rsaux9.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux9.EOF Then
                     var_saldo_titular = 0
                     While Not rsaux9.EOF
                           var_saldo_titular = var_saldo_titular + IIf(IsNull(rsaux9!Importe), 0, rsaux9!Importe)
                           rsaux9.MoveNext
                     Wend
                     'MsgBox CStr(var_saldo_titular)
                     If var_saldo_titular >= var_limite_credito Then
                        var_posible_venta = False
                        var_negado = 1
                     End If
                  Else
                     var_saldo_titular = 0
                  End If
                  rsaux9.Close
                  'var_cadena = " SELECT SUM(dbo.TB_SALDOS.FLOA_SAL_IMPORTE) AS importe, dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID FROM  dbo.TB_ENCABEZADO_CARTERA INNER JOIN dbo.TB_SALDOS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = dbo.TB_SALDOS.VCHA_SER_SERIE_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO AND dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID = dbo.TB_SALDOS.VCHA_CLI_CLAVE_ID AND dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO = dbo.TB_SALDOS.INTE_CAR_NUMERO WHERE (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C' OR dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL) AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_SALDOS.FLOA_SAL_IMPORTE > 0.5) AND (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA + dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO < GETDATE()) "
                  'var_cadena = var_cadena + " GROUP BY dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID HAVING      (dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID = '" + Me.txt_titular + "') AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')"
                  'If rsaux8.State = 1 Then
                  '   rsaux8.Close
                  'End If
                  'rsaux8.Open var_cadena, cnn_distribucion, adOpenDynamic, adLockOptimistic
                  'If Not rsaux8.EOF Then
                  '   MsgBox CStr(rsaux8!Importe)
                  '   var_posible_venta = False
                  '   var_negado = 2
                  'End If
                  'rsaux8.Close
               End If

               
            Else
               var_posible_venta = False
            End If
            rsaux10.Close
         Else
            var_posible_venta = True
         End If
         var_posible_venta = True
         If var_posible_venta = True Then
            txt_nombre_titular = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
            If txt_tipo_pedido = "V" Then
               rs.Close
               rs.Open "select distinct floa_gac_descuento_1, floa_gac_descuento_2,inte_pla_dias,inte_tpe_dias_caducidad,floa_gac_descuento_3,vcha_esb_establecimiento_id,vcha_esb_nombre,vcha_cli_clave_id,vcha_cli_nombre,vcha_lis_lista_id, vcha_can_canal_venta_id, inte_tpe_resurtible, vcha_mon_moneda_id from vw_pedidos_2 where VCHA_TCL_TIPO_CLIENTE_ID = '" + var_tipo_cliente + "' and vcha_tit_titular_id = '" + txt_titular + "' order by vcha_esb_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
               If Not rs.EOF Then
                  var_lista_precios = IIf(IsNull(rs!vcha_LIS_LISTA_iD), "", rs!vcha_LIS_LISTA_iD)
                  txt_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
                  txt_clave_cliente = rs!vcha_cli_clave_id
                  txt_nombre_cliente = rs!VCHA_CLI_NOMBRE
                  txt_establecimiento.Enabled = False
                  txt_titular.Enabled = False
                  txt_clave_cliente.Enabled = False
                  
                  var_lista_precios = IIf(IsNull(rs!vcha_LIS_LISTA_iD), "", rs!vcha_LIS_LISTA_iD)
                  var_canal_venta = IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id)
                  var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                  var_canal_venta = IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id)
                  var_resurtible = IIf(IsNull(rs!inte_tpe_resurtible), 0, rs!inte_tpe_resurtible)
                  
                  If IsNull(rs(0).Value) Then
                     var_descuento_1 = 0
                     txt_descuento1 = Format(var_descuento_1, "##0.000")
                  Else
                     var_descuento_1 = rs(0).Value
                     txt_descuento1 = Format(rs(0).Value, "##0.000")
                  End If
                  If IsNull(rs(1).Value) Then
                     var_descuento_2 = 0
                      txt_descuento2 = Format(var_descuento_2, "##0.000")
                  Else
                     var_descuento_2 = rs(1).Value
                     txt_descuento2 = Format(var_descuento_2, "##0.000")
                  End If
                  If IsNull(rs(2).Value) Then
                     txt_plazo = 0
                     var_dias_condiciones = 0
                  Else
                     txt_plazo = rs(2).Value
                     var_dias_condiciones = rs(2).Value
                  End If
                  If IsNull(rs(3).Value) Then
                     var_dias_caducidad = 0
                  Else
                     var_dias_caducidad = rs(3).Value
                  End If
                  txt_codigo.Enabled = True
                  txt_codigo.SetFocus
               Else
                  MsgBox "El titular no tiene relacionado algun establecimiento o un cliente", vbOKOnly, "ATENCION"
               End If
               rs.Close
            Else
               txt_establecimiento.Enabled = True
               rs.Close
               txt_titular.Enabled = False
            End If
         Else
            If rs.State = 1 Then
               rs.Close
            End If
            Me.txt_titular = ""
            Me.txt_nombre_titular = ""
            If var_negado = 1 Then
               MsgBox "El crédito del cliente excede al limite y la venta no puede ser hecha", vbOKOnly, "ATENCION"
            End If
            If var_negado = 2 Then
               MsgBox "El cliente ya tiene facturas vencidas no se puede hacer la venta", vbOKOnly, "ATENCION"
            End If
         End If
      Else
         rs.Close
         txt_titular = ""
         txt_nombre_titular = ""
         txt_establecimiento.Enabled = False
         MsgBox "Titular Incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

