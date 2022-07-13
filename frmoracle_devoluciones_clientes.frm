VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_devoluciones_clientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Devoluciones de clientes"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_asignar_causa_devolucion 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Picture         =   "frmoracle_devoluciones_clientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "Asignar causa de devoluci�n"
      Top             =   720
      Width           =   330
   End
   Begin VB.Frame frm_pasar_todo 
      Height          =   1500
      Left            =   2160
      TabIndex        =   2
      Top             =   2805
      Width           =   2430
      Begin VB.TextBox txt_serie 
         Height          =   315
         Left            =   735
         TabIndex        =   7
         Top             =   990
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txt_numero 
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
         Height          =   525
         Left            =   315
         TabIndex        =   6
         Top             =   885
         Width           =   1920
      End
      Begin VB.Frame Frame5 
         Height          =   30
         Left            =   0
         TabIndex        =   5
         Top             =   765
         Width           =   2415
      End
      Begin VB.CommandButton cmd_aceptar_pedidos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         Picture         =   "frmoracle_devoluciones_clientes.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   420
         Width           =   330
      End
      Begin VB.CommandButton cmd_cancelar_pedidos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmoracle_devoluciones_clientes.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   420
         Width           =   330
      End
      Begin VB.Label Label5 
         BackColor       =   &H000000C0&
         Caption         =   " Factura a Devolver"
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   9
         Top             =   120
         Width           =   2355
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   1050
         Visible         =   0   'False
         Width           =   15
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   1875
      TabIndex        =   65
      Top             =   765
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   285
      Left            =   1455
      TabIndex        =   62
      Top             =   735
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt_clave_titular 
      Height          =   285
      Left            =   5490
      TabIndex        =   61
      Text            =   "Text1"
      Top             =   705
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   360
      TabIndex        =   10
      Top             =   1200
      Width           =   7305
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1980
         Left            =   45
         TabIndex        =   11
         Top             =   375
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   3493
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
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Anterior"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H000000C0&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   120
         Width           =   7230
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   6255
      TabIndex        =   56
      Top             =   2190
      Width           =   2040
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         TabIndex        =   58
         Top             =   420
         Width           =   1830
      End
      Begin VB.Label lbl_total 
         Alignment       =   2  'Center
         Caption         =   "12345619999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         TabIndex        =   57
         Top             =   795
         Width           =   1830
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3570
      Left            =   90
      TabIndex        =   44
      Top             =   3660
      Width           =   8235
      Begin VB.CommandButton cmd_movimiento_masivo 
         Caption         =   "Carga masiva"
         Height          =   450
         Left            =   105
         TabIndex        =   66
         Top             =   450
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.CommandButton cmd_factura 
         Caption         =   "Factura"
         Height          =   375
         Left            =   5280
         TabIndex        =   64
         Top             =   480
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.CommandButton cmd_devolucion 
         Caption         =   "Devoluci�n"
         Height          =   360
         Left            =   4290
         TabIndex        =   63
         Top             =   510
         Visible         =   0   'False
         Width           =   990
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
         Left            =   1545
         TabIndex        =   49
         Top             =   405
         Width           =   2640
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   1785
         TabIndex        =   46
         Top             =   1755
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   60
            TabIndex        =   47
            Top             =   375
            Width           =   2745
         End
         Begin VB.Label lbl_nombre_eliminar 
            BackColor       =   &H000000C0&
            Caption         =   "Cantidad a eliminar"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   0
            TabIndex        =   48
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
         Left            =   6240
         TabIndex        =   45
         Top             =   465
         Width           =   1890
      End
      Begin MSComctlLib.ListView lv_entradas 
         Height          =   2430
         Left            =   45
         TabIndex        =   50
         Top             =   1065
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   4286
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "C�digo"
            Object.Width           =   2478
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripci�n"
            Object.Width           =   9349
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   2328
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "localizador"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C�digo del Art�culo:"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   585
         Width           =   1395
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Lectura de Art�culos"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   53
         Top             =   120
         Width           =   8160
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   5535
         TabIndex        =   52
         Top             =   585
         Width           =   675
      End
      Begin VB.Label lbl_cancelado 
         Alignment       =   2  'Center
         Caption         =   "MOVIMIENTO CANCELADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   4320
         TabIndex        =   51
         Top             =   420
         Width           =   3765
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   30
      TabIndex        =   43
      Top             =   570
      Width           =   8250
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   9015
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   2910
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      Height          =   1080
      Index           =   0
      Left            =   6240
      TabIndex        =   20
      Top             =   1110
      Width           =   2055
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
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   435
         Width           =   1950
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   1
         Left            =   30
         TabIndex        =   22
         Top             =   120
         Width           =   1980
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7920
      Picture         =   "frmoracle_devoluciones_clientes.frx":0396
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Salir"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   735
      Picture         =   "frmoracle_devoluciones_clientes.frx":09D0
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmoracle_devoluciones_clientes.frx":0AD2
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Buscar Movimiento Alt + B"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmoracle_devoluciones_clientes.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Nuevo Movimiento Alt + N"
      Top             =   720
      Width           =   330
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   1575
      TabIndex        =   13
      Top             =   1080
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         MaxLength       =   10
         TabIndex        =   14
         Top             =   495
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   30
         TabIndex        =   15
         Top             =   120
         Width           =   3060
      End
   End
   Begin VB.CheckBox chk_factura 
      Caption         =   "Check1"
      Height          =   255
      Left            =   7275
      TabIndex        =   1
      Top             =   765
      Width           =   420
   End
   Begin VB.TextBox txt_movimiento 
      Height          =   345
      Left            =   7635
      TabIndex        =   0
      Top             =   30
      Width           =   480
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   30
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
            Picture         =   "frmoracle_devoluciones_clientes.frx":0CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_devoluciones_clientes.frx":15B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_devoluciones_clientes.frx":1E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_devoluciones_clientes.frx":2426
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_devoluciones_clientes.frx":2D02
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_devoluciones_clientes.frx":35DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_devoluciones_clientes.frx":3EB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_devoluciones_clientes.frx":3FC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_devoluciones_clientes.frx":40DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_devoluciones_clientes.frx":41EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_devoluciones_clientes.frx":42FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_devoluciones_clientes.frx":4410
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   30
      TabIndex        =   55
      Top             =   975
      Width           =   8250
   End
   Begin VB.Frame Frame3 
      Height          =   2550
      Index           =   1
      Left            =   75
      TabIndex        =   23
      Top             =   1080
      Width           =   6150
      Begin VB.TextBox txt_nombre_titular 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1095
         Width           =   3750
      End
      Begin VB.TextBox txt_titular 
         Height          =   315
         Left            =   1290
         TabIndex        =   33
         Top             =   1095
         Width           =   1005
      End
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   1290
         TabIndex        =   35
         Top             =   1440
         Width           =   1005
      End
      Begin VB.TextBox txt_almacen 
         Height          =   315
         Left            =   1290
         TabIndex        =   32
         Top             =   420
         Width           =   1005
      End
      Begin VB.TextBox txt_establecimiento 
         Height          =   315
         Left            =   1290
         TabIndex        =   31
         Top             =   1785
         Width           =   1005
      End
      Begin VB.TextBox txt_referencia 
         Height          =   315
         Left            =   1290
         MaxLength       =   20
         TabIndex        =   30
         Top             =   2130
         Width           =   4365
      End
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   1320
         TabIndex        =   29
         Top             =   750
         Width           =   1005
      End
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   315
         Left            =   2325
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   405
         Width           =   3750
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   750
         Width           =   3750
      End
      Begin VB.TextBox txt_nombre_establecimiento 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1785
         Width           =   3750
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1440
         Width           =   3750
      End
      Begin VB.CommandButton cmd_pasar_todo 
         Height          =   330
         Left            =   5700
         Picture         =   "frmoracle_devoluciones_clientes.frx":4522
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Pasar una factura"
         Top             =   2115
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   60
         Top             =   1155
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   41
         Top             =   1500
         Width           =   525
      End
      Begin VB.Label label 
         BackColor       =   &H000000C0&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   40
         Top             =   120
         Width           =   6075
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   495
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   38
         Top             =   1845
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Referencia:"
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   37
         Top             =   2190
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   36
         Top             =   810
         Width           =   555
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
      Left            =   30
      TabIndex        =   59
      Top             =   105
      Width           =   8325
   End
End
Attribute VB_Name = "frmoracle_devoluciones_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_devolucion_costales As Integer
Dim var_tipo_pedido As Integer
Dim var_localizador_subinventario As String
Dim var_localizador As Integer
Dim var_a�o As Integer
Dim var_almacen_Destino As String
Dim var_primera_vez As Boolean
Dim var_numero_folio As Double
Dim var_cantidad_leida As Double
Dim var_costo As Double
Dim var_precio As Double
Dim var_descripcion_articulo As String
Dim var_estatus_movimiento As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_numero_causa As Integer
Dim var_elimina As Boolean
Dim var_clave_cliente As String
Dim var_clave_titular As String
Dim var_solo_lectura As Boolean
Dim var_clave_almacen_costo As String
Dim var_ventana As Integer
Dim var_clave_moneda As String
Dim VAR_TIPO_LISTA As Integer
Dim var_renglon As Double
Dim var_inventory_item_id As Double
Dim var_unidad_medida As String
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim clnt As New SoapClient30
Dim var_codigo_barras As String
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL = 1
Dim var_ruta_facturas As String



Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
  "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal _
    szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long






Sub ilumina_grid()
   var_n = lv_entradas.ListItems.Count
   For var_i = 1 To var_n
       If var_i = var_renglon Then
          lv_entradas.ListItems.Item(var_i).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_entradas.ListItems.Item(var_i).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
       Else
          lv_entradas.ListItems.Item(var_i).Bold = False
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).Bold = False
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).Bold = False
          lv_entradas.ListItems.Item(var_i).ForeColor = &H80000012
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000012
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000012
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_entradas.ListItems.Item(var_renglon).Selected = True
      lv_entradas.selectedItem.EnsureVisible
   End If
   If lv_entradas.ListItems.Count > 11 Then
      lv_entradas.ColumnHeaders(2).Width = 5050.22
   Else
      lv_entradas.ColumnHeaders(2).Width = 5300.22
   End If
   
   'lv_entradas.Refresh
   
End Sub

















Private Sub cmd_aceptar_pedidos_Click()
   If Me.txt_almacen <> "" Then
      strconsulta = "select * from ra_customer_trx_all where trx_number = ? AND attribute7 = 'FACT. DE COSTALES'"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_numero)
           .Parameters.Append parametro
       End With
       Set rsaux = comandoORA.execute
       Set comandoORA = Nothing
       Set parametro = Nothing
       If Not rsaux.EOF Then
          var_bill_to_site_id = rsaux!bill_to_site_use_id
          var_SHIP_to_site_id = rsaux!SHIP_TO_SITE_USE_ID
          VAR_BILL_TO_CUSTOMER_ID = rsaux!BILL_TO_CUSTOMER_ID
          var_customer_trx_id = rsaux!customer_Trx_id
          strconsulta = "SELECT  distinct  hcas.org_id , arc.collector_id as vcha_age_agente_id, arc.name as vcha_age_nombre FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcas.org_id = 92 AND HCSU.SITE_USE_ID = ? order by arc.name"

          With comandoORA
               .ActiveConnection = cnnoracle_4
               .CommandType = adCmdText
               .CommandText = strconsulta
               Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_bill_to_site_id)
               .Parameters.Append parametro
          End With
          Set rsaux1 = comandoORA.execute
          Set comandoORA = Nothing
          Set parametro = Nothing
          If Not rsaux1.EOF Then
             Me.txt_agente = rsaux1!VCHA_AGE_AGENTE_ID
             Me.txt_nombre_agente = rsaux1!VCHA_AGE_NOMBRE
             
             strconsulta = "SELECT * FROM XXVIA_VW_CLIENTES_BCP WHERE CUST_ACCOUNT_ID = ?"
   
             With comandoORA
                  .ActiveConnection = cnnoracle_4
                  .CommandType = adCmdText
                  .CommandText = strconsulta
                  Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, VAR_BILL_TO_CUSTOMER_ID)
                  .Parameters.Append parametro
             End With
             Set rsaux2 = comandoORA.execute
             Set comandoORA = Nothing
             Set parametro = Nothing
             If Not rsaux2.EOF Then
                Me.txt_clave_titular = rsaux2!CUST_ACCOUNT_ID
                Me.txt_nombre_titular = rsaux2!ACCOUNT_FULL_NAME
                Me.txt_titular = rsaux2!ACCOUNT_NUMBER
                var_clave_titular = rsaux2!CUST_ACCOUNT_ID
                strconsulta = "SELECT * FROM XXVIA_VW_CLIENTES_BCP WHERE SITE_USE_ID = ?"
        
                With comandoORA
                     .ActiveConnection = cnnoracle_4
                     .CommandType = adCmdText
                     .CommandText = strconsulta
                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_bill_to_site_id)
                     .Parameters.Append parametro
                End With
                Set rsaux3 = comandoORA.execute
                Set comandoORA = Nothing
                Set parametro = Nothing
                If Not rsaux3.EOF Then
                   txt_cliente = rsaux3!site_use_id
                   txt_nombre_cliente = rsaux3!razon_social_cliente
                   var_tipo_pedido = rsaux2!ORDER_TYPE_ID
                   strconsulta = "SELECT * FROM XXVIA_VW_CLIENTES_BCP WHERE SITE_USE_ID = ?"
                   With comandoORA
                        .ActiveConnection = cnnoracle_4
                        .CommandType = adCmdText
                        .CommandText = strconsulta
                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_SHIP_to_site_id)
                        .Parameters.Append parametro
                   End With
                   Set rsaux4 = comandoORA.execute
                   Set comandoORA = Nothing
                   Set parametro = Nothing
                   If Not rsaux4.EOF Then
                      Me.txt_establecimiento = rsaux4!site_use_id
                      Me.txt_nombre_establecimiento = rsaux4!razon_social_cliente
                      Me.txt_referencia = "DC BULTOS " + Me.txt_numero
                      
                      var_cadena = "SELECT  hcsu.price_list_id, hcsu.order_type_id,hca.cust_account_id, hcp.site_use_id AS VCHA_CLI_CLAVE_ID,hl.address1 VCHA_CLI_NOMBRE FROM hz_parties hp,hz_party_sites hps,hz_cust_accounts hca,hz_cust_acct_sites_all hcas,hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND  hcp.site_use_id = " + Me.txt_cliente
                      rsaux10.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                      If Not rsaux10.EOF Then
                         var_clave_lista_precios = rsaux10!price_list_id
                      Else
                         var_clave_lista_precios = 0
                      End If
                      rsaux10.Close
                      
                      
                      strconsulta = "SELECT A.INVENTORY_ITEM_ID, B.SEGMENT1, QUANTITY_ORDERED, a.description, PRIMARY_UOM_CODE FROM RA_CUSTOMER_TRX_LINES_ALL A, XXVIA_SYSTEM_ITEMS_B B WHERE CUSTOMER_TRX_ID = ? AND QUANTITY_ORDERED IS NOT  NULL AND A.INVENTORY_ITEM_ID = B.INVENTORY_ITEM_ID AND B.ORGANIZATION_ID = 93 and extended_amount >= 0"
                      With comandoORA
                           .ActiveConnection = cnnoracle_4
                           .CommandType = adCmdText
                           .CommandText = strconsulta
                           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, var_customer_trx_id)
                           .Parameters.Append parametro
                      End With
                      Set rsaux5 = comandoORA.execute
                      Set comandoORA = Nothing
                      Set parametro = Nothing
                      While Not rsaux5.EOF
                            Me.txt_codigo = rsaux5!SEGMENT1
                            var_cantidad_leida = rsaux5!QUANTITY_ORDERED
                            Me.txt_codigo.Enabled = True
                            var_inventory_item_id = rsaux5!inventory_item_id
                            var_descripcion_articulo = rsaux5!Description
                            var_unidad_medida = rsaux5!PRIMARY_UOM_CODE
                            var_localizador_subinventario = ""
                            Call txt_foco_GotFocus
                            rsaux5.MoveNext
                            var_devolucion_costales = 1
                      Wend
                      rsaux5.Close
                      
                      Me.txt_agente.Enabled = False
                      Me.txt_nombre_agente.Enabled = False
                      Me.txt_titular.Enabled = False
                      Me.txt_nombre_titular.Enabled = False
                      Me.txt_cliente.Enabled = False
                      Me.txt_nombre_cliente.Enabled = False
                      Me.txt_establecimiento.Enabled = False
                      Me.txt_nombre_establecimiento.Enabled = False
                      Me.txt_codigo.Enabled = False
                      
                   Else
                      MsgBox "El establecimiento no existe", vbOKOnly, "ATENCION"
                   End If
                   rsaux4.Close
                Else
                   MsgBox "El cliente no existe", vbOKOnly, "ATENCION"
                End If
                rsaux3.Close
             Else
                MsgBox "El titular no existe", vbOKOnly, "ATENCION"
             End If
             rsaux2.Close
          Else
             MsgBox "El agente no existe", vbOKOnly, "ATENCION"
          End If
          rsaux1.Close
       Else
          MsgBox "La factura no existe", vbOKOnly, "ATENCION"
       End If
       If rsaux1.State = 1 Then
          rsaux.Close
       End If
    Else
       MsgBox "No se a seleccionado un almac�n destino", vbOKOnly, "ATENCION"
    End If
    Me.frm_pasar_todo.Visible = False
End Sub

Private Sub cmd_asignar_causa_devolucion_Click()
      rs.Open "select * from xxvia_tb_Devoluciones_clientes where numero = " + CStr(var_numero_folio) + " and organizacion = " + var_unidad_organizacional + " AND MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_estatus_movimiento = IIf(IsNull(rs!estatus), "", rs!estatus)
      End If
      rs.Close
      If var_clave_movimiento = "DC" Or var_clave_movimiento = "SNC" Then
         If var_estatus_movimiento = "I" Then
            var_numero_folio_devoluciones = CDbl(Me.txt_folio)
            var_clave_almacen_devolucion = Me.txt_almacen
            var_referencia_global_dev = Me.txt_referencia
            frmoracle_asignacion_causas_devolucion.Show 1
         Else
            MsgBox "El movimiento no a sido cerrado.", vbOKOnly, "ATENCION"
         End If
      End If
End Sub

Private Sub cmd_buscar_Click()
   var_ventana = 1
   frm_busqueda.Visible = True
   txt_busqueda_folio.SetFocus
End Sub


Private Sub cmd_cancelar_pedidos_Click()
   Me.frm_pasar_todo.Visible = False
End Sub

Private Sub cmd_devolucion_Click()
   Dim var_inserta As Boolean
   Dim var_factura As Integer
   Dim var_posible_cliente As Boolean
   
   If rsaux11.State = 1 Then
      rsaux11.Close
   End If
   'rsaux11.Open "select codigo, cantidad from pedido_301014", cnn, adOpenDynamic, adLockOptimistic
   'If rsaux11.State = 1 Then
   '   rsaux11.Close
   'End If
   rsaux12.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   'rsaux11.Open "select oha.header_id, segment1, SUM(pricing_quantity) as cantidad from oe_order_headers_all oha, oe_order_lines_all ola, xxvia_system_items_b b where oha.header_id = ola.header_id and oha.ship_from_org_id = b.organization_id and ola.inventory_item_id = b.inventory_item_id and order_number = " + Me.txt_referencia + " GROUP BY oha.header_id ,SEGMENT1", cnnoracle_4, adOpenDynamic, adLockOptimistic
   
   'rsaux11.Open "select segment1 as codigo, sum(floa_sal_cantidad_leida) as cantidad from xxvia_Tb_Salidas_cajas where source_header_number  in ('380026','380028','380029','380030','380031') group by segment1", cnnoracle_4, adOpenDynamic, adLockOptimistic
   rsaux11.Open "select ordered_item as segment1, ordered_quantity as cantidad from oe_order_headers_all a, oe_order_lines_all b where order_number in ('434817','434812', '434836') and a.header_id = b.header_id", cnnoracle_4, adOpenDynamic, adLockOptimistic
   
   'rsaux11.Open "select B.SEGMENT1 , SHIPPED_QUANTITY as cantidad from wsh_deliverables_v A, XXVIA_SYSTEM_ITEMS_B B where source_header_number = 413146  AND A.INVENTORY_ITEM_ID = B.INVENTORY_ITEM_ID  AND B.ORGANIZATION_ID = 93 AND SHIPPED_QUANTITY > 0", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_faltantes = ""
   Me.txt_codigo = rsaux11!SEGMENT1
   'While Not rsaux11.EOF
   '      rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
   '      If rsaux8.EOF Then
   '         If var_Cadena_faltantes = "" Then
   '            var_Cadena_faltantes = rsaux11!SEGMENT1
   '         Else
   '            var_Cadena_faltantes = var_Cadena_faltantes + ", " + rsaux11!SEGMENT1
   '         End If
   '      End If
   '      rsaux8.Close
   '      rsaux11.MoveNext
   'Wend
   rsaux11.MoveFirst
   If var_Cadena_faltantes = "" Then
      var_primera_vez = True
      While Not rsaux11.EOF
            Me.txt_codigo = rsaux11!SEGMENT1
            If Trim(Me.txt_codigo) <> "" Then
                rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                If Not rsaux8.EOF Then
                   var_unidad_medida = rsaux8!PRIMARY_UOM_CODE
                   var_descripcion_articulo = rsaux8!Description
                   var_inventory_item_id = rsaux8!inventory_item_id
                   var_salida_masiva = IIf(IsNull(rsaux8!attribute10), "N", rsaux8!attribute10)
                End If
                rsaux8.Close
            End If
         
            var_cantidad_leida = rsaux11!cantidad
            If Trim(txt_codigo.Text) <> "" Then
               If var_primera_vez = True Then
                  rs.Open "select * from xxvia_tb_folios_dev_clientes WHERE MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_numero_folio = rs(0).Value + 1
                     Me.txt_folio = rs(0).Value + 1
                     rsaux.Open "update xxvia_tb_folios_dev_clientes set folio =  folio + 1 WHERE MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux.Open "insert into xxvia_tb_folios_dev_clientes (folio, MOVIMIENTO) values (1,'" + var_clave_movimiento + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_numero_folio = 1
                     Me.txt_folio = 1
                  End If
                  rs.Close
                  var_primera_vez = False
               End If
               Cadena = "select * from xxvia_tb_devoluciones_clientes where numero = " + Str(var_numero_folio) + " and codigo = '" + txt_codigo + "' and inventory_item_id = " + CStr(var_inventory_item_id) + " and localizador = '" + var_localizador_subinventario + "' AND MOVIMIENTO = '" + var_clave_movimiento + "'"
               rs.Open Cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If rs.EOF Then
                  lbl_total = CStr(CDbl(lbl_total) + var_cantidad_leida)
                  var_inserta = False
                  rs.Close
                  If Me.txt_establecimiento = "" Then
                     Me.txt_establecimiento = 0
                  End If
                  var_cadena = "insert into xxvia_tb_devoluciones_clientes (numero, organizacion, inventory_item_id, codigo, cantidad, descripcion, estatus, agente, cliente, establecimiento, titular, nombre_agente, almacen, nombre_almacen, nombre_cliente, nombre_establecimiento, referencia, usuario, maquina, fecha_inicio, unidad_medida, precio, localizador, movimiento,tipo_pedido, factura)"
                  var_cadena = var_cadena + " values (" + CStr(var_numero_folio) + "," + var_unidad_organizacional + "," + CStr(var_inventory_item_id) + ",'" + Me.txt_codigo + "'," + CStr(var_cantidad_leida) + ",'" + var_descripcion_articulo + "',''," + Me.txt_agente + "," + Me.txt_cliente + "," + Me.txt_establecimiento + "," + CStr(var_clave_titular) + ",'" + Me.txt_nombre_agente + "','" + Me.txt_almacen + "','" + Me.txt_nombre_almacen + "','" + Me.txt_nombre_cliente + "','" + Me.txt_nombre_establecimiento + "','" + Me.txt_referencia + "','" + var_clave_usuario_global + "', '" + fun_NombrePc + "','" + CStr(Date) + "','" + var_unidad_medida + "'," + CStr(0) + ",'" + var_localizador_subinventario + "','" + var_clave_movimiento + "'," + CStr(var_tipo_pedido) + ",0)"
                  'MsgBox var_cadena
                  rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  valor = Trim(txt_codigo)
       
                  Set list_item = lv_entradas.ListItems.Add(, , Trim(txt_codigo))
                  list_item.SubItems(1) = var_descripcion_articulo
                  list_item.SubItems(2) = var_cantidad_leida
                  list_item.SubItems(3) = var_localizador_subinventario
                  var_renglon = lv_entradas.ListItems.Count
                  Call ilumina_grid
                  txt_codigo = ""
               Else
                  var_inserta = False
                  lbl_total = CStr(CDbl(lbl_total) + var_cantidad_leida)
                  rs.Close
                  rs.Open "update xxvia_tb_devoluciones_clientes set cantidad = cantidad +" + CStr(var_cantidad_leida) + " where numero = " + CStr(var_numero_folio) + " and inventory_item_id = " + CStr(var_inventory_item_id) + " and codigo = '" + Me.txt_codigo + "' and localizador = '" + var_localizador_subinventario + "' and movimiento = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  valor = Me.txt_codigo
                  var_j = 1
                  For var_j = 1 To Me.lv_entradas.ListItems.Count
                      Me.lv_entradas.ListItems.Item(var_j).Selected = True
                      If Me.lv_entradas.selectedItem = Me.txt_codigo And Trim(Me.lv_entradas.selectedItem.SubItems(3)) = Trim(var_localizador_subinventario) Then
                         Me.lv_entradas.selectedItem.SubItems(2) = CDbl(Me.lv_entradas.selectedItem.SubItems(2)) + var_cantidad_leida
                         var_renglon = var_j
                      End If
                  Next var_j
                  Call ilumina_grid
                  txt_codigo = ""
               End If
               txt_codigo = ""
            End If
           rsaux11.MoveNext
      Wend
      rsaux11.Close
   Else
      MsgBox "Faltan los siguientes c�digos " + var_Cadena_faltantes, vbOKOnly, "ATENCION"
   End If


End Sub

Private Sub cmd_factura_Click()
   Dim var_inserta As Boolean
   Dim var_factura As Integer
   Dim var_posible_cliente As Boolean
   If rsaux11.State = 1 Then
      rsaux11.Close
   End If
   
   'rsaux11.Open "select segment1, sum(floa_sal_cantidad_leida) cantidad from xxvia_tb_Salidas_Cajas where source_header_number = 658602 and inte_paq_caja <> 22 and floa_Sal_Cantidad_leida > 0 group by segment1", cnnoracle_4, adOpenDynamic, adLockOptimistic
   rsaux11.Open "select segment1, sum(floa_sal_cantidad_leida) cantidad from xxvia_tb_Salidas_Cajas where source_header_number = 662655 and floa_Sal_Cantidad_leida > 0 group by segment1", cnnoracle_4, adOpenDynamic, adLockOptimistic
   'rsaux11.Open "select segment1 , empacado cantidad from tb_oracle_pedidos where source_header_number = " + Me.txt_referencia + " and empacado >0", cnn, adOpenDynamic, adLockOptimistic
   var_primera_vez = True
   While Not rsaux11.EOF
         Me.txt_codigo = rsaux11!SEGMENT1
         If Len(Me.txt_codigo) = 5 Then
            Me.txt_codigo = "000" + Me.txt_codigo
         End If
         If Len(Me.txt_codigo) = 4 Then
            Me.txt_codigo = "0000" + Me.txt_codigo
         End If
         If Trim(Me.txt_codigo) <> "" Then
             rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
             If Not rsaux8.EOF Then
                var_unidad_medida = rsaux8!PRIMARY_UOM_CODE
                var_descripcion_articulo = rsaux8!Description
                var_inventory_item_id = rsaux8!inventory_item_id
                var_salida_masiva = IIf(IsNull(rsaux8!attribute10), "N", rsaux8!attribute10)
             End If
             rsaux8.Close
         End If
         
         var_cantidad_leida = rsaux11!cantidad
         If Trim(txt_codigo.Text) <> "" Then
            If var_primera_vez = True Then
               rs.Open "select * from xxvia_tb_folios_dev_clientes WHERE MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_numero_folio = rs(0).Value + 1
                  Me.txt_folio = rs(0).Value + 1
                  rsaux.Open "update xxvia_tb_folios_dev_clientes set folio =  folio + 1 WHERE MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               Else
                  rsaux.Open "insert into xxvia_tb_folios_dev_clientes (folio, MOVIMIENTO) values (1,'" + var_clave_movimiento + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  var_numero_folio = 1
                  Me.txt_folio = 1
               End If
               rs.Close
               var_primera_vez = False
            End If
            Cadena = "select * from xxvia_tb_devoluciones_clientes where numero = " + Str(var_numero_folio) + " and codigo = '" + txt_codigo + "' and inventory_item_id = " + CStr(var_inventory_item_id) + " and localizador = '" + var_localizador_subinventario + "' AND MOVIMIENTO = '" + var_clave_movimiento + "'"
            rs.Open Cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If rs.EOF Then
               lbl_total = CStr(CDbl(lbl_total) + var_cantidad_leida)
               var_inserta = False
               rs.Close
               If Me.txt_establecimiento = "" Then
                  Me.txt_establecimiento = 0
               End If
               var_cadena = "insert into xxvia_tb_devoluciones_clientes (numero, organizacion, inventory_item_id, codigo, cantidad, descripcion, estatus, agente, cliente, establecimiento, titular, nombre_agente, almacen, nombre_almacen, nombre_cliente, nombre_establecimiento, referencia, usuario, maquina, fecha_inicio, unidad_medida, precio, localizador, movimiento,tipo_pedido)"
               var_cadena = var_cadena + " values (" + CStr(var_numero_folio) + "," + var_unidad_organizacional + "," + CStr(var_inventory_item_id) + ",'" + Me.txt_codigo + "'," + CStr(var_cantidad_leida) + ",'" + var_descripcion_articulo + "',''," + Me.txt_agente + "," + Me.txt_cliente + "," + Me.txt_establecimiento + "," + CStr(var_clave_titular) + ",'" + Me.txt_nombre_agente + "','" + Me.txt_almacen + "','" + Me.txt_nombre_almacen + "','" + Me.txt_nombre_cliente + "','" + Me.txt_nombre_establecimiento + "','" + Me.txt_referencia + "','" + var_clave_usuario_global + "', '" + fun_NombrePc + "','" + CStr(Date) + "','" + var_unidad_medida + "',0,'" + var_localizador_subinventario + "','" + var_clave_movimiento + "'," + CStr(var_tipo_pedido) + ")"
               'MsgBox var_cadena
               rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               valor = Trim(txt_codigo)
       
               Set list_item = lv_entradas.ListItems.Add(, , Trim(txt_codigo))
               list_item.SubItems(1) = var_descripcion_articulo
               list_item.SubItems(2) = var_cantidad_leida
               list_item.SubItems(3) = var_localizador_subinventario
               var_renglon = lv_entradas.ListItems.Count
               Call ilumina_grid
               txt_codigo = ""
            Else
               var_inserta = False
               lbl_total = CStr(CDbl(lbl_total) + var_cantidad_leida)
               rs.Close
               rs.Open "update xxvia_tb_devoluciones_clientes set cantidad = cantidad +" + CStr(var_cantidad_leida) + " where numero = " + CStr(var_numero_folio) + " and inventory_item_id = " + CStr(var_inventory_item_id) + " and codigo = '" + Me.txt_codigo + "' and localizador = '" + var_localizador_subinventario + "' and movimiento = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               valor = Me.txt_codigo
               var_j = 1
               For var_j = 1 To Me.lv_entradas.ListItems.Count
                   Me.lv_entradas.ListItems.Item(var_j).Selected = True
                   If Me.lv_entradas.selectedItem = Me.txt_codigo And Trim(Me.lv_entradas.selectedItem.SubItems(3)) = Trim(var_localizador_subinventario) Then
                      Me.lv_entradas.selectedItem.SubItems(2) = CDbl(Me.lv_entradas.selectedItem.SubItems(2)) + var_cantidad_leida
                      var_renglon = var_j
                   End If
               Next var_j
               Call ilumina_grid
               txt_codigo = ""
            End If
            txt_codigo = ""
         End If
        rsaux11.MoveNext
   Wend
   rsaux11.Close
         If var_clave_movimiento = "VDIII" Then
            var_numero_folio_devoluciones = CDbl(Me.txt_folio)
            If var_estatus_movimiento = "I" Then
            Else
               rsaux7.Open "select name from qp_secu_list_headers_v where list_header_id = " + CStr(var_clave_lista_precios), cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_lista_precios = rsaux7!Name
               rsaux7.Close
               var_clave_tipo_pedido = 1464
               If var_clave_tipo_pedido > 0 Then
                  If var_lista_precios <> "" Then
                     If rs.State = 1 Then
                        rs.Close
                     End If
                     rs.Open "SELECT A.ESTATUS, A.MOVIMIENTO, A.NUMERO, A.ORGANIZACION, A.inventory_item_id, A.almacen, A.titular, A.unidad_medida, A.precio, A.TITULAR, A.CLIENTE, A.ESTABLECIMIENTO, A.LOCALIZADOR, a.codigo, SUM(A.cantidad) AS CANTIDAD FROM XXVIA_TB_DEVOLUCIONES_CLIENTES A WHERE A.NUMERO = " + CStr(var_numero_folio_devoluciones) + " AND A.ORGANIZACION = " + var_unidad_organizacional + "  AND A.MOVIMIENTO = '" + var_clave_movimiento + "' GROUP BY A.ESTATUS, A.MOVIMIENTO, A.NUMERO, A.ORGANIZACION, A.inventory_item_id, A.almacen, A.titular, A.unidad_medida, A.precio, A.TITULAR, A.CLIENTE, A.ESTABLECIMIENTO, A.LOCALIZADOR, a.codigo", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If IIf(IsNull(rs!estatus), "", rs!estatus) = "" Then
                        If Not rs.EOF Then
                           rs.MoveFirst
                           var_cadena_posible_existencias = ""
                           x = 0
                           If x = 1 Then
                           While Not rs.EOF
                                 strconsulta = "select * from Xxvia_vw_existencias_inv where organization_id = ? and subinventory_code = ? and segment1 = ?"
                                 With comandoORA
                                      .ActiveConnection = cnnoracle_4
                                      .CommandType = adCmdText
                                      .CommandText = strconsulta
                                      Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_unidad_organizacional))
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_almacen)
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!codigo)
                                      .Parameters.Append parametro
                                 End With
                                 Set rsaux = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing
                           
                                 'rsaux.Open "select * from Xxvia_vw_existencias_inv where organization_id = " + var_unidad_organizacional + " and subinventory_code = '" + Me.txt_almacen + "' and segment1 = '" + rs!CODIGO + "'"
                                 If Not rsaux.EOF Then
                                    var_disponible = IIf(IsNull(rsaux!Disponible), 0, rsaux!Disponible)
                                 Else
                                    var_disponible = 0
                                 End If
                                 If var_disponible < rs!cantidad Then
                                    If var_cadena_posible_existencias = "" Then
                                       var_cadena_posible_existencias = "Movimiento = " + CStr(rs!cantidad) + " Disponible = " + CStr(var_disponible)
                                    Else
                                       var_cadena_posible_existencias = var_cadena_posible_existencias + "--- Movimiento = " + CStr(rs!cantidad) + " Disponible = " + CStr(var_disponible)
                                    End If
                                 End If
                                 rsaux.Close
                                 rs.MoveNext
                           Wend
                           End If
                           rs.MoveFirst
var_cadena_posible_existencias = ""
                           If var_cadena_posible_existencias = "" Then
                              'var_clave_lista_precios = 719011
                              rsaux7.Open "select name from qp_secu_list_headers_v where list_header_id = " + CStr(var_clave_lista_precios), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              var_lista_precios = rsaux7(0).Value
                              rsaux7.Close
                              
                              var_cadena = "INSERT INTO oe_headers_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref, creation_date, created_by, last_update_date, last_updated_by, operation_code , sold_to_org_id        , SHIP_TO_ORG_id                   ,INVOICE_TO_ORG_ID     , Order_type_ID, PRICE_LIST, org_id, ship_from_org_id)"
                              var_cadena = var_cadena + "  VALUES (1001,'VDISID_" + Me.txt_folio + "',SYSDATE,-1,SYSDATE, -1,'INSERT', " + CStr(rs!TITULAR) + "," + CStr(rs!establecimiento) + "," + CStr(rs!Cliente) + "," + CStr(var_clave_tipo_pedido) + ",'" + var_lista_precios + "'," + var_empresa + "," + var_unidad_organizacional + ")"
                              'MsgBox var_numero_folio_devoluciones
                              If rsaux.State = 1 Then
                                 rsaux.Close
                              End If
                              rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              var_i = 0
                              While Not rs.EOF
                                    var_i = var_i + 1
                                    rsaux10.Open "SELECT PRIMARY_UOM_CODE FROM xxvia_system_items_b WHERE INVENTORY_ITEM_ID = " + CStr(rs!inventory_item_id) + " AND ORGANIZATION_ID = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    If Not rsaux10.EOF Then
                                       VAR_MEDIDA = rsaux10(0).Value
                                    End If
                                    rsaux10.Close
                                 
                                    var_cadena = "INSERT INTO oe_lines_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref,orig_sys_line_ref,inventory_item_id,ordered_quantity, operation_code, created_by, creation_date, last_updated_by, last_update_date, unit_selling_price, unit_list_price, calculate_price_flag, PRICING_QUANTITY, PRICING_QUANTITY_UOM, ATTRIBUTE1, subinventory, org_id, ship_from_org_id)"
                                    var_cadena = var_cadena + " VALUES (1001,'SIDVDI_" + Trim(CStr(var_numero_folio_devoluciones)) + "','" + CStr(var_i) + "', " + CStr(rs!inventory_item_id) + ", " + CStr(rs!cantidad) + ",'INSERT', -1,SYSDATE, -1,SYSDATE,0,0,'Y', " + CStr(rs!cantidad) + ", '" + VAR_MEDIDA + "','','" + Me.txt_almacen + "'," + var_empresa + "," + var_unidad_organizacional + ")"
                                    'MsgBox var_cadena
                                    rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    rs.MoveNext
                              Wend
                              On Error GoTo SALIR
                              rsaux.Open "INSERT INTO oe_actions_iface_all (order_source_ID, orig_sys_document_ref, operation_code) VALUES (1001, 'SIDVDI_" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "','BOOK_ORDER')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If rsaux.State = 1 Then
                                 rsaux.Close
                              End If
                              rsaux.Open "CALL XXVIA_PK_INTERFACES_OM.importar_pedido('SIDVDI_" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "'," + var_empresa + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If rsaux.State = 1 Then
                                 rsaux.Close
                              End If
                              rsaux.Open "UPDATE XXVIA_TB_DEVOLUCIONES_CLIENTES A SET ESTATUS = 'I' WHERE A.NUMERO = " + CStr(var_numero_folio_devoluciones) + " AND A.ORGANIZACION = " + var_unidad_organizacional + "  AND A.MOVIMIENTO = '" + var_clave_movimiento + "'"
                              rsaux.Open "select order_number from oe_order_headers_all where orig_sys_document_ref = 'SIDVDI_" + Trim(CStr(var_numero_folio_devoluciones)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              var_pedido = rsaux(0).Value
                              rsaux.Close
                              MsgBox var_pedido
                           Else
                              MsgBox "No se puede dar salida debido a: " + var_cadena_posible_existencias, vbOKOnly, "ATENCION"
                           End If
                        End If
                     End If
                     rs.Close
                  Else
                     MsgBox "No se a indicado una lista de precios", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "No se a indicado un tipo de pedido", vbOKOnly, "ATENCION"
               End If
            End If
         End If
SALIR:
   If Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
       rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      Resume
   Else
      MsgBox Err.Description
      'Resume
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
   End If
   Exit Sub
salir_factura:
   MsgBox "Surgio un error al generar los documentos electr�nicos", vbOKOnly, "ATENCION"
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
   If objConn.State = 1 Then
      objConn.RollbackTrans
      objConn.Close
   End If



End Sub

Private Sub cmd_imprimir_Click()
   Dim var_precio_inflado As Double
   Dim var_precio_descuento As Double
   Dim objConn As New ADODB.Connection
   Dim objCmd As New ADODB.Command
   Dim objParm As ADODB.Parameter
   Dim clnt As New SoapClient30
   Dim var_con As String
   Dim var_customer_trx_id As Double
   Dim var_estatus_factura As String
   Dim numero_req As Double
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
   'On Error GoTo salir2
   If var_numero_folio > 0 Then
      rs.Open "select * from xxvia_tb_Devoluciones_clientes where numero = " + CStr(var_numero_folio) + " and organizacion = " + var_unidad_organizacional + " AND MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_estatus_movimiento = IIf(IsNull(rs!estatus), "", rs!estatus)
      End If
      rs.Close
      If var_clave_movimiento = "DC" Or var_clave_movimiento = "SNC" Then
         If var_estatus_movimiento = "I" Then
            var_numero_folio_devoluciones = CDbl(Me.txt_folio)
            var_clave_almacen_devolucion = Me.txt_almacen
            var_referencia_global_dev = Me.txt_referencia
            frmoracle_devoluciones_desgloce.Show 1
         Else
            
            rs.Open "select codigo from xxvia_tb_devoluciones_clientes where numero = " + Me.txt_folio + " and movimiento = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = ""
            While Not rs.EOF
                  If rsaux1.State = 1 Then
                     rsaux1.Close
                  End If
                  rsaux1.Open "select CUSTOMER_ORDER_FLAG, customer_order_enabled_flag, SHIPPABLE_ITEM_FLAG, INTERNAL_ORDER_FLAG, internal_order_enabled_flag, SO_TRANSACTIONS_FLAG, RETURNABLE_FLAG, INVOICEABLE_ITEM_FLAG from xxvia_system_items_b where segment1='" + rs!codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If IIf(IsNull(rsaux1!RETURNABLE_FLAG), "", rsaux1!RETURNABLE_FLAG) <> "Y" Then
                     If var_cadena = "" Then
                        var_cadena = var_cadena + " " + rs!codigo
                     Else
                        var_cadena = var_cadena + ", " + rs!codigo
                     End If
                  End If
                  rsaux1.Close
                  rs.MoveNext
            Wend
            rs.Close
            var_cadena_codigos_retornables = var_cadena
            If var_cadena = "" Then
               rsaux8.Open "SELECT * FROM XXVIA_TB_DEV_CLIENTES_DESGLOCE WHERE NUMERO = " + Me.txt_folio + " AND ORGANIZACION = " + var_unidad_organizacional + " AND MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If rsaux8.EOF Then
                  rsaux9.Open "SELECT * FROM XXVIA_TB_DEVOLUCIONES_CLIENTES WHERE NUMERO = " + Me.txt_folio + " AND ORGANIZACION = " + var_unidad_organizacional + " AND MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If rsaux7.State = 1 Then
                     rsaux7.Close
                  End If
                  'var_cadena = "SELECT  hcsu.attribute1 AS RFC,hcsu.order_type_id,hca.cust_account_id,hcp.site_use_id AS VCHA_CLI_CLAVE_ID,hl.address1 VCHA_CLI_NOMBRE FROM hz_parties hp,hz_party_sites hps,hz_cust_accounts hca,hz_cust_acct_sites_all hcas,hz_cust_site_uses_all hcsu, hz_locations hl,hr_operating_units hr,hz_customer_profiles hcp,ar_collectors arc,ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id=hps.party_id AND hps.party_site_id=hcas.party_site_id AND hca.cust_account_id=hcas.cust_account_id AND hcas.cust_acct_site_id=hcsu.cust_acct_site_id AND hps.location_id=hl.location_id AND hcas.org_id=hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcp.collector_id = " + Me.txt_agente + " AND hcp.site_use_id = " + Me.txt_cliente
                  'MsgBox Me.txt_cliente
                  var_cadena = "select * from xxvia_vw_clientes_bcp where SITE_USE_ID = " + Me.txt_cliente
                  If rsaux7.State = 1 Then
                     rsaux7.Close
                  End If
                  rsaux7.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  var_rfc = IIf(IsNull(rsaux7!rfc), "", rsaux7!rfc)
                  rsaux7.Close
                  var_consecutivo = 0
                  While Not rsaux9.EOF
                        var_numero_factura = 0
                        VAR_PORCENTAJE_FIN = 0
                        VAR_NOTA_CREDITO_DF = 0
                        If rsaux.State = 1 Then
                           rsaux.Close
                        End If
                        x = 1
                        ' esto no recuerdo porque se comento
                        'rsaux.Open "SELECT L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id, NVL(SALES_ORDER_LINE,ROWNUM) AS LINEA, SUM(NVL(L.GROSS_unit_selling_price,l.unit_selling_price)) AS PRECIO FROM RA_CUSTOMER_TRX_LINES_ALL L, RA_CUSTOMER_TRX_ALL E, ra_cust_trx_types_all TYPES Where TYPES.TYPE = 'INV' AND TYPES.cust_trx_type_id = E.cust_trx_type_id AND TYPES.org_id = E.org_id AND l.customer_trx_id = E.customer_trx_id AND L.inventory_item_id = " + CStr(rsaux9!INVENTORY_ITEM_ID) + " AND E.sold_to_customer_id = " + CStr(rsaux9!TITULAR) + "  GROUP BY L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id, NVL(SALES_ORDER_LINE,ROWNUM) ORDER BY trx_date desc", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux.Open "SELECT L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id, SUM(NVL(L.GROSS_unit_selling_price,l.unit_selling_price)) AS PRECIO FROM RA_CUSTOMER_TRX_LINES_ALL L, RA_CUSTOMER_TRX_ALL E, ra_cust_trx_types_all TYPES Where TYPES.TYPE = 'INV' AND TYPES.cust_trx_type_id = E.cust_trx_type_id AND TYPES.org_id = E.org_id AND l.customer_trx_id = E.customer_trx_id AND L.inventory_item_id = " + CStr(rsaux9!inventory_item_id) + " AND E.sold_to_customer_id = " + CStr(rsaux9!TITULAR) + " and trx_date >= to_date('01/01/2016','DD/MM/YYYY') and trx_number not like 'F%'  GROUP BY L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id  ORDER BY trx_date desc", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
''''''


                        'rsaux.Open "SELECT L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id, SUM(NVL(L.GROSS_unit_selling_price,l.unit_selling_price)) AS PRECIO FROM RA_CUSTOMER_TRX_LINES_ALL L, RA_CUSTOMER_TRX_ALL E, ra_cust_trx_types_all TYPES Where TYPES.TYPE = 'INV' AND TYPES.cust_trx_type_id = E.cust_trx_type_id AND TYPES.org_id = E.org_id AND l.customer_trx_id = E.customer_trx_id AND L.inventory_item_id = " + CStr(rsaux9!inventory_item_id) + " AND E.sold_to_customer_id = " + CStr(rsaux9!TITULAR) + "  GROUP BY L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id  ORDER BY trx_date desc", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           var_veces = 1
                           If rsaux7.State = 1 Then
                              rsaux7.Close
                           End If
                           rsaux7.Open "select count(*) from RA_CUSTOMER_TRX_LINES_ALL where customer_trx_id= " + CStr(rsaux!customer_Trx_id) + " and inventory_item_id = " + CStr(rsaux9!inventory_item_id) + " and unit_selling_price >0", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux7.EOF Then
                              var_veces = IIf(IsNull(rsaux7(0).Value), 1, rsaux7(0).Value)
                           End If
                           rsaux7.Close
                           If rsaux!Precio = 0 Then
                              var_precio = 0
                           Else
                              var_precio = rsaux!Precio / var_veces
                           End If
                           
                           
                           If rsaux!Precio = 0 Then
                              var_precio_entero = 0
                           Else
                              var_precio_entero = rsaux!Precio / var_veces
                           End If
                           var_numero_factura = rsaux!customer_Trx_id
                           rsaux5.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           
                           
                           
                           If var_attribute10 = "0" Then
                              var_cadena = "SELECT ARPA.APPLIED_CUSTOMER_TRX_ID AS FACTURA_ID, ARPA.CUSTOMER_TRX_ID AS NOTA_CREDITO_ID, ARPA.ACCTD_AMOUNT_APPLIED_TO AS MONTO_APLICADO, RCT.CUST_TRX_TYPE_ID, RCTL.ATTRIBUTE11, RCTL.ATTRIBUTE10, ARPA.AMOUNT_APPLIED, acr.amount FROM AR_RECEIVABLE_APPLICATIONS_ALL ARPA, RA_CUSTOMER_TRX_ALL RCT, RA_CUSTOMER_TRX_LINES_ALL RCTL, ar_cash_receipts_all acr WHERE ARPA.APPLICATION_TYPE = 'CM' AND ARPA.CUSTOMER_TRX_ID = RCT.CUSTOMER_TRX_ID AND RCT.CUST_TRX_TYPE_ID IN (SELECT ATTRIBUTE2 From RA_CUST_TRX_TYPES_ALL WHERE ATTRIBUTE2 IS NOT NULL) AND ARPA.CUSTOMER_TRX_ID  = RCTL.CUSTOMER_TRX_ID AND RCTL.ATTRIBUTE11 IS NOT NULL AND ARPA.APPLIED_CUSTOMER_TRX_ID = " + CStr(var_numero_factura) + " and RCTL.ATTRIBUTE10 = acr.cash_receipt_id and ARPA.ACCTD_AMOUNT_APPLIED_TO > 0 order by arpa.last_update_date desc"
                              rsaux5.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              var_attribute10 = ""
                              If Not rsaux5.EOF Then
                                 While Not rsaux5.EOF
                                      If var_attribute10 = "" Then
                                         var_attribute10 = rsaux5!attribute10
                                      Else
                                         var_attribute10 = var_attribute10 + "," + rsaux5!attribute10
                                      End If
                                       rsaux5.MoveNext
                                 Wend
                              Else
                                 var_attribute10 = 0
                              End If
                              rsaux5.Close
                           End If
                           
                           
                           var_cadena = "select rec.CUSTOMER_TRX_ID, nvl(sum(rec.amount_applied),0) as importe_df from ar_receivable_applications_all rec Inner join ar_payment_schedules_all pay on rec.payment_schedule_id = pay.payment_schedule_id Inner join ra_cust_trx_types_all on pay.cust_trx_type_id = ra_cust_trx_types_all.cust_trx_type_id Where rec.applied_customer_trx_id = " + CStr(var_numero_factura) + " and rec.apply_date < sysdate and rec.display = 'Y' and application_type = 'CM' and ra_cust_trx_types_all.cust_trx_type_id in (1564,1028) group by rec.CUSTOMER_TRX_ID "
                           rsaux5.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux5.EOF Then
                              var_importe_total_df = 0
                              var_notas_credito_df = ""
                              While Not rsaux5.EOF
                                    var_importe_total_df = var_importe_total_df + IIf(IsNull(rsaux5!importe_df), 0, rsaux5!importe_df)
                                    If var_notas_credito_df = "" Then
                                       var_notas_credito_df = CStr(rsaux5!customer_Trx_id)
                                    Else
                                       var_notas_credito_df = var_notas_credito_df + ", " + CStr(rsaux5!customer_Trx_id)
                                    End If
                                    rsaux5.MoveNext
                              Wend
                              rsaux5.MoveFirst
                              'var_cadena = "select sum(amount_applied) amount_applied from ar_receivable_applications_all Where applied_customer_trx_id = " + CStr(VAR_NUMERO_fACTURA) + " and display = 'Y' and application_type = 'CASH' and cash_receipt_id in( " + CStr(var_attribute10) + ")"
                              var_cadena = "select SUM(nvl(gross_extended_amount, extended_amount)) AS amount_applied from ra_customer_trx_lines_all where customer_trx_id = " + CStr(var_numero_factura) + " and line_type = 'LINE'"
                              rsaux6.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux6.EOF Then
                                 'var_importe_total = IIf(IsNull(rsaux6!amount_applied), 0, rsaux6!amount_applied) + var_importe_total_df
                                 VAR_IMPORTE_TOTAL = IIf(IsNull(rsaux6!amount_applied), 0, rsaux6!amount_applied)
                                 If VAR_IMPORTE_TOTAL = 0 Then
                                    VAR_PORCENTAJE_FIN = 0
                                 Else
                                    'VAR_PORCENTAJE_FIN = 100 - (Round((IIf(IsNull(rsaux6!amount_applied), 0, rsaux6!amount_applied) * 100) / var_importe_total, 2))
                                    VAR_PORCENTAJE_FIN = (Round((IIf(IsNull(var_importe_total_df), 0, var_importe_total_df) * 100) / VAR_IMPORTE_TOTAL, 2))
                                 End If
                                 VAR_NOTA_CREDITO_DF = var_notas_credito_df
                                 var_precio = var_precio * (1 - (IIf(IsNull(VAR_PORCENTAJE_FIN), 0, VAR_PORCENTAJE_FIN) / 100))
                              End If
                              rsaux6.Close
                           End If
                           rsaux5.Close
                        Else
                           'MsgBox "SELECT PRODUCT_UOM_CODE, CURRENCY_CODE  * FROM qp_list_lines_v A, qp_list_HEADERS_v B WHERE A.list_header_id = B.list_header_id AND  A.list_header_id = " + CStr(var_clave_lista_precios) + " and  A.product_attr_value = " + CStr(rsaux9!INVENTORY_ITEM_ID) + " AND A.product_attr_val_disp = '" + rsaux9!CODIGO + "'"
                           If rsaux11.State = 1 Then
                              rsaux11.Close
                           End If
                           rsaux11.Open "SELECT PRODUCT_UOM_CODE, CURRENCY_CODE  FROM qp_list_lines_v A, qp_list_HEADERS_v B WHERE A.list_header_id = B.list_header_id AND  A.list_header_id = " + CStr(var_clave_lista_precios) + " and  A.product_attr_value = " + CStr(rsaux9!inventory_item_id) + " AND A.product_attr_val_disp = '" + rsaux9!codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           var_unidad_medida = rsaux11!PRODUCT_UOM_CODE
                           VAR_CURRENCY = rsaux11!CURRENCY_CODE
                           rsaux11.Close
                           VAR_ZZ = 0
                           If var_unidad_organizacional = "93" And VAR_ZZ = 1 Then
                              objConn.Open var_conexion_oracle
                              '� Establecer conexi�n a la base de datos con el objeto objConn.
                              With objCmd
                                   objConn.BeginTrans
                                   .ActiveConnection = objConn
                                   .CommandText = "APPS.xxvia_descuento_linea"
                                   .CommandType = adCmdStoredProc
                                
                                   rsaux10.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                                   If Not rsaux10.EOF Then
                                      var_responsabilidad_facturacion = IIf(IsNull(rsaux10!RESPONSABILIDAD_FACTURACION), "", rsaux10!RESPONSABILIDAD_FACTURACION)
                                   End If
                                   rsaux10.Close
                                         
                                         
                                   Set objParm = .CreateParameter("p_responsabilidad", adVarChar, adParamInput, 100, var_responsabilidad_facturacion)
                                   .Parameters.Append objParm
      
                                        
                                   Set objParm = .CreateParameter("p_org_id", adNumeric, adParamInput, 50, CDbl(var_empresa))
                                   .Parameters.Append objParm
                                   
                                   Set objParm = .CreateParameter("p_uom", adVarChar, adParamInput, 50, var_unidad_medida)
                                   .Parameters.Append objParm
                                
                                   Set objParm = .CreateParameter("p_currency", adVarChar, adParamInput, 50, VAR_CURRENCY)
                                   .Parameters.Append objParm
                                
                                   var_estatus_factura = ""
                                   Set objParm = .CreateParameter("p_inventory_item_id", adNumeric, adParamInput, 50, rsaux9!inventory_item_id)
                                   .Parameters.Append objParm
                                   
                                   Set objParm = .CreateParameter("p_price_list_id", adNumeric, adParamInput, 50, var_clave_lista_precios)
                                   .Parameters.Append objParm
                                   
                                   Set objParm = .CreateParameter("p_cust_account_id", adNumeric, adParamInput, 50, var_clave_titular)
                                   .Parameters.Append objParm
                                
                                   Set objParm = .CreateParameter("xx_unit_price", adNumeric, adParamOutput, 50, var_precio_inflado)
                                   .Parameters.Append objParm
                                   
                                   Set objParm = .CreateParameter("xx_adjusted_price", adNumeric, adParamOutput, 50, var_precio_descuento)
                                   .Parameters.Append objParm
                                   
                                   rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                   rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                   On Error GoTo salir2:
                                   .execute
                                 
                                   var_precio_inflado = .Parameters("xx_unit_price").Value
                                   var_precio_entero = var_precio_inflado
                                   var_precio_descuento = .Parameters("xx_adjusted_price").Value
                                   var_precio = var_precio_descuento
                                   objConn.CommitTrans
                              End With
                              Set objConn = Nothing
                              Set objCmd = Nothing
                           Else
                              x = 0
                              If x = 0 Then
                                 If rsaux10.State = 1 Then
                                    rsaux10.Close
                                 End If
                                 rsaux11.Open "select * from xxvia_system_items_b where segment1 = '" + rsaux9!codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux11.EOF Then
                                    rsaux10.Open "SELECT OPERAND FROM qp_list_lines_v WHERE list_header_id =  " + CStr(var_clave_lista_precios) + " AND  PRODUCT_ATTR_VAL_DISP = '" + rsaux9!codigo + "' and start_date_active <= sysdate and (end_date_active is null or end_date_active >= sysdate) and Product_Attr_Value = " + CStr(rsaux11!inventory_item_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    'rsaux10.Open "SELECT OPERAND FROM qp_list_lines_v WHERE list_header_id =  " + CStr(var_clave_lista_precios) + " AND  PRODUCT_ATTR_VAL_DISP = '" + rsaux9!codigo + "' AND Product_Attr_Value = " + CStr(rsaux11!INVENTORY_ITEM_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 Else
                                    rsaux10.Open "SELECT OPERAND FROM qp_list_lines_v WHERE list_header_id =  " + CStr(var_clave_lista_precios) + " AND  PRODUCT_ATTR_VAL_DISP = '" + rsaux9!codigo + "' and start_date_active <= sysdate and (end_date_active is null or end_date_active >= sysdate) and ", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux11.Close
                                 Product_Attr_Value = 16217
                                 var_precio_entero = IIf(IsNull(rsaux10!OPERAND), 0, rsaux10!OPERAND)
                                 var_precio = IIf(IsNull(rsaux10!OPERAND), 0, rsaux10!OPERAND)
                                 rsaux10.Close
                                 VAR_DESCUENTO = 0
                                 rsaux11.Open "SELECT DISTINCT(list_header_id) as calificador FROM qp_qualifiers_v WHERE list_header_id IS NOT NULL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 While Not rsaux11.EOF
                                       rsaux10.Open "select xxvia_fn_descuento_titular(" + CStr(rsaux11!calificador) + ",'" + Me.txt_titular + "') as descuento from dual ", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                       If Not rsaux10.EOF Then
                                          If rsaux10!DESCUENTO > VAR_DESCUENTO Then
                                             VAR_DESCUENTO = rsaux10!DESCUENTO
                                          End If
                                       End If
                                       rsaux10.Close
                                       rsaux11.MoveNext
                                 Wend
                                 rsaux11.Close
                                 var_precio = var_precio * (1 - (VAR_DESCUENTO / 100))
                              Else
                                 var_precio_entero = IIf(IsNull(rsaux9!Precio), 0, rsaux9!Precio)
                                 var_precio = IIf(IsNull(rsaux9!Precio), 0, rsaux9!Precio)
                              End If
                           End If
                       End If






                        
''''''
                        Else
                        rsaux.Close
                        rsaux.Open "SELECT L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id, SUM(NVL(L.GROSS_unit_selling_price,l.unit_selling_price)) AS PRECIO FROM RA_CUSTOMER_TRX_LINES_ALL L, RA_CUSTOMER_TRX_ALL E, ra_cust_trx_types_all TYPES Where TYPES.TYPE = 'INV' AND TYPES.cust_trx_type_id = E.cust_trx_type_id AND TYPES.org_id = E.org_id AND l.customer_trx_id = E.customer_trx_id AND L.inventory_item_id = " + CStr(rsaux9!inventory_item_id) + " AND E.sold_to_customer_id = " + CStr(rsaux9!TITULAR) + " and TRX_NUMBER not like 'F%' GROUP BY L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id  ORDER BY trx_date desc", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           var_veces = 1
                           If rsaux7.State = 1 Then
                              rsaux7.Close
                           End If
                           rsaux7.Open "select count(*) from RA_CUSTOMER_TRX_LINES_ALL where customer_trx_id= " + CStr(rsaux!customer_Trx_id) + " and inventory_item_id = " + CStr(rsaux9!inventory_item_id) + " and unit_selling_price >0", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux7.EOF Then
                              var_veces = IIf(IsNull(rsaux7(0).Value), 1, rsaux7(0).Value)
                           End If
                           rsaux7.Close
                           If rsaux!Precio = 0 Then
                              var_precio = 0
                           Else
                              var_precio = rsaux!Precio / var_veces
                           End If
                           
                           
                           If rsaux!Precio = 0 Then
                              var_precio_entero = 0
                           Else
                              var_precio_entero = rsaux!Precio / var_veces
                           End If
                           var_numero_factura = rsaux!customer_Trx_id
                           rsaux5.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           
                           
                           
                           If var_attribute10 = "0" Then
                              var_cadena = "SELECT ARPA.APPLIED_CUSTOMER_TRX_ID AS FACTURA_ID, ARPA.CUSTOMER_TRX_ID AS NOTA_CREDITO_ID, ARPA.ACCTD_AMOUNT_APPLIED_TO AS MONTO_APLICADO, RCT.CUST_TRX_TYPE_ID, RCTL.ATTRIBUTE11, RCTL.ATTRIBUTE10, ARPA.AMOUNT_APPLIED, acr.amount FROM AR_RECEIVABLE_APPLICATIONS_ALL ARPA, RA_CUSTOMER_TRX_ALL RCT, RA_CUSTOMER_TRX_LINES_ALL RCTL, ar_cash_receipts_all acr WHERE ARPA.APPLICATION_TYPE = 'CM' AND ARPA.CUSTOMER_TRX_ID = RCT.CUSTOMER_TRX_ID AND RCT.CUST_TRX_TYPE_ID IN (SELECT ATTRIBUTE2 From RA_CUST_TRX_TYPES_ALL WHERE ATTRIBUTE2 IS NOT NULL) AND ARPA.CUSTOMER_TRX_ID  = RCTL.CUSTOMER_TRX_ID AND RCTL.ATTRIBUTE11 IS NOT NULL AND ARPA.APPLIED_CUSTOMER_TRX_ID = " + CStr(var_numero_factura) + " and RCTL.ATTRIBUTE10 = acr.cash_receipt_id and ARPA.ACCTD_AMOUNT_APPLIED_TO > 0 order by arpa.last_update_date desc"
                              rsaux5.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              var_attribute10 = ""
                              If Not rsaux5.EOF Then
                                 While Not rsaux5.EOF
                                      If var_attribute10 = "" Then
                                         var_attribute10 = rsaux5!attribute10
                                      Else
                                         var_attribute10 = var_attribute10 + "," + rsaux5!attribute10
                                      End If
                                       rsaux5.MoveNext
                                 Wend
                              Else
                                 var_attribute10 = 0
                              End If
                              rsaux5.Close
                           End If
                           
                           
                           var_cadena = "select rec.CUSTOMER_TRX_ID, nvl(sum(rec.amount_applied),0) as importe_df from ar_receivable_applications_all rec Inner join ar_payment_schedules_all pay on rec.payment_schedule_id = pay.payment_schedule_id Inner join ra_cust_trx_types_all on pay.cust_trx_type_id = ra_cust_trx_types_all.cust_trx_type_id Where rec.applied_customer_trx_id = " + CStr(var_numero_factura) + " and rec.apply_date < sysdate and rec.display = 'Y' and application_type = 'CM' and ra_cust_trx_types_all.cust_trx_type_id in (1564,1028) group by rec.CUSTOMER_TRX_ID "
                           rsaux5.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux5.EOF Then
                              var_importe_total_df = 0
                              var_notas_credito_df = ""
                              While Not rsaux5.EOF
                                    var_importe_total_df = var_importe_total_df + IIf(IsNull(rsaux5!importe_df), 0, rsaux5!importe_df)
                                    If var_notas_credito_df = "" Then
                                       var_notas_credito_df = CStr(rsaux5!customer_Trx_id)
                                    Else
                                       var_notas_credito_df = var_notas_credito_df + ", " + CStr(rsaux5!customer_Trx_id)
                                    End If
                                    rsaux5.MoveNext
                              Wend
                              rsaux5.MoveFirst
                              'var_cadena = "select sum(amount_applied) amount_applied from ar_receivable_applications_all Where applied_customer_trx_id = " + CStr(VAR_NUMERO_fACTURA) + " and display = 'Y' and application_type = 'CASH' and cash_receipt_id in( " + CStr(var_attribute10) + ")"
                              var_cadena = "select SUM(nvl(gross_extended_amount, extended_amount)) AS amount_applied from ra_customer_trx_lines_all where customer_trx_id = " + CStr(var_numero_factura) + " and line_type = 'LINE'"
                              rsaux6.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux6.EOF Then
                                 'var_importe_total = IIf(IsNull(rsaux6!amount_applied), 0, rsaux6!amount_applied) + var_importe_total_df
                                 VAR_IMPORTE_TOTAL = IIf(IsNull(rsaux6!amount_applied), 0, rsaux6!amount_applied)
                                 If VAR_IMPORTE_TOTAL = 0 Then
                                    VAR_PORCENTAJE_FIN = 0
                                 Else
                                    'VAR_PORCENTAJE_FIN = 100 - (Round((IIf(IsNull(rsaux6!amount_applied), 0, rsaux6!amount_applied) * 100) / var_importe_total, 2))
                                    VAR_PORCENTAJE_FIN = (Round((IIf(IsNull(var_importe_total_df), 0, var_importe_total_df) * 100) / VAR_IMPORTE_TOTAL, 2))
                                 End If
                                 VAR_NOTA_CREDITO_DF = var_notas_credito_df
                                 var_precio = var_precio * (1 - (IIf(IsNull(VAR_PORCENTAJE_FIN), 0, VAR_PORCENTAJE_FIN) / 100))
                              End If
                              rsaux6.Close
                           End If
                           rsaux5.Close
                        Else
                           'MsgBox "SELECT PRODUCT_UOM_CODE, CURRENCY_CODE  * FROM qp_list_lines_v A, qp_list_HEADERS_v B WHERE A.list_header_id = B.list_header_id AND  A.list_header_id = " + CStr(var_clave_lista_precios) + " and  A.product_attr_value = " + CStr(rsaux9!INVENTORY_ITEM_ID) + " AND A.product_attr_val_disp = '" + rsaux9!CODIGO + "'"
                           If rsaux11.State = 1 Then
                              rsaux11.Close
                           End If
                           rsaux11.Open "SELECT PRODUCT_UOM_CODE, CURRENCY_CODE  FROM qp_list_lines_v A, qp_list_HEADERS_v B WHERE A.list_header_id = B.list_header_id AND  A.list_header_id = " + CStr(var_clave_lista_precios) + " and  A.product_attr_value = " + CStr(rsaux9!inventory_item_id) + " AND A.product_attr_val_disp = '" + rsaux9!codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           var_unidad_medida = rsaux11!PRODUCT_UOM_CODE
                           VAR_CURRENCY = rsaux11!CURRENCY_CODE
                           rsaux11.Close
                           VAR_ZZ = 0
                           If var_unidad_organizacional = "93" And VAR_ZZ = 1 Then
                              objConn.Open var_conexion_oracle
                              '� Establecer conexi�n a la base de datos con el objeto objConn.
                              With objCmd
                                   objConn.BeginTrans
                                   .ActiveConnection = objConn
                                   .CommandText = "APPS.xxvia_descuento_linea"
                                   .CommandType = adCmdStoredProc
                                
                                   rsaux10.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                                   If Not rsaux10.EOF Then
                                      var_responsabilidad_facturacion = IIf(IsNull(rsaux10!RESPONSABILIDAD_FACTURACION), "", rsaux10!RESPONSABILIDAD_FACTURACION)
                                   End If
                                   rsaux10.Close
                                         
                                         
                                   Set objParm = .CreateParameter("p_responsabilidad", adVarChar, adParamInput, 100, var_responsabilidad_facturacion)
                                   .Parameters.Append objParm
      
                                        
                                   Set objParm = .CreateParameter("p_org_id", adNumeric, adParamInput, 50, CDbl(var_empresa))
                                   .Parameters.Append objParm
                                   
                                   Set objParm = .CreateParameter("p_uom", adVarChar, adParamInput, 50, var_unidad_medida)
                                   .Parameters.Append objParm
                                
                                   Set objParm = .CreateParameter("p_currency", adVarChar, adParamInput, 50, VAR_CURRENCY)
                                   .Parameters.Append objParm
                                
                                   var_estatus_factura = ""
                                   Set objParm = .CreateParameter("p_inventory_item_id", adNumeric, adParamInput, 50, rsaux9!inventory_item_id)
                                   .Parameters.Append objParm
                                   
                                   Set objParm = .CreateParameter("p_price_list_id", adNumeric, adParamInput, 50, var_clave_lista_precios)
                                   .Parameters.Append objParm
                                   
                                   Set objParm = .CreateParameter("p_cust_account_id", adNumeric, adParamInput, 50, var_clave_titular)
                                   .Parameters.Append objParm
                                
                                   Set objParm = .CreateParameter("xx_unit_price", adNumeric, adParamOutput, 50, var_precio_inflado)
                                   .Parameters.Append objParm
                                   
                                   Set objParm = .CreateParameter("xx_adjusted_price", adNumeric, adParamOutput, 50, var_precio_descuento)
                                   .Parameters.Append objParm
                                   
                                   rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                   rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                   On Error GoTo salir2:
                                   .execute
                                 
                                   var_precio_inflado = .Parameters("xx_unit_price").Value
                                   var_precio_entero = var_precio_inflado
                                   var_precio_descuento = .Parameters("xx_adjusted_price").Value
                                   var_precio = var_precio_descuento
                                   objConn.CommitTrans
                              End With
                              Set objConn = Nothing
                              Set objCmd = Nothing
                           Else
                              x = 0
                              If x = 0 Then
                                 If rsaux10.State = 1 Then
                                    rsaux10.Close
                                 End If
                                 rsaux11.Open "select * from xxvia_system_items_b where segment1 = '" + rsaux9!codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux11.EOF Then
                                    rsaux10.Open "SELECT OPERAND FROM qp_list_lines_v WHERE list_header_id =  " + CStr(var_clave_lista_precios) + " AND  PRODUCT_ATTR_VAL_DISP = '" + rsaux9!codigo + "' and start_date_active <= sysdate and (end_date_active is null or end_date_active >= sysdate) and Product_Attr_Value = " + CStr(rsaux11!inventory_item_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    'rsaux10.Open "SELECT OPERAND FROM qp_list_lines_v WHERE list_header_id =  " + CStr(var_clave_lista_precios) + " AND  PRODUCT_ATTR_VAL_DISP = '" + rsaux9!codigo + "' AND Product_Attr_Value = " + CStr(rsaux11!INVENTORY_ITEM_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 Else
                                    rsaux10.Open "SELECT OPERAND FROM qp_list_lines_v WHERE list_header_id =  " + CStr(var_clave_lista_precios) + " AND  PRODUCT_ATTR_VAL_DISP = '" + rsaux9!codigo + "' and start_date_active <= sysdate and (end_date_active is null or end_date_active >= sysdate) and ", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux11.Close
                                 Product_Attr_Value = 16217
                                 var_precio_entero = IIf(IsNull(rsaux10!OPERAND), 0, rsaux10!OPERAND)
                                 var_precio = IIf(IsNull(rsaux10!OPERAND), 0, rsaux10!OPERAND)
                                 rsaux10.Close
                                 VAR_DESCUENTO = 0
                                 rsaux11.Open "SELECT DISTINCT(list_header_id) as calificador FROM qp_qualifiers_v WHERE list_header_id IS NOT NULL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 While Not rsaux11.EOF
                                 On Error GoTo desc:
                                       If rsaux11!calificador <> 1144154 Then
                                          If rsaux11!calificador <> 1778169 Then
                                             If Me.txt_titular <> "74667" Then
                                                var_cadena = "select xxvia_fn_descuento_titular(" + CStr(rsaux11!calificador) + ",'" + Me.txt_titular + "') as descuento from dual "
                                                rsaux10.Open "select xxvia_fn_descuento_titular(" + CStr(rsaux11!calificador) + ",'" + Me.txt_titular + "') as descuento from dual ", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                                If Not rsaux10.EOF Then
                                                       If rsaux10!DESCUENTO > VAR_DESCUENTO Then
                                                          VAR_DESCUENTO = rsaux10!DESCUENTO
                                                       End If
                                                End If
                                                rsaux10.Close
                                             End If
                                          End If
                                       End If
desc:
                                       rsaux11.MoveNext
                                 Wend
                                 rsaux11.Close
                                 var_precio = var_precio * (1 - (VAR_DESCUENTO / 100))
                              Else
                                 var_precio_entero = IIf(IsNull(rsaux9!Precio), 0, rsaux9!Precio)
                                 var_precio = IIf(IsNull(rsaux9!Precio), 0, rsaux9!Precio)
                              End If
                           End If
                        End If
                        End If 'PARA QUE TOME EN CUENTA EL PRECIO VIEJO DE ARTICULOS QUE CAMBIARON DE PRECIO
                        'quitar cuando se termine lod el cambio de cantia
                        If rsaux.State = 1 Then
                           rsaux.Close
                        End If
                        rsaux.Open "UPDATE XXVIA_TB_DEVOLUCIONES_CLIENTES SET PRECIO = " + CStr(var_precio) + ",FACTURA = '" + CStr(var_numero_factura) + "', DESCUENTO_FINANCIERO = '" + CStr(IIf(IsNull(VAR_PORCENTAJE_FIN), 0, VAR_PORCENTAJE_FIN)) + "', NOTA_CREDITO_DESC_FIN = '" + Mid(CStr(VAR_NOTA_CREDITO_DF), 1, 100) + "', precio_entero = " + CStr(var_precio_entero) + " WHERE NUMERO = " + Me.txt_folio + " AND ORGANIZACION = " + var_unidad_organizacional + " AND INVENTORY_ITEM_ID = " + CStr(rsaux9!inventory_item_id) + " AND MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        var_cantidad = rsaux9!cantidad
                        x = 0
                        If x = 0 Then
                           While var_cantidad > 0
                                 VAR_CANTIDAD_RESTANTE = var_cantidad - 1
                                 If VAR_CANTIDAD_RESTANTE >= 1 Then
                                    var_cantidad_LEER = 1
                                    var_cantidad = var_cantidad - 1
                                 Else
                                    var_cantidad_LEER = var_cantidad
                                    var_cantidad = 0
                                 End If
                                 If Mid(Me.txt_referencia, 1, 9) = "DC BULTOS" Then
                                    var_devolucion_costales = 1
                                 Else
                                    var_devolucion_costales = 0
                                 End If
                                 var_consecutivo = var_consecutivo + 1
                                 If var_devolucion_costales = 1 Then
                                    rsaux10.Open "INSERT INTO XXVIA_TB_dEV_CLIENTES_DESGLOCE (NUMERO, ORGANIZACION, INVENTORY_ITEM_ID, CANTIDAD, CAUSA_DEVOLUCION, CONSECUTIVO, DESCRIPCION_CAUSA, ESTATUS, CODIGO, DESCRIPCION, LOCALIZADOR, MOVIMIENTO, TIPO_PEDIDO) VALUES (" + Me.txt_folio + "," + var_unidad_organizacional + "," + CStr(rsaux9!inventory_item_id) + "," + CStr(var_cantidad_LEER) + ",'210'," + CStr(var_consecutivo) + ",'NEGOCIACION ESPECIAL','','" + rsaux9!codigo + "','" + rsaux9!Descripcion + "','" + IIf(IsNull(rsaux9!localizador), "", rsaux9!localizador) + "','" + var_clave_movimiento + "'," + CStr(IIf(IsNull(rsaux9!tipo_pedido), 0, rsaux9!tipo_pedido)) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 Else
                                    rsaux10.Open "INSERT INTO XXVIA_TB_dEV_CLIENTES_DESGLOCE (NUMERO, ORGANIZACION, INVENTORY_ITEM_ID, CANTIDAD, CAUSA_DEVOLUCION, CONSECUTIVO, DESCRIPCION_CAUSA, ESTATUS, CODIGO, DESCRIPCION, LOCALIZADOR, MOVIMIENTO, TIPO_PEDIDO) VALUES (" + Me.txt_folio + "," + var_unidad_organizacional + "," + CStr(rsaux9!inventory_item_id) + "," + CStr(var_cantidad_LEER) + ",''," + CStr(var_consecutivo) + ",'','','" + rsaux9!codigo + "','" + rsaux9!Descripcion + "','" + IIf(IsNull(rsaux9!localizador), "", rsaux9!localizador) + "','" + var_clave_movimiento + "'," + CStr(IIf(IsNull(rsaux9!tipo_pedido), 0, rsaux9!tipo_pedido)) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 End If
                           Wend
                        Else
                             If Mid(Me.txt_referencia, 1, 9) = "DC BULTOS" Then
                                var_devolucion_costales = 1
                             Else
                                var_devolucion_costales = 0
                             End If
                             
                             var_consecutivo = var_consecutivo + 1
                             var_cantidad_LEER = var_cantidad
                             If var_devolucion_costales = 1 Then
                                rsaux10.Open "INSERT INTO XXVIA_TB_dEV_CLIENTES_DESGLOCE (NUMERO, ORGANIZACION, INVENTORY_ITEM_ID, CANTIDAD, CAUSA_DEVOLUCION, CONSECUTIVO, DESCRIPCION_CAUSA, ESTATUS, CODIGO, DESCRIPCION, LOCALIZADOR, MOVIMIENTO, TIPO_PEDIDO) VALUES (" + Me.txt_folio + "," + var_unidad_organizacional + "," + CStr(rsaux9!inventory_item_id) + "," + CStr(var_cantidad_LEER) + ",'210'," + CStr(var_consecutivo) + ",'NEGOCIACION ESPECIAL','','" + rsaux9!codigo + "','" + rsaux9!Descripcion + "','" + IIf(IsNull(rsaux9!localizador), "", rsaux9!localizador) + "','" + var_clave_movimiento + "'," + CStr(IIf(IsNull(rsaux9!tipo_pedido), 0, rsaux9!tipo_pedido)) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                             Else
                                rsaux10.Open "INSERT INTO XXVIA_TB_dEV_CLIENTES_DESGLOCE (NUMERO, ORGANIZACION, INVENTORY_ITEM_ID, CANTIDAD, CAUSA_DEVOLUCION, CONSECUTIVO, DESCRIPCION_CAUSA, ESTATUS, CODIGO, DESCRIPCION, LOCALIZADOR, MOVIMIENTO, TIPO_PEDIDO) VALUES (" + Me.txt_folio + "," + var_unidad_organizacional + "," + CStr(rsaux9!inventory_item_id) + "," + CStr(var_cantidad_LEER) + ",''," + CStr(var_consecutivo) + ",'','','" + rsaux9!codigo + "','" + rsaux9!Descripcion + "','" + IIf(IsNull(rsaux9!localizador), "", rsaux9!localizador) + "','" + var_clave_movimiento + "'," + CStr(IIf(IsNull(rsaux9!tipo_pedido), 0, rsaux9!tipo_pedido)) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                             End If
                        End If
                        rsaux9.MoveNext
                  Wend
                  rsaux9.Close
                  var_numero_folio_devoluciones = CDbl(Me.txt_folio)
                  var_clave_almacen_devolucion = Me.txt_almacen
                  var_referencia_global_dev = Me.txt_referencia
                  
                  
                  
                  
                  
                  frmoracle_devoluciones_desgloce.Show 1
                  txt_codigo.Enabled = False
                  txt_foco.Enabled = False
                  var_devolucion_costales = 0
               Else
                  If IIf(IsNull(rsaux8!estatus), "", rsaux8!estatus) = "" Then
                     var_numero_folio_devoluciones = CDbl(Me.txt_folio)
                     frmoracle_devoluciones_desgloce.Show 1
                  Else
                     var_numero_folio_devoluciones = CDbl(Me.txt_folio)
                     frmoracle_devoluciones_desgloce.Show 1
                  End If
               End If
               If rsaux8.State = 1 Then
                  rsaux8.Close
               End If
            Else
               MsgBox "Los siguientes c�digos no son retornables " + var_cadena_codigos_retornables, vbOKOnly, "ATENCION"
            End If
         End If
      Else
         If var_clave_movimiento = "VDI" Or var_clave_movimiento = "SML" Then
            var_numero_folio_devoluciones = CDbl(Me.txt_folio)
            'var_estatus_movimiento = ""
            If var_estatus_movimiento = "I" Then
               rsaux.Open "select order_number from oe_order_headers_all where orig_sys_document_ref = 'SIDVDI_" + Trim(CStr(var_numero_folio_devoluciones)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_pedido = rsaux(0).Value
               rsaux.Close
               rsaux.Open "select order_number from oe_order_headers_all where orig_sys_document_ref = 'SIDVDI_" + Trim(CStr(var_numero_folio_devoluciones)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_pedido = rsaux(0).Value
               rsaux.Close
               MsgBox "El pedido generado es el " + CStr(var_pedido), vbOKOnly, "ATENCION"
               x = 0
               If x = 0 Then
               var_encontro = 0
               'var_cadena = "SELECT RCT.CUSTOMER_TRX_ID, APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, APS.STATUS, APS.CLASS, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME From hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND RCT.BILL_TO_SITE_USE_ID = HCSU.SITE_USE_ID AND INTERFACE_HEADER_ATTRIBUTE1 IN ('" + CStr(var_pedido) + "') AND RCT.customer_trx_id = APS.customer_trx_id "
               var_cadena = "SELECT * FROM RA_CUSTOMER_TRX_ALL WHERE INTERFACE_HEADER_ATTRIBUTE1 = '" + CStr(var_pedido) + "' AND INTERFACE_HEADER_CONTEXT = 'ORDER ENTRY'"
               rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  var_encontros = 1
                  var_customer_trx_id = rsaux2!customer_Trx_id
                  var_factura = rsaux2!trx_number
                  rs.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_serie_pedido = IIf(IsNull(rs!Serie), "", rs!Serie)
                     var_ruta_facturas = IIf(IsNull(rs!ruta_facturas), "", rs!ruta_facturas)
                  End If
                  rs.Close
                  rsaux1.Open "select customer_trx_id from xxvia_Tb_control_doc_fiscales where customer_trx_id = " + CStr(var_customer_trx_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  var_posible = 0
                  If rsaux1.EOF Then
                     var_posible = 1
                  Else
                     rsaux10.Open "CALL XXVIA_SEND_POST('" + CStr(rsaux1!customer_Trx_id) + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux1.Close
                  If var_posible = 0 Then
                     URL = "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=VTH981105F90&serie=" + Trim(var_serie_pedido) + "&folio=" + Trim(CStr(var_factura))
                     buf = Split(URL, ".")
                     ext = buf(UBound(buf))
                     strSavePath = "C:\SISTEMAS\" + Trim(var_serie_pedido) + Trim(CStr(var_factura)) + ".pdf"
                     ret = URLDownloadToFile(0, URL, strSavePath, 0, 0)
                     If ret = 0 Then
                        Call ShellExecute(Me.hwnd, "print", "C:\SISTEMAS\" + Trim(var_serie_pedido) + Trim(CStr(var_factura)) + ".PDF", vbNullString, vbNullString, SW_SHOWNORMAL = 1)
                     Else
                        MsgBox "Error en la factura " + Trim(var_serie_pedido) + Trim(CStr(var_factura))
                     End If
                  Else
                     MsgBox "No se a generado la factura", vbOKOnly, "ATENCION"
                  End If
                  
                  'Open (App.Path & "\EJPDF" + Trim(var_serie_pedido) + Trim(CStr(var_factura)) + ".bat") For Output As #2
                  'Print #2, "START " + var_ruta_facturas + var_serie_pedido + "\"; Trim(var_serie_pedido) + Trim(CStr(var_factura)) + ".PDF"
                  'Close #2
                  'var_Archivo = App.Path & "\EJPDF" + Trim(var_serie_pedido) + Trim(CStr(var_factura)) + ".bat"
                  'x = Shell(var_Archivo, vbHide)
               Else
                  MsgBox "No se a generado la factura", vbOKOnly, "ATENCION"
               End If
               rsaux2.Close
               End If
            Else
               If Me.txt_almacen = "TEX_VB1" Or Me.txt_almacen = "TEX_VB2" Or Me.txt_almacen = "TEX_VB4" Then
                  If Me.txt_almacen = "TEX_VB1" Then
                     var_tipo_pedido_ventas_directas = 1121
                  End If
                  If Me.txt_almacen = "TEX_VB2" Then
                     var_tipo_pedido_ventas_directas = 1123
                  End If
                  If Me.txt_almacen = "TEX_VB4" Then
                     var_tipo_pedido_ventas_directas = 1441
                  End If
               Else

                  If var_clave_usuario_global = "U0000001250" Then
                     var_tipo_pedido_ventas_directas = 2221
                  Else
                     frmoracle_tipo_pedido.Show 1
                  End If
               End If
               'var_tipo_pedido_ventas_directas = 1184
               'var_tipo_pedido_ventas_directas = 1048
               'var_tipo_pedido_ventas_directas = 1042}
               'var_tipo_pedido_ventas_directas = 1464
               rsaux7.Open "select name from qp_secu_list_headers_v where list_header_id = " + CStr(var_clave_lista_precios), cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_lista_precios = rsaux7!Name
               rsaux7.Close
               'var_lista_precios = var_clave_lista_precios
               var_clave_tipo_pedido = var_tipo_pedido_ventas_directas
               'var_clave_tipo_pedido = 1106
               
               If var_clave_tipo_pedido > 0 Then
                  If var_lista_precios <> "" Then
                     If rs.State = 1 Then
                        rs.Close
                     End If
                     rs.Open "SELECT A.ESTATUS, A.MOVIMIENTO, A.NUMERO, A.ORGANIZACION, A.inventory_item_id, A.almacen, A.titular, A.unidad_medida, A.precio, A.TITULAR, A.CLIENTE, A.ESTABLECIMIENTO, A.LOCALIZADOR, a.codigo, SUM(A.cantidad) AS CANTIDAD FROM XXVIA_TB_DEVOLUCIONES_CLIENTES A WHERE A.NUMERO = " + CStr(var_numero_folio_devoluciones) + " AND A.ORGANIZACION = " + var_unidad_organizacional + "  AND A.MOVIMIENTO = '" + var_clave_movimiento + "' and A.cantidad > 0 GROUP BY A.ESTATUS, A.MOVIMIENTO, A.NUMERO, A.ORGANIZACION, A.inventory_item_id, A.almacen, A.titular, A.unidad_medida, A.precio, A.TITULAR, A.CLIENTE, A.ESTABLECIMIENTO, A.LOCALIZADOR, a.codigo", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     
                     If IIf(IsNull(rs!estatus), "", rs!estatus) = "" Then
                        If Not rs.EOF Then
                           rs.MoveFirst
                           var_cadena_posible_existencias = ""
                           GoTo x:
                           While Not rs.EOF
                                 'Me.txt_almacen = "PRIVALIA"
                                 strconsulta = "select * from Xxvia_vw_existencias_inv where organization_id = ? and subinventory_code = ? and segment1 = ?"
                                 With comandoORA
                                      .ActiveConnection = cnnoracle_4
                                      .CommandType = adCmdText
                                      .CommandText = strconsulta
                                      Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_unidad_organizacional))
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_almacen)
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!codigo)
                                      .Parameters.Append parametro
                                 End With
                                 Set rsaux = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing
                           
                                 'rsaux.Open "select * from Xxvia_vw_existencias_inv where organization_id = " + var_unidad_organizacional + " and subinventory_code = '" + Me.txt_almacen + "' and segment1 = '" + rs!CODIGO + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux.EOF Then
                                    var_disponible = IIf(IsNull(rsaux!Disponible), 0, rsaux!Disponible)
                                 Else
                                    var_disponible = 0
                                 End If
                                 If var_disponible < rs!cantidad Then
                                    If var_cadena_posible_existencias = "" Then
                                       rsaux10.Open "insert into codigos_faltantes (codigo, cantidad, disponible) values ('" + rs!codigo + "'," + CStr(rs!cantidad) + "," + CStr(var_disponible) + ")", cnn, adOpenDynamic, adLockOptimistic
                                       var_cadena_posible_existencias = "Movimiento " + CStr(rs!codigo) + " = " + CStr(rs!cantidad) + " Disponible = " + CStr(var_disponible)
                                    Else
                                       rsaux10.Open "insert into codigos_faltantes (codigo, cantidad, disponible) values ('" + rs!codigo + "'," + CStr(rs!cantidad) + "," + CStr(var_disponible) + ")", cnn, adOpenDynamic, adLockOptimistic
                                       var_cadena_posible_existencias = var_cadena_posible_existencias + "--- Movimiento " + CStr(rs!codigo) + " =  " + CStr(rs!cantidad) + " Disponible = " + CStr(var_disponible)
                                    End If
                                 End If
                                 rsaux.Close
                                 rs.MoveNext
                           Wend
                           rs.MoveFirst
                           If (var_unidad_organizacional = 85 Or var_unidad_organizacional = 94) And Me.txt_cliente = "307953" Then
                              var_cadena_posible_existencias = ""
                           End If
x:
'var_cadena_posible_existencias = ""
                           
                           If var_clave_usuario_global = "U0000000763" Or var_clave_usuario_global = "U0000001250" Then
                              var_cadena_posible_existencias = ""
                           End If
                           If var_cadena_posible_existencias = "" Then
                              
                              
                             If rsaux12.State = 1 Then
                                rsaux12.Close
                             End If


                              rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              'var_numero_folio_devoluciones = CDbl(Me.txt_folio)
                              If (var_unidad_organizacional = 85 Or var_unidad_organizacional = 94) And Me.txt_cliente = "307953" Then
                                 If rsaux.State = 1 Then
                                    rsaux.Close
                                 End If
                                 rsaux.Open "CALL  XXVIA_SP_JOB_BATCH_SID (" + CStr(var_unidad_organizacional) + "," + Me.txt_folio.Text + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 var_posible_existencias = 0
                                 While var_posible_existencias = 0
                                       For var_i = 1 To Me.lv_entradas.ListItems.Count
                                           Me.lv_entradas.ListItems.Item(var_i).Selected = True
                                           var_codigo = Me.lv_entradas.selectedItem
                                           strconsulta = "select * from Xxvia_vw_existencias_inv where organization_id = ? and subinventory_code = ? and segment1 = ?"
                                           With comandoORA
                                                .ActiveConnection = cnnoracle_4
                                                .CommandType = adCmdText
                                                .CommandText = strconsulta
                                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_unidad_organizacional))
                                                .Parameters.Append parametro
                                                txt_almacen = "PRODTER"
                                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_almacen)
                                                .Parameters.Append parametro
                                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(var_codigo))
                                                .Parameters.Append parametro
                                           End With
                                           Set rsaux12 = comandoORA.execute
                                           Set comandoORA = Nothing
                                           Set parametro = Nothing
                                           If Not rsaux12.EOF Then
                                              var_cantidad_pedida = CDbl(Me.lv_entradas.selectedItem.SubItems(2))
                                              var_cantidad = IIf(IsNull(rsaux12!Disponible), 0, rsaux12!Disponible)
                                              If var_cantidad < var_cantidad_pedida Then
                                                 var_posible_existencias = 0
                                                 var_i = 0
                                                 strconsulta = "select * from xxvia_tb_job_batch_log where vcha_transaction_reference = ?"
                                                 With comandoORA
                                                      .ActiveConnection = cnnoracle_4
                                                      .CommandType = adCmdText
                                                      .CommandText = strconsulta
                                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(Me.txt_folio))
                                                      .Parameters.Append parametro
                                                End With
                                                Set rsaux12 = comandoORA.execute
                                                Set comandoORA = Nothing
                                                Set parametro = Nothing
                                                var_log = ""
                                                If Not rsaux12.EOF Then
                                                   var_log = IIf(IsNull(rsaux12!VCHA_MENSAJE_MENSAJE), "", rsaux12VCHA_MENSAJE_MENSAJE)
                                                   If var_log <> "" Then
                                                      GoTo no_job
                                                   End If
                                                End If
                                                rsaux12.Close
                                              Else
                                                 var_posible_existencias = 1
                                              End If
                                           Else
                                              var_posible_existencias = 0
                                              var_i = 0
                                                 strconsulta = "select * from xxvia_tb_job_batch_log where vcha_transaction_reference = ?"
                                                 With comandoORA
                                                      .ActiveConnection = cnnoracle_4
                                                      .CommandType = adCmdText
                                                      .CommandText = strconsulta
                                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(Me.txt_folio))
                                                      .Parameters.Append parametro
                                                End With
                                                Set rsaux12 = comandoORA.execute
                                                Set comandoORA = Nothing
                                                Set parametro = Nothing
                                                var_log = ""
                                                If Not rsaux12.EOF Then
                                                   var_log = IIf(IsNull(rsaux12!VCHA_MENSAJE_MENSAJE), "", rsaux12VCHA_MENSAJE_MENSAJE)
                                                   If var_log <> "" Then
                                                      GoTo no_job
                                                   End If
                                                End If
                                                rsaux12.Close
                                           
                                           End If
                                           If rsaux12.State = 1 Then
                                           rsaux12.Close
                                           End If
                                       Next var_i
                                 Wend
                              End If
                              
                              
                              
                              
                              
                              
                              If Me.txt_almacen = "TEX_VB1" Or Me.txt_almacen = "TEX_VB2" Or Me.txt_almacen = "TEX_VB4" Then
                                 rsaux11.Open "select * from tb_oracle_pedidos_vb where almacen = '" + Me.txt_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux11.EOF Then
                                    var_pedido = IIf(IsNull(rsaux11!pedido), 100000, rsaux11!pedido)
                                    rsaux12.Open "update tb_oracle_pedidos_vb set pedido = " + CStr(var_pedido + 1), cnn, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux11.Close
                                 var_cadena = "INSERT INTO oe_headers_iface_all (order_number,ORDER_SOURCE_ID, orig_sys_document_ref, creation_date, created_by, last_update_date, last_updated_by, operation_code , sold_to_org_id        , SHIP_TO_ORG_id                   ,INVOICE_TO_ORG_ID     , Order_type_ID, PRICE_LIST, org_id, ship_from_org_id)"
                                 var_cadena = var_cadena + "  VALUES ('" + CStr(var_pedido) + "',1001,'SIDVDI_" + Trim(CStr((var_numero_folio_devoluciones))) + "',SYSDATE,-1,SYSDATE, -1,'INSERT', " + CStr(rs!TITULAR) + "," + CStr(rs!establecimiento) + "," + CStr(rs!Cliente) + "," + CStr(var_clave_tipo_pedido) + ",'" + var_lista_precios + "'," + var_empresa + "," + var_unidad_organizacional + ")"
                              Else
                                 If var_clave_usuario_global = "U0000000763" Or var_clave_usuario_global = "U0000001250" Then
                                 'SALESREP_ID
                                    If CDbl(var_clave_tipo_pedido) = 2061 Then
                                       var_cadena = "INSERT INTO oe_headers_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref, creation_date, created_by, last_update_date, last_updated_by, operation_code , sold_to_org_id        , SHIP_TO_ORG_id                   ,INVOICE_TO_ORG_ID     , Order_type_ID, PRICE_LIST, org_id, ship_from_org_id, SALESREP_ID, PAYMENT_TERM_ID, payment_term)"
                                       var_cadena = var_cadena + "  VALUES (1001,'SIDVDI_" + Trim(CStr((var_numero_folio_devoluciones))) + "',SYSDATE,-1,SYSDATE, -1,'INSERT', " + CStr(rs!TITULAR) + "," + CStr(rs!establecimiento) + "," + CStr(rs!Cliente) + "," + CStr(var_clave_tipo_pedido) + ",'" + var_lista_precios + "'," + var_empresa + "," + var_unidad_organizacional + ",100076045,38441,'PLAZO_ESPECIAL')"
                                    Else
                                       If var_clave_usuario_global = "U0000001250" Then
                                          var_cadena = "INSERT INTO oe_headers_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref, creation_date, created_by, last_update_date, last_updated_by, operation_code , sold_to_org_id        , SHIP_TO_ORG_id                   ,INVOICE_TO_ORG_ID     , Order_type_ID, PRICE_LIST, org_id, ship_from_org_id, SALESREP_ID)"
                                          var_cadena = var_cadena + "  VALUES (1001,'SID" + var_clave_movimiento + "_" + Trim(CStr((var_numero_folio_devoluciones))) + "',SYSDATE,-1,SYSDATE, -1,'INSERT', " + CStr(rs!TITULAR) + "," + CStr(rs!establecimiento) + "," + CStr(rs!Cliente) + "," + CStr(var_clave_tipo_pedido) + ",'" + var_lista_precios + "'," + var_empresa + "," + var_unidad_organizacional + ",100067042)"
                                       Else
                                          var_cadena = "INSERT INTO oe_headers_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref, creation_date, created_by, last_update_date, last_updated_by, operation_code , sold_to_org_id        , SHIP_TO_ORG_id                   ,INVOICE_TO_ORG_ID     , Order_type_ID, PRICE_LIST, org_id, ship_from_org_id, SALESREP_ID)"
                                          var_cadena = var_cadena + "  VALUES (1001,'SIDVDI_" + Trim(CStr((var_numero_folio_devoluciones))) + "',SYSDATE,-1,SYSDATE, -1,'INSERT', " + CStr(rs!TITULAR) + "," + CStr(rs!establecimiento) + "," + CStr(rs!Cliente) + "," + CStr(var_clave_tipo_pedido) + ",'" + var_lista_precios + "'," + var_empresa + "," + var_unidad_organizacional + ",100076045)"
                                       End If
                                    End If
                                 Else
                                    var_cadena = "INSERT INTO oe_headers_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref, creation_date, created_by, last_update_date, last_updated_by, operation_code , sold_to_org_id        , SHIP_TO_ORG_id                   ,INVOICE_TO_ORG_ID     , Order_type_ID, PRICE_LIST, org_id, ship_from_org_id)"
                                    var_cadena = var_cadena + "  VALUES (1001,'SIDVDI_" + Trim(CStr((var_numero_folio_devoluciones))) + "',SYSDATE,-1,SYSDATE, -1,'INSERT', " + CStr(rs!TITULAR) + "," + CStr(rs!establecimiento) + "," + CStr(rs!Cliente) + "," + CStr(var_clave_tipo_pedido) + ",'" + CStr(var_lista_precios) + "'," + var_empresa + "," + var_unidad_organizacional + ")"
                                 End If
                              End If
                              If rsaux.State = 1 Then
                                 rsaux.Close
                              End If
                              rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              var_i = 0
                              While Not rs.EOF
                                    var_i = var_i + 1
                                    rsaux10.Open "SELECT PRIMARY_UOM_CODE FROM xxvia_system_items_b WHERE INVENTORY_ITEM_ID = " + CStr(rs!inventory_item_id) + " AND ORGANIZATION_ID = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    If Not rsaux10.EOF Then
                                       VAR_MEDIDA = rsaux10(0).Value
                                    End If
                                    rsaux10.Close
                                    If Me.txt_almacen = "TEX_VB1" Or Me.txt_almacen = "TEX_VB2" Or Me.txt_almacen = "TEX_VB4" Then
                                       var_cadena = "INSERT INTO oe_lines_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref,orig_sys_line_ref,inventory_item_id,ordered_quantity, operation_code, created_by, creation_date, last_updated_by, last_update_date, unit_selling_price, unit_list_price, calculate_price_flag, PRICING_QUANTITY, PRICING_QUANTITY_UOM, ATTRIBUTE1, subinventory, org_id, ship_from_org_id)"
                                       var_cadena = var_cadena + " VALUES (1001,'SIDVDI_" + Trim(CStr(var_numero_folio_devoluciones)) + "','" + CStr(var_i) + "', " + CStr(rs!inventory_item_id) + ", " + CStr(rs!cantidad) + ",'INSERT', -1,SYSDATE, -1,SYSDATE," + CStr(rs!Precio) + "," + CStr(rs!Precio) + ",'Y', " + CStr(rs!cantidad) + ", '" + VAR_MEDIDA + "','" + IIf(IsNull(rs!localizador), "", rs!localizador) + "','" + Me.txt_almacen + "'," + var_empresa + "," + var_unidad_organizacional + ")"
                                    Else
                                       var_cadena = "INSERT INTO oe_lines_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref,orig_sys_line_ref,inventory_item_id,ordered_quantity, operation_code, created_by, creation_date, last_updated_by, last_update_date, unit_selling_price, unit_list_price, calculate_price_flag, PRICING_QUANTITY, PRICING_QUANTITY_UOM, ATTRIBUTE1, subinventory, org_id, ship_from_org_id)"
                                       var_cadena = var_cadena + " VALUES (1001,'SID" + var_clave_movimiento + "_" + Trim(CStr(var_numero_folio_devoluciones)) + "','" + CStr(var_i) + "', " + CStr(rs!inventory_item_id) + ", " + CStr(rs!cantidad) + ",'INSERT', -1,SYSDATE, -1,SYSDATE," + CStr(rs!Precio) + "," + CStr(rs!Precio) + ",'Y', " + CStr(rs!cantidad) + ", '" + VAR_MEDIDA + "','" + IIf(IsNull(rs!localizador), "", rs!localizador) + "','" + Me.txt_almacen + "'," + var_empresa + "," + var_unidad_organizacional + ")"
                                    End If
                                    rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    rs.MoveNext
                              Wend
                              On Error GoTo salir2
                              If Me.txt_almacen = "TEX_VB1" Or Me.txt_almacen = "TEX_VB2" Or Me.txt_almacen = "TEX_VB4" Then
                                 rsaux.Open "INSERT INTO oe_actions_iface_all (order_source_ID, orig_sys_document_ref, operation_code) VALUES (1001, 'SIDVDI_" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "','BOOK_ORDER')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              Else
                                 rsaux.Open "INSERT INTO oe_actions_iface_all (order_source_ID, orig_sys_document_ref, operation_code) VALUES (1001, 'SID" + var_clave_movimiento + "_" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "','BOOK_ORDER')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              End If
                              
                              
                              
                              
                              
                              rsaux.Open "CALL XXVIA_PK_INTERFACES_OM.importar_pedido('SID" + var_clave_movimiento + "_" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "'," + var_empresa + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If rsaux.State = 1 Then
                                 rsaux.Close
                              End If
                              rsaux.Open "UPDATE XXVIA_TB_DEVOLUCIONES_CLIENTES A SET ESTATUS = 'I' WHERE A.NUMERO = " + CStr(var_numero_folio_devoluciones) + " AND A.ORGANIZACION = " + var_unidad_organizacional + "  AND A.MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              rsaux.Open "select order_number from oe_order_headers_all where orig_sys_document_ref = 'SID" + var_clave_movimiento + "_" + Trim(CStr(var_numero_folio_devoluciones)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              var_pedido = rsaux(0).Value
                              rsaux.Close
                              rsaux.Open "select * from tb_oracle_pedidos_cerrados where pedido = " + CStr(var_pedido), cnn, adOpenDynamic, adLockOptimistic
                              If rsaux.EOF Then
                                 rsaux10.Open "insert into tb_oracle_pedidos_Cerrados(PEDIDO, REQUEST_ID) VALUES (" + CStr(var_pedido) + ",0)", cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux.Close
                              
                              MsgBox "El pedido generado es el " + CStr(var_pedido), vbOKOnly, "ATENCION"
                              Me.txt_codigo.Enabled = False
                              
                              
                              
                              
                              
                              
                              
                              x = 0
                              If x = 1 Then
                              While var_encontro = 0
                                    var_cadena = "SELECT * FROM RA_INTERFACE_LINES_ALL WHERE SALES_ORDER = " + CStr(var_pedido)
                                    rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    If Not rsaux.EOF Then
                                       var_encontro = 1
                                    End If
                                    rsaux.Close
                              Wend
                              clnt.MSSoapInit var_webservice
                              For var_j = 1 To 2
                                  var_con = clnt.ejecutar_autoinvoice("OM_FACTURAS", 4002)
                              Next var_j
                              Set clint = Nothing
                           
                              
                              rsaux.Open "SELECT oha.header_id, oha.ordered_date, oha.order_number,  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  E.NAME, f.orig_system_reference from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors E, hz_cust_acct_sites_all f Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id and HCSU.site_use_code = 'BILL_TO' and f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID and order_number  = '" + CStr(var_pedido) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux.EOF Then
                                 'var_cadena = "SELECT e.collector_id, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  E.NAME , sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors e Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id "
                                 'var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + CStr(var_pedido) + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id  group by  e.collector_id, A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status,  E.NAME"
                                 'rsaux1.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 'If Not rsaux1.EOF Then
                                    var_encontros = 0
                                    VAR_Z = 0
                                    While var_encontros = 0
                                          If VAR_Z = 1000 Then
                                             VAR_Z = 0
                                          End If
                                          'MsgBox VAR_PEDIDO
                                          var_cadena = "SELECT RCT.CUSTOMER_TRX_ID, APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, APS.STATUS, APS.CLASS, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME From hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND RCT.BILL_TO_SITE_USE_ID = HCSU.SITE_USE_ID AND INTERFACE_HEADER_ATTRIBUTE1 IN ('" + CStr(var_pedido) + "') AND RCT.customer_trx_id = APS.customer_trx_id "
                                          rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          If Not rsaux2.EOF Then
                                             var_encontros = 1
                                             var_customer_trx_id = rsaux2!customer_Trx_id
                                             var_factura = rsaux2!trx_number
                                          End If
                                          rsaux2.Close
                                          VAR_Z = VAR_Z + 1
                                    Wend
                                    x = 0
                                    If x = 1 Then
                                    var_cadena = "SELECT RCT.CUSTOMER_TRX_ID, APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, APS.STATUS, APS.CLASS, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  E.COLLECTOR_ID, E.NAME From hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS, ar_collectors E, hz_customer_profiles D Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND RCT.BILL_TO_SITE_USE_ID = HCSU.SITE_USE_ID AND INTERFACE_HEADER_ATTRIBUTE1 IN ('" + CStr(var_pedido) + "') AND RCT.customer_trx_id = APS.customer_trx_id AND E.collector_id = D.COLLECTOR_ID AND D.site_use_id = HCSU.SITE_USE_ID "
                                    rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    While Not rsaux2.EOF
                                          var_customer_trx_id = rsaux2!customer_Trx_id
                                          If objConn.State = 1 Then
                                             objConn.RollbackTrans
                                             objConn.Close
                                          End If
                                          objConn.Open var_conexion_oracle
                                          '� Establecer conexi�n a la base de datos con el objeto objConn.
                                          With objCmd
                                               objConn.BeginTrans
                                               .ActiveConnection = objConn
                                               'LISTO
                                               .CommandText = "xxvia_pk_fact_pos_ar_VIANNEY.ejecuta_conc_fact_VIANNEY"
                                               .CommandType = adCmdStoredProc
                                               
                                               rsaux10.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                                               If Not rsaux10.EOF Then
                                                  var_responsabilidad_facturacion = IIf(IsNull(rsaux10!RESPONSABILIDAD_FACTURACION), "", rsaux10!RESPONSABILIDAD_FACTURACION)
                                               End If
                                               rsaux10.Close
                                    
                                               Set objParm = .CreateParameter("p_responsabilidad", adVarChar, adParamInput, 100, var_responsabilidad_facturacion)
                                               .Parameters.Append objParm
                                            
                                               Set objParm = .CreateParameter("p_customer_trx_id", adNumeric, adParamInput, 50, var_customer_trx_id)
                                               .Parameters.Append objParm
                                     
                                               Set objParm = .CreateParameter("p_esperar", adNumeric, adParamInput, 50, 1)
                                               .Parameters.Append objParm
                                   
                                               Set objParm = .CreateParameter("p_fact_pagada", adVarChar, adParamInput, 50, "Y")
                                               .Parameters.Append objParm
                                    
                                               var_estatus_factura = ""
                                               Set objParm = .CreateParameter("p_estatus", adVarChar, adParamOutput, 50, var_estatus_factura)
                                               .Parameters.Append objParm
                                               rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                               rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                               On Error GoTo salir2:
                                               .execute
                                    
                                               var_estatus_factura = .Parameters("p_estatus").Value
                                               If var_estatus_factura = "N" Then
                                                  GoTo salir_factura:
                                               End If
                                               objConn.CommitTrans
                                          End With
                                          Set objConn = Nothing
                                          Set objCmd = Nothing
                                          rsaux2.MoveNext
                                    Wend
                                    rsaux2.Close
                                    Else
                                       MsgBox "Favor de correr el concurrente XXVIA - Facturacion VIANNEY (Eflow) directamente en Oracle", vbOKOnly, "ATENCION"
                                    End If
                                 'End If
                                 If rsaux1.State = 1 Then
                                    rsaux1.Close
                                 End If
                                 Me.txt_codigo.Enabled = False
                                 
                                 var_encontro = 0
                                 var_cadena = "SELECT RCT.CUSTOMER_TRX_ID, APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, APS.STATUS, APS.CLASS, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME From hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND RCT.BILL_TO_SITE_USE_ID = HCSU.SITE_USE_ID AND INTERFACE_HEADER_ATTRIBUTE1 IN ('" + CStr(var_pedido) + "') AND RCT.customer_trx_id = APS.customer_trx_id "
                                 rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux2.EOF Then
                                    var_encontros = 1
                                    var_customer_trx_id = rsaux2!customer_Trx_id
                                    var_factura = rsaux2!trx_number
                                    rsaux8.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux8.EOF Then
                                       var_serie_pedido = IIf(IsNull(rsaux8!Serie), "", rsaux8!Serie)
                                       var_ruta_facturas = IIf(IsNull(rsaux8!ruta_facturas), "", rsaux8!ruta_facturas)
                                    End If
                                    rsaux8.Close
                                    
                                    
                                    rsaux1.Open "select customer_trx_id from xxvia_Tb_control_doc_fiscales where customer_trx_id = " + CStr(var_customer_trx_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    var_posible = 0
                                    If rsaux1.EOF Then
                                       var_posible = 1
                                    Else
                                       rsaux10.Open "CALL XXVIA_SEND_POST('" + CStr(rsaux1!customer_Trx_id) + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    End If
                                    rsaux1.Close
                                    If var_posible = 0 Then
                                       URL = "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=VTH981105F90&serie=" + Trim(var_serie_pedido) + "&folio=" + Trim(CStr(var_factura))
                                       buf = Split(URL, ".")
                                       ext = buf(UBound(buf))
                                       strSavePath = "C:\SISTEMAS\" + Trim(var_serie_pedido) + Trim(CStr(var_factura)) + ".pdf"
                                       ret = URLDownloadToFile(0, URL, strSavePath, 0, 0)
                                       If ret = 0 Then
                                          Call ShellExecute(Me.hwnd, "print", "C:\SISTEMAS\" + Trim(var_serie_pedido) + Trim(CStr(var_factura)) + ".PDF", vbNullString, vbNullString, SW_SHOWNORMAL = 1)
                                       Else
                                          MsgBox "Error en la factura " + Trim(var_serie_pedido) + Trim(CStr(var_factura))
                                       End If
                                    Else
                                       MsgBox "No se a generado la factura", vbOKOnly, "ATENCION"
                                    End If
                                    
                                    
                                    
                                    
                                    'Open (App.Path & "\EJPDF" + Trim(var_serie_pedido) + Trim(CStr(var_factura)) + ".bat") For Output As #2
                                    'Print #2, "START " + var_ruta_facturas + var_serie_pedido + "\"; Trim(var_serie_pedido) + Trim(Str(var_factura)) + ".PDF"
                                    'Close #2
                                    'var_Archivo = App.Path & "\EJPDF" + Trim(var_serie_pedido) + Trim(CStr(var_factura)) + ".bat"
                                    'x = Shell(var_Archivo, vbHide)
                                    
                                 Else
                                    MsgBox "No se a generado la factura", vbOKOnly, "ATENCION"
                                 End If
                                 rsaux2.Close
                                 
                                 
                                 
                                 
                                 
                              End If
                              rsaux.Close
                              End If
                           Else
                              MsgBox "No se puede dar salida debido a: " + var_cadena_posible_existencias, vbOKOnly, "ATENCION"
                           End If
                        End If
                     End If
                     If rs.State = 1 Then
                        rs.Close
                     End If
                  Else
                     MsgBox "No se a indicado una lista de precios", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "No se a indicado un tipo de pedido", vbOKOnly, "ATENCION"
               End If
            End If
         Else
            If var_clave_movimiento = "SP" Or var_clave_movimiento = "EP" Then
               If var_estatus_movimiento <> "I" Then
                  rsaux.Open "UPDATE XXVIA_TB_DEVOLUCIONES_CLIENTES A SET ESTATUS = 'I', fecha_fin = SYSDATE WHERE A.NUMERO = " + Me.txt_folio + " AND A.ORGANIZACION = " + var_unidad_organizacional + "  AND A.MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  cnn.BeginTrans
                  rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_SALIDAS_PRIVALIA", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
                  End If
                  rsaux.Close
                  rsaux.Open "INSERT INTO TB_TEMP_ORACLE_SALIDAS_PRIVALIA (INTE_tEM_CONSECUTIVO)  VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  cnn.CommitTrans
                  rsaux.Open "ALTER SESSION SET NLS_DATE_FORMAT = 'DD/MM/YYYY HH24:MI:SS'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  rsaux.Open "select fecha_inicio, fecha_fin, CODIGO, DESCRIPTION, CANTIDAD from xxvia_Tb_devoluciones_Clientes A, XXVIA_SYSTEM_ITEMS_B B where movimiento = '" + var_clave_movimiento + "' and numero = " + Me.txt_folio + " AND A.ORGANIZACION = B.ORGANIZATION_ID AND CODIGO = SEGMENT1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rsaux.EOF
                        rsaux1.Open "INSERT INTO TB_TEMP_ORACLE_SALIDAS_PRIVALIA (INTE_TEM_CONSECUTIVO, FECHA_INICIO, FECHA_FIN, CODIGO, DESCRIPCION, CANTIDAD, FOLIO, MOVIMIENTO) VALUES (" + CStr(var_consecutivo) + ",'" + CStr(IIf(IsNull(rsaux!FECHA_INICIO), "", rsaux!FECHA_INICIO)) + "','" + CStr(IIf(IsNull(rsaux!fecha_fin), "", rsaux!fecha_fin)) + "','" + rsaux!codigo + "','" + rsaux!Description + "'," + CStr(rsaux!cantidad) + "," + Me.txt_folio + ",'" + var_clave_movimiento + "')", cnn, adOpenDynamic, adLockOptimistic
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  Set reporte = appl.OpenReport(App.Path + "\rep_oracle_salidas_privalia.rpt")
                  reporte.RecordSelectionFormula = "{VW_ORACLE_SALIDAS_PRIVALIA.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Privalia"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  rsaux.Open "delete from TB_TEMP_ORACLE_SALIDAS_PRIVALIA where inte_tem_consecutivo = " + Str(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  Me.txt_codigo.Enabled = False
               Else
                  cnn.BeginTrans
                  rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_SALIDAS_PRIVALIA", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
                  End If
                  rsaux.Close
                  rsaux.Open "INSERT INTO TB_TEMP_ORACLE_SALIDAS_PRIVALIA (INTE_tEM_CONSECUTIVO)  VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  cnn.CommitTrans
                  rsaux.Open "ALTER SESSION SET NLS_DATE_FORMAT = 'DD/MM/YYYY HH24:MI:SS'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  rsaux.Open "select fecha_inicio, fecha_fin, CODIGO, DESCRIPTION, CANTIDAD from xxvia_Tb_devoluciones_Clientes A, XXVIA_SYSTEM_ITEMS_B B where movimiento = '" + var_clave_movimiento + "' and numero = " + Me.txt_folio + " AND A.ORGANIZACION = B.ORGANIZATION_ID AND CODIGO = SEGMENT1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rsaux.EOF
                        rsaux1.Open "INSERT INTO TB_TEMP_ORACLE_SALIDAS_PRIVALIA (INTE_TEM_CONSECUTIVO, FECHA_INICIO, FECHA_FIN, CODIGO, DESCRIPCION, CANTIDAD, FOLIO, MOVIMIENTO) VALUES (" + CStr(var_consecutivo) + ",'" + CStr(IIf(IsNull(rsaux!FECHA_INICIO), "", rsaux!FECHA_INICIO)) + "','" + CStr(IIf(IsNull(rsaux!fecha_fin), "", rsaux!fecha_fin)) + "','" + rsaux!codigo + "','" + rsaux!Description + "'," + CStr(rsaux!cantidad) + "," + Me.txt_folio + ",'" + var_clave_movimiento + "')", cnn, adOpenDynamic, adLockOptimistic
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  Set reporte = appl.OpenReport(App.Path + "\rep_oracle_salidas_privalia.rpt")
                  reporte.RecordSelectionFormula = "{VW_ORACLE_SALIDAS_PRIVALIA.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Ordenes de surtido historica"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  rsaux.Open "delete from TB_TEMP_ORACLE_SALIDAS_PRIVALIA where inte_tem_consecutivo = " + Str(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               End If
            End If
         End If
         ''''
      End If
   Else
      MsgBox "No se a seleccionado ning�n movimiento", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir2:
   'MsgBox Err.Number
   
   If Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      Resume
   Else
      '
      'MsgBox Err.Description
      Resume
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
   End If
   Exit Sub
salir_factura:
   MsgBox "No se pudo generar el documento electr�nico", vbOKOnly, "ATENCION"
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
      Exit Sub
no_job:
    MsgBox "Vueleva a ejecutar la impresi�n", vbOKOnly, "ATENCION"
                             strconsulta = "delete from xxvia_tb_job_batch_log where vcha_transaction_reference = ?"
                             With comandoORA
                                  .ActiveConnection = cnnoracle_4
                                  .CommandType = adCmdText
                                  .CommandText = strconsulta
                                  Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(Me.txt_folio))
                                  .Parameters.Append parametro
                             End With
                             Set rsaux12 = comandoORA.execute
                             Set comandoORA = Nothing
                             Set parametro = Nothing
      
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
      If rsaux11.State = 1 Then
         rsaux11.Close
      End If
      If rsaux12.State = 1 Then
         rsaux12.Close
      End If
      Exit Sub
    
End Sub


Private Sub cmd_movimiento_masivo_Click()
   Dim var_inserta As Boolean
   Dim var_factura As Integer
   Dim var_posible_cliente As Boolean
    
   Dim codigo As String
   Dim cantidad As Double

   On Error GoTo e
   Open "c:\sistemas\texto.txt" For Input As #1
   Do While Not EOF(1)
      Line Input #1, sCadena1
      palabras = Split(sCadena1, " ")
          
          Me.txt_codigo = Trim(palabras(0))
          var_cantidad_leida = CDbl(Trim(palabras(1)))
          If Trim(Me.txt_codigo) <> "" Then
             rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
             If Not rsaux8.EOF Then
                var_unidad_medida = rsaux8!PRIMARY_UOM_CODE
                var_descripcion_articulo = rsaux8!Description
                var_inventory_item_id = rsaux8!inventory_item_id
                var_salida_masiva = IIf(IsNull(rsaux8!attribute10), "N", rsaux8!attribute10)
                var_cadena = "select * from  qp_secu_list_headers_v a, qp_list_lines_v b Where a.list_header_id = b.list_header_id and  B.product_attr_value = " + CStr(var_inventory_item_id) + " AND a.list_header_id = " + CStr(var_clave_lista_precios) + " and   product_attr_val_disp = '" + Me.txt_codigo + "'"
                rsaux10.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                If Not rsaux10.EOF Then
                   If IIf(IsNull(rsaux10!OPERAND), 0, rsaux10!OPERAND) > 0 Then
                      var_cantidad_leida = var_cantidad_leida
                      Me.txt_foco.Enabled = True
                      Me.txt_foco.SetFocus
'----------------
                      If Trim(txt_codigo.Text) <> "" Then
                         If var_primera_vez = True Then
                            rs.Open "select * from xxvia_tb_folios_dev_clientes WHERE MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                            If Not rs.EOF Then
                               var_numero_folio = rs(0).Value + 1
                               Me.txt_folio = rs(0).Value + 1
                               rsaux.Open "update xxvia_tb_folios_dev_clientes set folio =  folio + 1 WHERE MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                            Else
                               rsaux.Open "insert into xxvia_tb_folios_dev_clientes (folio, MOVIMIENTO) values (1,'" + var_clave_movimiento + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                               var_numero_folio = 1
                               Me.txt_folio = 1
                            End If
                            rs.Close
                            var_primera_vez = False
                         End If
                         Cadena = "select * from xxvia_tb_devoluciones_clientes where numero = " + Str(var_numero_folio) + " and codigo = '" + txt_codigo + "' and inventory_item_id = " + CStr(var_inventory_item_id) + " and localizador = '" + var_localizador_subinventario + "' AND MOVIMIENTO = '" + var_clave_movimiento + "'"
                         rs.Open Cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                         If rs.EOF Then
                            lbl_total = CStr(CDbl(lbl_total) + var_cantidad_leida)
                            var_inserta = False
                            rs.Close
                            If Me.txt_establecimiento = "" Then
                               Me.txt_establecimiento = 0
            
                            End If
                            var_cadena = "insert into xxvia_tb_devoluciones_clientes (numero, organizacion, inventory_item_id, codigo, cantidad, descripcion, estatus, agente, cliente, establecimiento, titular, nombre_agente, almacen, nombre_almacen, nombre_cliente, nombre_establecimiento, referencia, usuario, maquina, fecha_inicio, unidad_medida, precio, localizador, movimiento,tipo_pedido)"
                            var_cadena = var_cadena + " values (" + CStr(var_numero_folio) + "," + var_unidad_organizacional + "," + CStr(var_inventory_item_id) + ",'" + Me.txt_codigo + "'," + CStr(var_cantidad_leida) + ",'" + var_descripcion_articulo + "',''," + Me.txt_agente + "," + Me.txt_cliente + "," + Me.txt_establecimiento + "," + CStr(var_clave_titular) + ",'" + Me.txt_nombre_agente + "','" + Me.txt_almacen + "','" + Me.txt_nombre_almacen + "','" + Me.txt_nombre_cliente + "','" + Me.txt_nombre_establecimiento + "','" + Me.txt_referencia + "','" + var_clave_usuario_global + "', '" + fun_NombrePc + "','" + CStr(Date) + "','" + var_unidad_medida + "',0,'" + var_localizador_subinventario + "','" + var_clave_movimiento + "'," + CStr(var_tipo_pedido) + ")"
                            'MsgBox var_cadena
                            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                            valor = Trim(txt_codigo)
         
                            Set list_item = lv_entradas.ListItems.Add(, , Trim(txt_codigo))
                            list_item.SubItems(1) = var_descripcion_articulo
                            list_item.SubItems(2) = var_cantidad_leida
                            list_item.SubItems(3) = var_localizador_subinventario
                            var_renglon = lv_entradas.ListItems.Count
                            'Call ilumina_grid
                            txt_codigo = ""
                         Else
                            var_inserta = False
                            lbl_total = CStr(CDbl(lbl_total) + var_cantidad_leida)
                            rs.Close
                            rs.Open "update xxvia_tb_devoluciones_clientes set cantidad = cantidad +" + CStr(var_cantidad_leida) + " where numero = " + CStr(var_numero_folio) + " and inventory_item_id = " + CStr(var_inventory_item_id) + " and codigo = '" + Me.txt_codigo + "' and localizador = '" + var_localizador_subinventario + "' and movimiento = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                            valor = Me.txt_codigo
                            var_j = 1
                            For var_j = 1 To Me.lv_entradas.ListItems.Count
                                Me.lv_entradas.ListItems.Item(var_j).Selected = True
                                If Me.lv_entradas.selectedItem = Me.txt_codigo And Trim(Me.lv_entradas.selectedItem.SubItems(3)) = Trim(var_localizador_subinventario) Then
                                   Me.lv_entradas.selectedItem.SubItems(2) = CDbl(Me.lv_entradas.selectedItem.SubItems(2)) + var_cantidad_leida
                                   var_renglon = var_j
                                End If
                            Next var_j
                            'Call ilumina_grid
                            txt_codigo = ""
                         End If
                         'txt_codigo.SetFocus
                         'txt_codigo = ""
                      End If
                                    
'----------------
                   Else
                      frmmensaje.lbl_mensaje = "El art�culo no tiene precio " + Me.txt_codigo
                      frmmensaje.Show
                      Me.txt_codigo = ""
                   End If
                Else
                   frmmensaje.lbl_mensaje = "El art�culo no se encuentra en la lista de precios del cliente " + Me.txt_codigo
                   frmmensaje.Show
                   Me.txt_codigo = ""
                End If
                rsaux10.Close
             Else
                frmmensaje.lbl_mensaje = "Error en c�digo " + Me.txt_codigo
                frmmensaje.Show
                txt_codigo = ""
             End If
             rsaux8.Close
          End If
   Loop
   Close #1
   Exit Sub
e:
MsgBox "A surgido un error al subir el archivo"
    
End Sub

Private Sub cmd_nuevo_Click()
   var_devolucion_costales = 0
   lbl_total = "0"
   lbl_cancelado = ""
   If var_numero_folio > 0 Then
      If rs.State = 1 Then
         rs.Close
      End If
     rs.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
   End If
   txt_codigo.Enabled = False
   var_primera_vez = True
   var_ventana = 0
   frm_busqueda.Visible = False
   lv_entradas.ListItems.Clear
   var_numero_folio = 0
   txt_folio = ""
   txt_codigo = ""
   var_estatus_movimiento = ""
   txt_cliente = ""
   txt_establecimiento = ""
   txt_agente = ""
   txt_almacen = ""
   txt_referencia = ""
   Me.txt_titular = ""
   Me.txt_nombre_titular = ""
   Me.txt_clave_titular = ""
   Me.txt_titular.Enabled = True
   Me.txt_nombre_titular.Enabled = True
   txt_almacen.Enabled = True
   txt_cliente.Enabled = False
   txt_agente.Enabled = False
   txt_establecimiento.Enabled = False
   txt_referencia.Enabled = False
   txt_codigo.Enabled = False
   txt_cliente.Enabled = False
   If var_clave_movimiento <> "SP" Then
      If var_clave_movimiento <> "EP" Then
         If var_clave_usuario_global <> "U0000000763" And var_clave_usuario_global <> "U0000001250" Then
            txt_almacen.SetFocus
         End If
      End If
   End If
   txt_nombre_almacen = ""
   txt_nombre_agente = ""
   txt_nombre_establecimiento = ""
   txt_nombre_cliente = ""
   If var_clave_movimiento = "SP" Or var_clave_movimiento = "EP" Then
      Me.txt_almacen = "PRIVALIA"
      Me.txt_nombre_almacen = "PT. ALMACEN PRIVALIA"
      Me.txt_agente = 1026
      Me.txt_nombre_agente = "JUAN MANUEL ROMERO FIGUEROA"
      Me.txt_titular = "T000009781"
      Me.txt_clave_titular = 52987
      var_clave_titular = 52987
      Me.txt_nombre_titular = "EDGAR IBARRA"
      Me.txt_cliente = 160613
      Me.txt_establecimiento = 160615
      Me.txt_nombre_cliente = "PRIVALIA VENTA DIRECTA,S.A DE C.V"
      Me.txt_nombre_establecimiento = "PRIVALIA VENTA DIRECTA,S.A DE C.V"
      var_cadena = "SELECT  hcsu.price_list_id, hcsu.order_type_id,hca.cust_account_id, hcp.site_use_id AS VCHA_CLI_CLAVE_ID,hl.address1 VCHA_CLI_NOMBRE FROM hz_parties hp,hz_party_sites hps,hz_cust_accounts hca,hz_cust_acct_sites_all hcas,hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcp.site_use_id = " + Me.txt_cliente
      rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_clave_lista_precios = IIf(IsNull(rs!price_list_id), 0, rs!price_list_id)
      End If
      rs.Close
      Me.txt_almacen.Enabled = False
      Me.txt_nombre_almacen.Enabled = False
      Me.txt_agente.Enabled = False
      Me.txt_nombre_agente.Enabled = False
      Me.txt_titular.Enabled = False
      Me.txt_clave_titular.Enabled = False
      Me.txt_nombre_titular.Enabled = False
      Me.txt_cliente.Enabled = False
      Me.txt_establecimiento.Enabled = False
      Me.txt_nombre_cliente.Enabled = False
      Me.txt_nombre_establecimiento.Enabled = False
      Me.txt_referencia.Enabled = True
      Me.txt_referencia.SetFocus
   End If
   
   If var_clave_movimiento = "SML" Or var_clave_movimiento = "EML" Then
      Me.txt_almacen = "MERCADOLIB"
      Me.txt_nombre_almacen = "PT. MERCADO LIBRE"
      Me.txt_agente = 85001
      Me.txt_nombre_agente = "0247 VHD VENTA EN LINEA 24/7"
      Me.txt_titular = "CN_VIA"
      Me.txt_clave_titular = 272259
      var_clave_titular = 272259
      Me.txt_nombre_titular = "VIANNEY TEXTIL HOGAR"
      Me.txt_cliente = 1030563
      Me.txt_establecimiento = 1121453
      Me.txt_nombre_cliente = "VENTAS AL PUBLICO EN GENERAL"
      Me.txt_nombre_establecimiento = "VENTAS AL PUBLICO EN GENERAL"
      var_cadena = "SELECT  hcsu.price_list_id, hcsu.order_type_id,hca.cust_account_id, hcp.site_use_id AS VCHA_CLI_CLAVE_ID,hl.address1 VCHA_CLI_NOMBRE FROM hz_parties hp,hz_party_sites hps,hz_cust_accounts hca,hz_cust_acct_sites_all hcas,hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcp.site_use_id = " + Me.txt_cliente
      rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_clave_lista_precios = IIf(IsNull(rs!price_list_id), 0, rs!price_list_id)
      End If
      rs.Close
      Me.txt_almacen.Enabled = False
      Me.txt_nombre_almacen.Enabled = False
      Me.txt_agente.Enabled = False
      Me.txt_nombre_agente.Enabled = False
      Me.txt_titular.Enabled = False
      Me.txt_clave_titular.Enabled = False
      Me.txt_nombre_titular.Enabled = False
      Me.txt_cliente.Enabled = False
      Me.txt_establecimiento.Enabled = False
      Me.txt_nombre_cliente.Enabled = False
      Me.txt_nombre_establecimiento.Enabled = False
      Me.txt_referencia.Enabled = True
      Me.txt_referencia.SetFocus
   End If
   
   
   
   
   
   
   
   If var_clave_usuario_global = "U0000000763" Then
      Me.txt_almacen = "PRIVALIA"
      Me.txt_nombre_almacen = "PT. ALMACEN PRIVALIA"
      Me.txt_agente = 95003
      Me.txt_nombre_agente = "INTERNET"
      Me.txt_titular = "T000009781"
      Me.txt_clave_titular = 52987
      var_clave_titular = 52987
      Me.txt_nombre_titular = "EDGAR IBARRA"
      Me.txt_cliente = 160613
      Me.txt_establecimiento = 160615
      Me.txt_nombre_cliente = "PRIVALIA VENTA DIRECTA,S.A DE C.V"
      Me.txt_nombre_establecimiento = "PRIVALIA VENTA DIRECTA,S.A DE C.V"
      var_cadena = "SELECT  hcsu.price_list_id, hcsu.order_type_id,hca.cust_account_id, hcp.site_use_id AS VCHA_CLI_CLAVE_ID,hl.address1 VCHA_CLI_NOMBRE FROM hz_parties hp,hz_party_sites hps,hz_cust_accounts hca,hz_cust_acct_sites_all hcas,hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcp.site_use_id = " + Me.txt_cliente
      rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_clave_lista_precios = IIf(IsNull(rs!price_list_id), 0, rs!price_list_id)
      End If
      rs.Close
      Me.txt_almacen.Enabled = False
      Me.txt_nombre_almacen.Enabled = False
      Me.txt_agente.Enabled = False
      Me.txt_nombre_agente.Enabled = False
      Me.txt_titular.Enabled = False
      Me.txt_clave_titular.Enabled = False
      Me.txt_nombre_titular.Enabled = False
      Me.txt_cliente.Enabled = False
      Me.txt_establecimiento.Enabled = False
      Me.txt_nombre_cliente.Enabled = False
      Me.txt_nombre_establecimiento.Enabled = False
      Me.txt_referencia.Enabled = True
      Me.txt_referencia.SetFocus
   End If
End Sub

Private Sub cmd_pasar_todo_Click()
   Me.frm_pasar_todo.Visible = True
   Me.txt_numero.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub



Private Sub Command1_Click()
On Error GoTo SALIR:
rsaux.Open "CALL XXVIA_PK_INTERFACES_OM.importar_pedido('TIE_01'," + var_empresa + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
rsaux.Open "select * from oe_order_headers_all where orig_sys_document_ref = 'TIE_01'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_nUMERO_pedido = rsaux!order_number
rsaux.Close

Exit Sub
SALIR:
   If Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
       rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      Resume
   End If


End Sub

Private Sub Command2_Click()
    rs.Open "select codigo from xxvia_tb_devoluciones_clientes where numero = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
    var_cadena = ""
    While Not rs.EOF
          rsaux1.Open "select CUSTOMER_ORDER_FLAG, customer_order_enabled_flag, SHIPPABLE_ITEM_FLAG, INTERNAL_ORDER_FLAG, internal_order_enabled_flag, SO_TRANSACTIONS_FLAG, RETURNABLE_FLAG, INVOICEABLE_ITEM_FLAG from xxvia_system_items_b where segment1='" + rs!codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
          If IIf(IsNull(rsaux1!RETURNABLE_FLAG), "", rsaux1!RETURNABLE_FLAG) <> "Y" Then
             If var_cadena = "" Then
                var_cadena = var_cadena + " " + rs!codigo
             Else
                var_cadena = var_cadena + ", " + rs!codigo
             End If
          End If
          rsaux1.Close
          rs.MoveNext
    Wend
    rs.Close
    Me.txt_clave_titular = var_cadena
    MsgBox var_cadena
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
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If var_ventana = 0 Then
         Unload Me
      Else
         Me.frm_busqueda.Visible = False
         Me.frm_eliminar.Visible = False
         Me.frm_lista.Visible = False
         var_ventana = 0
      End If
   End If
End Sub

Private Sub Form_Load()
   If var_unidad_organizacional = 90 Then
      Me.txt_cantidad_eliminar.MaxLength = 90
   Else
      Me.txt_cantidad_eliminar.MaxLength = 10
   End If
   Me.chk_factura = 0
   Me.chk_factura.Visible = False
   Me.txt_movimiento = var_clave_movimiento
   Me.txt_movimiento.Visible = False
   Me.frm_pasar_todo.Visible = False
   If var_clave_usuario_global = "11" Or var_clave_usuario_global = "8" Then
      Me.cmd_pasar_todo.Visible = True
   Else
      If var_unidad_organizacional = "93" Then
         Me.cmd_pasar_todo.Visible = True
      End If
   End If
   lbl_total = "0"
   lbl_cancelado = ""
   var_a�o = 2005
   var_numero_folio = 0
   var_cadena_seguridad = ""
   Top = 0
   Left = 1500
   frm_lista.Visible = False
   var_estatus_movimiento = ""
   var_ventana = 0
   frm_busqueda.Visible = False
   frm_eliminar.Visible = False
   lbl_cantidad.Visible = False
   txt_cantidad.Visible = False
   txt_cliente.Enabled = False
   txt_codigo.Enabled = False
   txt_agente.Enabled = False
   txt_establecimiento.Enabled = False
   var_primera_vez = True
   var_cantidad_leida = 1#
   var_ventana = 0
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   'Me.cmd_pasar_todo.Visible = False
   If var_clave_movimiento = "SNC" Then
      Me.cmd_devolucion.Visible = True
   Else
      Me.cmd_devolucion.Value = False
   End If
   If var_clave_usuario_global = "U0000000763" Or var_clave_usuario_global = "U0000001250" Then
      Me.cmd_movimiento_masivo.Visible = True
   End If
   If var_clave_movimiento <> "DC" Then
      Me.cmd_asignar_causa_devolucion.Visible = False
   End If
    Me.cmd_movimiento_masivo.Visible = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If var_solo_lectura = False Then
   End If
   Call activa_forma(var_activa_forma_entradas_devoluciones)
   If var_numero_folio > 0 Then
     If rs.State = 1 Then
        rs.Close
     End If
     rs.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
   End If
End Sub

Private Sub lv_entradas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         MsgBox "Imposible modificar el movimiento", vbOKOnly, "ATENCION"
      Else
         If var_unidad_organizacional = 90 Then
            Me.lbl_nombre_eliminar = "C�digo de barras a eliminar"
         Else
         End If
         var_elimina = False
         var_ventana = 1
         frm_eliminar.Visible = True
         txt_cantidad_eliminar.SetFocus
      End If
   End If
End Sub


Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_GotFocus()
   var_ventana = 1
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 0 Then
         If VAR_TIPO_LISTA = 1 Then
            txt_almacen = lv_lista.selectedItem
            txt_nombre_almacen = lv_lista.selectedItem.SubItems(1)
            txt_almacen.Enabled = True
            txt_almacen.SetFocus
         End If
         If VAR_TIPO_LISTA = 2 Then
            txt_agente = lv_lista.selectedItem
            txt_nombre_agente = lv_lista.selectedItem.SubItems(1)
            txt_agente.Enabled = True
            txt_agente.SetFocus
         End If
         If VAR_TIPO_LISTA = 3 Then
            txt_establecimiento = lv_lista.selectedItem
            txt_nombre_establecimiento = lv_lista.selectedItem.SubItems(1)
            txt_establecimiento.Enabled = True
            txt_establecimiento.SetFocus
         End If
         If VAR_TIPO_LISTA = 4 Then
            txt_cliente = lv_lista.selectedItem
            txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
            txt_cliente.Enabled = True
            txt_cliente.SetFocus
         End If
         If VAR_TIPO_LISTA = 5 Then
            Me.txt_titular = lv_lista.selectedItem
            Me.txt_nombre_titular = lv_lista.selectedItem.SubItems(1)
            Me.txt_clave_titular = lv_lista.selectedItem.SubItems(2)
            Me.txt_titular.Enabled = True
            Me.txt_titular.SetFocus
         End If
      End If
      frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
      
   End If
End Sub

Private Sub lv_lista_LostFocus()
   var_ventana = 0
   frm_lista.Visible = False
End Sub

Private Sub txt_agente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci�n disponible"
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_unidad_organizacional = "85" Or var_unidad_organizacional = "94" Then
         If var_clave_usuario_global = "U0000000314" Or var_clave_usuario_global = "U0000000390" Then
            rs.Open "SELECT  distinct arc.collector_id as vcha_age_agente_id, arc.name as vcha_age_nombre  FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcas.org_id = " + var_empresa + " AND arc.collector_id = 1016 order by arc.name", cnnoracle_4, adOpenDynamic, adLockOptimistic
         Else
            rs.Open "SELECT  distinct arc.collector_id as vcha_age_agente_id, arc.name as vcha_age_nombre  FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcas.org_id = " + var_empresa + " AND arc.collector_id = 1028 order by arc.name", cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
      Else
         'rs.Open "SELECT  distinct arc.collector_id as vcha_age_agente_id, arc.name as vcha_age_nombre  FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcas.org_id = " + var_empresa + " order by arc.name", cnnoracle_4, adOpenDynamic, adLockOptimistic
          rs.Open "select distinct collector_id as vcha_age_agente_id, name as vcha_age_nombre from ar_collectors", cnnoracle_4, adOpenDynamic, adLockOptimistic
      End If
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Agentes"
      VAR_TIPO_LISTA = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Me.txt_titular.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txt_agente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_agente) <> "" Then
      If IsNumeric(Me.txt_agente) Then
         rs.Open "SELECT  distinct arc.collector_id as vcha_age_agente_id, arc.name as vcha_age_nombre  FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcas.org_id = " + var_empresa + " and arc.collector_id= " + Me.txt_agente, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_nombre_agente = rs!VCHA_AGE_NOMBRE
            rs.Close
            txt_agente.Enabled = False
            txt_titular.Enabled = True
            txt_titular.SetFocus
         Else
            rs.Close
            MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
            txt_agente = ""
            txt_nombre_agente = ""
         End If
      Else
         MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
         txt_agente = ""
         txt_nombre_agente = ""
      End If
   End If
End Sub

Private Sub txt_almacen_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci�n disponible"
End Sub

Private Sub txt_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_unidad_organizacional = "93" Then
         If var_clave_usuario_global = "U0000000315" Then
            'rs.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
            rs.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " AND secondary_inventory_name in ('RECEPMEX')", cnnoracle_4, adOpenDynamic, adLockOptimistic
            'rs.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " AND secondary_inventory_name in ('CDI_ALMCAL','CDI_ALMPT')", cnnoracle_4, adOpenDynamic, adLockOptimistic
         Else
            rs.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " AND secondary_inventory_name in ('CDI_ALMCAL','CDI_ALMPT')", cnnoracle_4, adOpenDynamic, adLockOptimistic
            'rs.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!vcha_alm_almacen_id)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
               rs.MoveNext
         Wend
         rs.Close
      End If
      If var_unidad_organizacional = "90" Then
         If var_clave_usuario_global = "U0000000241" Then
            rs.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " AND secondary_inventory_name in ('TEX_MP_INV')", cnnoracle_4, adOpenDynamic, adLockOptimistic
         Else
            If var_clave_usuario_global = "U0000000430" Then
               rs.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " AND secondary_inventory_name in ('TEX_PT_CON')", cnnoracle_4, adOpenDynamic, adLockOptimistic
            Else
               If var_clave_usuario_global = "U0000000164" Or var_clave_usuario_global = "U0000000046" Then
                  rs.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " AND secondary_inventory_name in ('TEXCALIDAD')", cnnoracle_4, adOpenDynamic, adLockOptimistic
               Else
                  rs.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " AND secondary_inventory_name in ('TEX_PT_QL','TEXCALIDAD','TEX_ALM_PT','TEX_VB1','TEX_VB2','TEX_VB4','TEX_ALM_VR','TEX_VB8')", cnnoracle_4, adOpenDynamic, adLockOptimistic
               End If
            End If
         End If
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!vcha_alm_almacen_id)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
               rs.MoveNext
         Wend
         rs.Close
      End If
      
      
      If var_unidad_organizacional = "86" Then
               rs.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!vcha_alm_almacen_id)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
               rs.MoveNext
         Wend
         rs.Close
      End If
      
      
      If var_unidad_organizacional = "89" Then
         rs.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " and secondary_inventory_name = 'EMBPT'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!vcha_alm_almacen_id)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
               rs.MoveNext
         Wend
         rs.Close
      End If
      If var_unidad_organizacional = "85" Then
         If var_clave_usuario_global = "U0000000314" Then
            rs.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " AND secondary_inventory_name in ('PRODTER')", cnnoracle_4, adOpenDynamic, adLockOptimistic
         Else
            If var_clave_usuario_global = "U0000000430" Then
               rs.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + "  AND secondary_inventory_name in ('ALMMP')", cnnoracle_4, adOpenDynamic, adLockOptimistic
            Else
               rs.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + "  AND secondary_inventory_name in ('ALMSEGYOBS')", cnnoracle_4, adOpenDynamic, adLockOptimistic
            End If
         End If
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!vcha_alm_almacen_id)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
               rs.MoveNext
         Wend
         rs.Close
      End If
      
      If var_unidad_organizacional = "94" Then
         rs.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + "  AND secondary_inventory_name in ('PRODTER')", cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!vcha_alm_almacen_id)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
               rs.MoveNext
         Wend
         rs.Close
      End If
      
      
      lbl_lista = "Almacenes"
      VAR_TIPO_LISTA = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_almacen_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_agente.Enabled = True
      txt_agente.SetFocus
   End If
End Sub

Private Sub txt_almacen_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_almacen) <> "" Then
      If var_tipo_permiso = 1 Then
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre, nvl(locator_type,0) as localizador from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " and secondary_inventory_name = '" + Me.txt_almacen + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_localizador = rsaux!localizador
            txt_nombre_almacen = rsaux!vcha_alm_nombre
            var_almacen_Destino = txt_almacen
            txt_almacen.Enabled = False
            txt_nombre_almacen.Enabled = False
            txt_agente.Enabled = True
            
            If Me.txt_almacen = "TEX_VB4" Then
               strconsulta = "SELECT  distinct arc.collector_id as vcha_age_agente_id, arc.name as vcha_age_nombre  FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcas.org_id = ? and arc.collector_id= ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, var_empresa)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, 46000)
                    .Parameters.Append parametro
               End With
               Set rsaux = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               Me.txt_agente = rsaux!VCHA_AGE_AGENTE_ID
               Me.txt_nombre_agente = rsaux!VCHA_AGE_NOMBRE
               rsaux.Close
            End If
            
             If Me.txt_almacen = "TEX_VB1" Then
               strconsulta = "SELECT  distinct arc.collector_id as vcha_age_agente_id, arc.name as vcha_age_nombre  FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcas.org_id = ? and arc.collector_id= ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, var_empresa)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, 1004)
                    .Parameters.Append parametro
               End With
               Set rsaux = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               Me.txt_agente = rsaux!VCHA_AGE_AGENTE_ID
               Me.txt_nombre_agente = rsaux!VCHA_AGE_NOMBRE
               rsaux.Close
            End If
             If Me.txt_almacen = "TEX_VB2" Then
               strconsulta = "SELECT  distinct arc.collector_id as vcha_age_agente_id, arc.name as vcha_age_nombre  FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcas.org_id = ? and arc.collector_id= ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, var_empresa)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, 47000)
                    .Parameters.Append parametro
               End With
               Set rsaux = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               Me.txt_agente = rsaux!VCHA_AGE_AGENTE_ID
               Me.txt_nombre_agente = rsaux!VCHA_AGE_NOMBRE
               rsaux.Close
            End If
            
         Else
            MsgBox "Clave de almacen incorrecto", vbOKOnly, "ATENCION"
            txt_almacen = ""
            txt_nombre_almacen = ""
         End If
         If rsaux.State = 1 Then
            rsaux.Close
         End If
      Else
         rs.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre, nvl(locator_type,0) as localizador  from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " and secondary_inventory_name = '" + Me.txt_almacen + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_localizador = rs!localizador
            txt_nombre_almacen = rs!vcha_alm_nombre
            var_almacen_Destino = txt_almacen
            txt_almacen.Enabled = False
            txt_nombre_almacen.Enabled = False
            txt_agente.Enabled = True
         Else
            MsgBox "Clave de almacen incorrecto", vbOKOnly, "ATENCION"
            txt_almacen = ""
         End If
         rs.Close
      End If
   End If
End Sub

Private Sub txt_busqueda_folio_GotFocus()
   var_ventana = 1
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
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
      If IsNumeric(Me.txt_busqueda_folio) Then
         rs.Open "select * from xxvia_tb_devoluciones_clientes where numero = " + Me.txt_busqueda_folio + " and organizacion = " + var_unidad_organizacional + " AND MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_clave_titular = rs!TITULAR
            var_tipo_pedido = IIf(IsNull(rs!tipo_pedido), 0, rs!tipo_pedido)
            var_estatus_movimiento = IIf(IsNull(rs!estatus), "", rs!estatus)
            var_almacen_Destino = rs!ALMACEN
            rsaux8.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre, nvl(locator_type,0) as localizador  from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " and secondary_inventory_name = '" + rs!ALMACEN + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
               var_localizador = rsaux8!localizador
            End If
            rsaux8.Close
            txt_almacen = var_almacen_Destino
            txt_referencia = rs!Referencia
            txt_cliente = rs!Cliente
            txt_nombre_cliente = rs!nombre_cliente
            txt_agente = rs!Agente
            txt_nombre_agente = rs!NOMBRE_AGENTE
            txt_establecimiento = rs!establecimiento
            txt_nombre_establecimiento = rs!nombre_Establecimiento
            
            'var_cadena = "SELECT  hcsu.price_list_id, hcsu.order_type_id,hca.cust_account_id, hcp.site_use_id AS VCHA_CLI_CLAVE_ID,hl.address1 VCHA_CLI_NOMBRE FROM hz_parties hp,hz_party_sites hps,hz_cust_accounts hca,hz_cust_acct_sites_all hcas,hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcp.collector_id = " + Me.txt_agente + " AND hcp.site_use_id = " + Me.txt_cliente
            var_cadena = "SELECT  hcsu.price_list_id, hcsu.order_type_id,hca.cust_account_id, hcp.site_use_id AS VCHA_CLI_CLAVE_ID,hl.address1 VCHA_CLI_NOMBRE FROM hz_parties hp,hz_party_sites hps,hz_cust_accounts hca,hz_cust_acct_sites_all hcas,hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND  hcp.site_use_id = " + Me.txt_cliente
            rsaux10.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux10.EOF Then
               var_clave_lista_precios = rsaux10!price_list_id
            Else
               var_clave_lista_precios = 0
            End If
            rsaux10.Close
            If var_unidad_organizacional = "90" Then
               If Me.txt_cliente <> "8049" Then
                  var_clave_lista_precios = 9016
               End If
            End If
            If var_unidad_organizacional = "89" Or var_unidad_organizacional = "85" Or var_unidad_organizacional = "94" Then
               If Me.txt_agente <> "1028" Then
                  If Me.txt_agente = "1016" Then
                     var_clave_lista_precios = 719011
                  Else
                     var_clave_lista_precios = 9019
                  End If
               End If
               'var_clave_lista_precios = 719011
            End If
            'var_clave_lista_precios = 9007
            
            
            var_cadena = "SELECT  distinct hp.party_name as nombre_titular,  account_number as vcha_tit_titular_id, hcas.cust_account_id AS VCHA_CLI_CLAVE_ID FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors Arc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND hcp.collector_id = " + txt_agente + " and hcas.cust_account_id = " + CStr(rs!TITULAR)
            rsaux8.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
               Me.txt_titular = IIf(IsNull(rsaux8!vcha_tit_titular_id), "", rsaux8!vcha_tit_titular_id)
               Me.txt_nombre_titular = IIf(IsNull(rsaux8!NOMBRE_TITULAR), "", rsaux8!NOMBRE_TITULAR)
               Me.txt_clave_titular = IIf(IsNull(rsaux8!vcha_cli_clave_id), "", rsaux8!vcha_cli_clave_id)
            Else
               Me.txt_titular = ""
               Me.txt_nombre_titular = ""
               Me.txt_clave_titular = ""
            End If
            rsaux8.Close
            Me.txt_titular.Enabled = False
            Me.txt_nombre_titular.Enabled = False
            txt_cliente.Enabled = False
            txt_agente.Enabled = False
            txt_establecimiento.Enabled = False
            txt_cliente.Enabled = False
            txt_almacen.Enabled = False
            txt_referencia.Enabled = False
            lv_entradas.ListItems.Clear
            var_primera_vez = False
            var_numero_folio = rs!numero
            txt_folio = var_numero_folio
            txt_almacen_destino = rs!ALMACEN
            txt_nombre_almacen = rs!ALMACEN
            lbl_total = "0"
            While Not rs.EOF
                  Set list_item = lv_entradas.ListItems.Add(, , rs!codigo)
                  list_item.SubItems(1) = IIf(IsNull(rs!Descripcion), "", rs!Descripcion)
                  list_item.SubItems(2) = IIf(IsNull(rs!cantidad), "", rs!cantidad)
                  list_item.SubItems(3) = IIf(IsNull(rs!localizador), "", rs!localizador)
                  lbl_total = CStr(CDbl(lbl_total) + IIf(IsNull(rs!cantidad), "", rs!cantidad))
                  rs.MoveNext
            Wend
            If lv_entradas.ListItems.Count > 11 Then
               lv_entradas.ColumnHeaders(2).Width = 5050.22
            Else
               lv_entradas.ColumnHeaders(2).Width = 5300.22
            End If
            rs.MoveFirst
            
            If IIf(IsNull(rs!estatus), "", rs!estatus) = "" Then
               Me.txt_codigo.Enabled = True
               Me.txt_foco.Enabled = True
            Else
               Me.txt_codigo.Enabled = False
               Me.txt_foco.Enabled = False
            End If
         Else
            MsgBox "El n�mero de movimiento no existe ", vbOKOnly, "ATENCION"
         End If
         rs.Close
         var_ventana = 0
         frm_busqueda.Visible = False
      Else
         MsgBox ""
      End If
   End If
End Sub

Private Sub txt_busqueda_folio_LostFocus()
   var_ventana = 0
   frm_busqueda.Visible = False
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)

   If var_unidad_organizacional <> 90 Then
      Select Case KeyAscii
      Case 48 To 57, 52, 13, 8, 46, 27
      Case Else
          KeyAscii = 0
      End Select
   End If
   If KeyAscii = 13 Then
      If var_unidad_organizacional = 90 Then
         Dim var_es_caja As Integer
         Dim var_caja As String
         var_posible_caja = 0
         var_es_caja = 0
         var_caja = ""
         If Mid(Me.txt_cantidad_eliminar, 1, 2) = "CA" Or var_tela = "---" Then
            rs.Open "SELECT * FROM XXVIA_TB_CAJAS_PROD WHERE vcha_caj_caja_id = '" + UCase(Me.txt_cantidad_eliminar) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_es_caja = 1
               var_caja = Me.txt_cantidad_eliminar
               If rsaux8.State = 1 Then
                  rsaux8.Close
               End If
               rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + IIf(IsNull(rs!VCHA_ART_ARTICULO_ID), "", rs!VCHA_ART_ARTICULO_ID) + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux8.EOF Then
                  var_estatus_caja = IIf(IsNull(rs!vcha_caj_staus), "", rs!vcha_caj_staus)
                  'var_estatus_caja = "A"
                  
                  If var_estatus_caja <> "PASAR" Then
                     var_posible_caja = 1
                     var_codigo_caja = Me.txt_cantidad_eliminar
                     Me.txt_codigo = IIf(IsNull(rs!VCHA_ART_ARTICULO_ID), "", rs!VCHA_ART_ARTICULO_ID)
                     var_codigo_barras = IIf(IsNull(rs!vcha_Codigo_barras), "", rs!vcha_Codigo_barras)
                     If var_codigo_barras = "" Then
                        var_codigo_barras = Me.txt_codigo + "0119"
                     End If
                     VAR_CANTIDAD_ELIMINAR = rs!numb_caj_cantidad
                  Else
                     var_posible_caja = 0
                  End If
               End If
               rsaux8.Close
            End If
            rs.Close
         Else
            var_codigo_barras = Me.txt_cantidad_eliminar
            var_posible_caja = 1
            VAR_CANTIDAD_ELIMINAR = 1
         End If
         If var_posible_caja = 1 Then
            strconsulta = "select * from xxvia_tb_transacciones where organizacion_id = ? and almacen_id = ? and movimiento_id = ? and numero = ? and codigo_barras = ? AND BULTO = ?"
            With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_almacen)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_clave_movimiento)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.txt_folio))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_codigo_barras)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 50, 0)
                    .Parameters.Append parametro
            End With
            Set rsaux17 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If Not rsaux17.EOF Then
               strconsulta = "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, NVL(a.description,'') AS localizador, B.UNIT_WEIGHT, a.cantidad FROM (select INVENTORY_ITEM_ID, description, cross_reference, nvl(attribute1,1) as cantidad from mtl_cross_references_b) A, (select inventory_item_id, DESCRIPTION, organization_id, segment1, UNIT_WEIGHT from xxvia_system_items_b) B Where a.inventory_item_id = B.inventory_item_id AND B.organization_id = ? AND CROSS_REFERENCE = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_codigo_barras)
                    .Parameters.Append parametro
               End With
               Set rsaux8 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               var_cantidad_leida = 1
               If Not rsaux8.EOF Then
                  VAR_CANTIDAD_ELIMINAR = IIf(IsNull(rsaux8!cantidad), 1, rsaux8!cantidad)
                  Me.txt_codigo = rsaux8!SEGMENT1
               End If
               rsaux8.Close
               If VAR_CANTIDAD_ELIMINAR <= lv_entradas.selectedItem.SubItems(2) * 1 = True Then
                  If Me.txt_codigo = lv_entradas.selectedItem Then
                  strconsulta = "call XXVIA_TB_TRANSACCIONS_CB (?,?,?,?,?,?,?,?,?) "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_almacen)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_clave_movimiento)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, CStr(Me.txt_folio))
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_codigo_barras)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, Me.lv_entradas.selectedItem)
                       .Parameters.Append parametro
                       If var_clave_movimiento = "VDI" Then
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, VAR_CANTIDAD_ELIMINAR)
                          .Parameters.Append parametro
                       Else
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, VAR_CANTIDAD_ELIMINAR * -1)
                          .Parameters.Append parametro
                       End If
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 500, var_nombre_usuario_global)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 0)
                       .Parameters.Append parametro
                  End With
                  Set rsaux16 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  
                  
                     strconsulta = "update xxvia_Tb_cajas_prod set vcha_caj_staus = 'A' where vcha_caj_Caja_id = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_caja)
                          .Parameters.Append parametro
                     End With
                     Set rsaux9 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                  
                  
                  
                  If rsaux8.State = 1 Then
                     rsaux8.Close
                  End If
                  rs.Open "SELECT TO_CHAR(SYSDATE,'dd/mm/yyyy hh24:mi:ss') AS FECHA FROM DUAL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  VAR_FECHA_HORA = rs(0).Value
                  rs.Close
                  rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + Me.lv_entradas.selectedItem + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux8.EOF Then
                     var_inventory_item_id = rsaux8!inventory_item_id
                     var_localizador_subinventario = Me.lv_entradas.selectedItem.SubItems(3)
                     rs.Open "UPDATE xxvia_tb_devoluciones_clientes SET CANTIDAD = CANTIDAD -" + CStr(VAR_CANTIDAD_ELIMINAR) + " where numero = " + Str(var_numero_folio) + " and codigo = '" + Me.lv_entradas.selectedItem + "' and inventory_item_id = " + CStr(var_inventory_item_id) + " and localizador = '" + var_localizador_subinventario + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     On Error GoTo xx:
                     clnt.MSSoapInit var_webservice_texto
                     'var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MOVIMIENTO: " + var_clave_movimiento + Me.txt_folio + "   MAQUINA: " + fun_NombrePc + ", USUARIO: " + var_nombre_usuario_global + Chr(13) + "C�digo: " + Me.lv_entradas.selectedItem + " " + Me.lv_entradas.selectedItem.SubItems(1) + " Cantidad eliminada: " + CStr(VAR_CANTIDAD_ELIMINAR))
                     var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MOVIMIENTO: " + var_clave_movimiento + Me.txt_folio + "  MAQUINA: " + fun_NombrePc + Chr(13) + " USUARIO: " + var_nombre_usuario_global + Chr(13) + "FECHA Y HORA: " + VAR_FECHA_HORA + Chr(13) + "C�digo: " + Me.lv_entradas.selectedItem + " " + Me.lv_entradas.selectedItem.SubItems(1) + " Cantidad eliminada: " + CStr(VAR_CANTIDAD_ELIMINAR))
                  
                     Set clnt = Nothing
xx:
                     lv_entradas.selectedItem.SubItems(2) = lv_entradas.selectedItem.SubItems(2) - Val(VAR_CANTIDAD_ELIMINAR)
                     lbl_total = CStr(CDbl(lbl_total) - Val(VAR_CANTIDAD_ELIMINAR))
                     var_renglon = lv_entradas.selectedItem.Index
                     Call ilumina_grid
                   End If
                   rsaux8.Close
                   Else
                      MsgBox "El c�digo de barras no corresponde al c�digo seleccionado", vbOKOnly, "ATENCION"
                   End If
               Else
                  MsgBox "La cantidad a eliminar supera a la cantidad asignada a la causa de devoluci�n seleccionada", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El c�digo de barras a eliminar no se encuentra en el movimiento", vbOKOnly, "ATENCION"
            End If
            rsaux17.Close
         Else
            MsgBox "La caja no existe", vbOKOnly, "ATENCION"
         End If
         var_ventana = 0
         frm_eliminar.Visible = False
         If Me.txt_codigo.Enabled = True Then
            txt_codigo.SetFocus
         End If
         
      Else
         If IsNumeric(txt_cantidad_eliminar) Then
            VAR_CANTIDAD_ELIMINAR = Val(txt_cantidad_eliminar)
            var_posible_eliminar = True
            If VAR_CANTIDAD_ELIMINAR <= lv_entradas.selectedItem.SubItems(2) * 1 = True Then
               If rsaux8.State = 1 Then
                  rsaux8.Close
               End If
               rs.Open "SELECT TO_CHAR(SYSDATE,'dd/mm/yyyy hh24:mi:ss') AS FECHA FROM DUAL", cnnoracle_4, adOpenDynamic, adLockOptimistic
               VAR_FECHA_HORA = rs(0).Value
               rs.Close
               rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + Me.lv_entradas.selectedItem + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux8.EOF Then
                  var_inventory_item_id = rsaux8!inventory_item_id
                  var_localizador_subinventario = Me.lv_entradas.selectedItem.SubItems(3)
                  rs.Open "UPDATE xxvia_tb_devoluciones_clientes SET CANTIDAD = CANTIDAD -" + CStr(VAR_CANTIDAD_ELIMINAR) + " where numero = " + Str(var_numero_folio) + " and codigo = '" + Me.lv_entradas.selectedItem + "' and inventory_item_id = " + CStr(var_inventory_item_id) + " and localizador = '" + var_localizador_subinventario + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  On Error GoTo x:
                  clnt.MSSoapInit var_webservice_texto
                  'var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MOVIMIENTO: " + var_clave_movimiento + Me.txt_folio + "   MAQUINA: " + fun_NombrePc + ", USUARIO: " + var_nombre_usuario_global + Chr(13) + "C�digo: " + Me.lv_entradas.selectedItem + " " + Me.lv_entradas.selectedItem.SubItems(1) + " Cantidad eliminada: " + CStr(VAR_CANTIDAD_ELIMINAR))
                  var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MOVIMIENTO: " + var_clave_movimiento + Me.txt_folio + "  MAQUINA: " + fun_NombrePc + Chr(13) + " USUARIO: " + var_nombre_usuario_global + Chr(13) + "FECHA Y HORA: " + VAR_FECHA_HORA + Chr(13) + "C�digo: " + Me.lv_entradas.selectedItem + " " + Me.lv_entradas.selectedItem.SubItems(1) + " Cantidad eliminada: " + CStr(VAR_CANTIDAD_ELIMINAR))
               
                  Set clnt = Nothing
x:
                  lv_entradas.selectedItem.SubItems(2) = lv_entradas.selectedItem.SubItems(2) - Val(txt_cantidad_eliminar)
                  lbl_total = CStr(CDbl(lbl_total) - Val(txt_cantidad_eliminar))
                  var_renglon = lv_entradas.selectedItem.Index
                  Call ilumina_grid
                End If
                rsaux8.Close
            Else
               MsgBox "La cantidad a eliminar supera a la cantidad asignada a la causa de devoluci�n seleccionada", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
         End If
         var_ventana = 0
         frm_eliminar.Visible = False
         If Me.txt_codigo.Enabled = True Then
            txt_codigo.SetFocus
         End If
      End If
   End If
   If KeyAscii = 27 Then
      var_ventana = 0
      frm_eliminar.Visible = False
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_cantidad_GotFocus()
   txt_cantidad = ""
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
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

Private Sub txt_cliente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci�n disponible"
End Sub

Private Sub txt_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Me.txt_clave_titular <> "" Then
         lv_lista.ListItems.Clear
         If var_unidad_organizacional = "85" Or var_unidad_organizacional = "94" Then
            If var_clave_usuario_global = "U0000000314" Or var_clave_usuario_global = "U0000000390" Then
               var_cadena = "SELECT  hps.party_site_number, hcp.site_use_id AS VCHA_CLI_CLAVE_ID, hl.address1 VCHA_CLI_NOMBRE "
               var_cadena = var_cadena + " FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcp.collector_id = " + Me.txt_agente + " and hcas.cust_account_id = " + Me.txt_clave_titular + " AND hcp.site_use_id = 307953 ORDER BY hl.address1"
            Else
               var_cadena = "SELECT  hps.party_site_number, hcp.site_use_id AS VCHA_CLI_CLAVE_ID, hl.address1 VCHA_CLI_NOMBRE "
               var_cadena = var_cadena + " FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcp.collector_id = " + Me.txt_agente + " and hcas.cust_account_id = " + Me.txt_clave_titular + " ORDER BY hl.address1"
            End If
         Else
            If Me.txt_titular = "T000001052" Then
               var_cadena = "SELECT  hps.party_site_number, hcp.site_use_id AS VCHA_CLI_CLAVE_ID, hl.address1 VCHA_CLI_NOMBRE "
               var_cadena = var_cadena + " FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcp.collector_id = " + Me.txt_agente + " and hcas.cust_account_id = " + Me.txt_clave_titular + " AND hcp.site_use_id = '160667'  ORDER BY hl.address1"
            Else
               var_cadena = "SELECT  hps.party_site_number, hcp.site_use_id AS VCHA_CLI_CLAVE_ID, hl.address1 VCHA_CLI_NOMBRE "
               var_cadena = var_cadena + " FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcp.collector_id = " + Me.txt_agente + " and hcas.cust_account_id = " + Me.txt_clave_titular + " ORDER BY hl.address1"
               var_cadena = "select site_use_id as VCHA_CLI_CLAVE_ID, razon_social_cliente as VCHA_CLI_NOMBRE, party_site_number as party_site_number from xxvia_vw_clientes_bcp where account_number = '" + Me.txt_titular + "' and site_use_code = 'BILL_TO'"
            End If
         End If
         If rs.State = 1 Then
            rs.Close
         End If
         rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
               list_item.SubItems(2) = IIf(IsNull(rs!party_site_number), "", rs!party_site_number)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "Clientes"
         VAR_TIPO_LISTA = 4
         Dim var_n As Integer
         var_n = lv_lista.ListItems.Count
         If var_n > 6 Then
            lv_lista.ColumnHeaders(2).Width = 4270.71
         Else
            lv_lista.ColumnHeaders(2).Width = 4499.71
         End If
         frm_lista.Visible = True
         lv_lista.SetFocus
      Else
         MsgBox "No se selecciono un titular", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_cliente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   
   If Trim(txt_cliente) <> "" Then
      If IsNumeric(Me.txt_cliente) Then
         var_cadena = "SELECT  hcsu.price_list_id, hcsu.order_type_id,hca.cust_account_id, hcp.site_use_id AS VCHA_CLI_CLAVE_ID,hl.address1 VCHA_CLI_NOMBRE FROM hz_parties hp,hz_party_sites hps,hz_cust_accounts hca,hz_cust_acct_sites_all hcas,hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcp.site_use_id = " + Me.txt_cliente
         'var_cadena = "select * from xxvia_vw_clientes_bcp where cust_acct_site_id = " + Me.txt_cliente + " and  account_number = '" + Me.txt_titular + "' and site_use_code = 'BILL_TO'"
         rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_clave_lista_precios = IIf(IsNull(rs!price_list_id), 0, rs!price_list_id)
            If var_unidad_organizacional = "90" Then
               If Me.txt_cliente <> "8049" Then
                  var_clave_lista_precios = 9016
               End If
            End If
            If var_unidad_organizacional = "89" Or var_unidad_organizacional = "85" Or var_unidad_organizacional = "94" Then
               If Me.txt_agente <> "1028" Then
                  If Me.txt_agente = "1016" Then
                     var_clave_lista_precios = 719011
                  Else
                     var_clave_lista_precios = 9019
                  End If
               End If
            End If
            
            var_tipo_pedido = IIf(IsNull(rs!ORDER_TYPE_ID), 0, rs!ORDER_TYPE_ID)
            txt_nombre_cliente = rs!VCHA_CLI_NOMBRE
            'txt_nombre_cliente = rs!razon_social_cliente
            var_clave_titular = rs!CUST_ACCOUNT_ID
            rs.Close
            txt_establecimiento.Enabled = True
            txt_establecimiento.SetFocus
         Else
            rs.Close
            MsgBox "Clave de Cliente Incorrecta", vbOKOnly, "ATENCION"
            txt_cliente = ""
            txt_nombre_cliente = ""
         End If
      Else
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
         txt_cliente = ""
         txt_nombre_cliente = ""
      End If
   End If
End Sub

Private Sub txt_codigo_GotFocus()
'   txt_codigo = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Dim var_recontable As Integer
   Dim var_cantidad_caja As Integer
   Dim var_caja As String
   Dim var_estatus_caja As String
   Dim var_posible_caja As Integer
   Dim var_codigo_caja As String
   txt_codigo = Trim(txt_codigo)
   If KeyAscii = 13 Then
      var_codigo_barras = Me.txt_codigo
      If Mid(Me.txt_codigo, 1, 2) = "C3" Then
         var_localizador_subinventario = " "
         var_encontro = 0
         var_cantidad_leida = 1
         var_cantidad_leida_seg_nivel = 1
         var_posible_caja = 1
         var_cantidad_leida_caja = 0
      
         var_embarque = CDbl(Mid(Me.txt_codigo, 2, 6))
         var_caja = CDbl(Mid(Me.txt_codigo, 8, 3))
         strconsulta = "select * from xxvia_Tb_salidas_cajas where inte_emb_embarque = ? and inte_paq_caja = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, var_embarque)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, var_caja)
              .Parameters.Append parametro
         End With
         Set rsaux9 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rsaux9.EOF Then
            While Not rsaux9.EOF
                  Me.txt_codigo = rsaux9!SEGMENT1
                  If rsaux8.State = 1 Then
                     rsaux8.Close
                  End If
                  rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux8.EOF Then
                     var_unidad_medida = rsaux8!PRIMARY_UOM_CODE
                     var_descripcion_articulo = rsaux8!Description
                     var_inventory_item_id = rsaux8!inventory_item_id
                     var_salida_masiva = IIf(IsNull(rsaux8!attribute10), "N", rsaux8!attribute10)
                     '9019
                      var_cadena = "select * from  qp_secu_list_headers_v a, qp_list_lines_v b Where a.list_header_id = b.list_header_id and  B.product_attr_value = " + CStr(var_inventory_item_id) + " AND a.list_header_id = " + CStr(var_clave_lista_precios) + " and   product_attr_val_disp = '" + Me.txt_codigo + "' AND NVL(OPERAND,0) > 0"
                     'MsgBox var_cadena
                     If rsaux10.State = 1 Then
                        rsaux10.Close
                     End If
                     'SE QUITA PARA HACER PRUEBA DEL PRECIO
                     rsaux10.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux10.EOF Then
                        If IIf(IsNull(rsaux10!OPERAND), 0, rsaux10!OPERAND) > 0 Then
                           var_cantidad_leida = rsaux9!FLOA_SAL_CANTIDAD_LEIDA
                           'Me.txt_foco.Enabled = True
                           'Me.txt_foco.SetFocus
                           
                           
                           If var_primera_vez = True Then
                              rs.Open "select * from xxvia_tb_folios_dev_clientes WHERE MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rs.EOF Then
                                 var_numero_folio = rs(0).Value + 1
                                 Me.txt_folio = rs(0).Value + 1
                                 rsaux11.Open "update xxvia_tb_folios_dev_clientes set folio =  folio + 1 WHERE MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              Else
                                 rsaux11.Open "insert into xxvia_tb_folios_dev_clientes (folio, MOVIMIENTO) values (1,'" + var_clave_movimiento + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 var_numero_folio = 1
                                 Me.txt_folio = 1
                              End If
                              rs.Close
                              var_primera_vez = False
                           End If
                           Cadena = "select * from xxvia_tb_devoluciones_clientes where numero = " + Str(var_numero_folio) + " and codigo = '" + txt_codigo + "' and inventory_item_id = " + CStr(var_inventory_item_id) + " and localizador = '" + var_localizador_subinventario + "' AND MOVIMIENTO = '" + var_clave_movimiento + "'"
                           rs.Open Cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If rs.EOF Then
                              lbl_total = CStr(CDbl(lbl_total) + var_cantidad_leida)
                              var_inserta = False
                              rs.Close
                              If Me.txt_establecimiento = "" Then
                                 Me.txt_establecimiento = 0
            
                              End If
                              var_cadena = "insert into xxvia_tb_devoluciones_clientes (numero, organizacion, inventory_item_id, codigo, cantidad, descripcion, estatus, agente, cliente, establecimiento, titular, nombre_agente, almacen, nombre_almacen, nombre_cliente, nombre_establecimiento, referencia, usuario, maquina, fecha_inicio, unidad_medida, precio, localizador, movimiento,tipo_pedido)"
                              var_cadena = var_cadena + " values (" + CStr(var_numero_folio) + "," + var_unidad_organizacional + "," + CStr(var_inventory_item_id) + ",'" + Me.txt_codigo + "'," + CStr(var_cantidad_leida) + ",'" + Replace(var_descripcion_articulo, "'", " ") + "',''," + Me.txt_agente + "," + Me.txt_cliente + "," + Me.txt_establecimiento + "," + CStr(var_clave_titular) + ",'" + Me.txt_nombre_agente + "','" + Me.txt_almacen + "','" + Me.txt_nombre_almacen + "','" + Replace(Me.txt_nombre_cliente, "'", " ") + "','" + Replace(Me.txt_nombre_establecimiento, "'", " ") + "','" + Me.txt_referencia + "','" + var_clave_usuario_global + "', '" + fun_NombrePc + "',sysdate,'" + var_unidad_medida + "',0,'" + var_localizador_subinventario + "','" + var_clave_movimiento + "'," + CStr(var_tipo_pedido) + ")"
                              'MsgBox var_cadena
                              rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              valor = Trim(txt_codigo)
         
                              Set list_item = lv_entradas.ListItems.Add(, , Trim(txt_codigo))
                              list_item.SubItems(1) = var_descripcion_articulo
                              list_item.SubItems(2) = var_cantidad_leida
                              list_item.SubItems(3) = var_localizador_subinventario
                              var_renglon = lv_entradas.ListItems.Count
                              On Error GoTo x:
                              'clnt.MSSoapInit var_webservice_texto
                              'var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MOVIMIENTO: " + var_clave_movimiento + Me.txt_folio + "  MAQUINA: " + fun_NombrePc + Chr(13) + " USUARIO: " + var_nombre_usuario_global + Chr(13) + "FECHA Y HORA: " + VAR_FECHA_HORA + Chr(13) + "C�digo: " + Me.txt_codigo + " " + var_descripcion_articulo + Chr(13) + " Cantidad: " + CStr(var_cantidad_leida))
                              'Set clnt = Nothing
x:
                              Call ilumina_grid
                              txt_codigo = ""
                           Else
                              var_inserta = False
                              lbl_total = CStr(CDbl(lbl_total) + var_cantidad_leida)
                              rs.Close
                              rs.Open "update xxvia_tb_devoluciones_clientes set cantidad = cantidad +" + CStr(var_cantidad_leida) + " where numero = " + CStr(var_numero_folio) + " and inventory_item_id = " + CStr(var_inventory_item_id) + " and codigo = '" + Me.txt_codigo + "' and localizador = '" + var_localizador_subinventario + "' and movimiento = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              valor = Me.txt_codigo
                              var_j = 1
                              For var_j = 1 To Me.lv_entradas.ListItems.Count
                                  Me.lv_entradas.ListItems.Item(var_j).Selected = True
                                  If Me.lv_entradas.selectedItem = Me.txt_codigo And Trim(Me.lv_entradas.selectedItem.SubItems(3)) = Trim(var_localizador_subinventario) Then
                                     Me.lv_entradas.selectedItem.SubItems(2) = CDbl(Me.lv_entradas.selectedItem.SubItems(2)) + var_cantidad_leida
                                     var_renglon = var_j
                                  End If
                              Next var_j
                              Call ilumina_grid
                              On Error GoTo Z:
                              'clnt.MSSoapInit var_webservice_texto
                              'var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MOVIMIENTO: " + var_clave_movimiento + Me.txt_folio + "  MAQUINA: " + fun_NombrePc + Chr(13) + " USUARIO: " + var_nombre_usuario_global + Chr(13) + "FECHA Y HORA: " + VAR_FECHA_HORA + Chr(13) + "C�digo: " + Me.txt_codigo + " " + var_descripcion_articulo + Chr(13) + " Cantidad: " + CStr(var_cantidad_leida))
                              'Set clnt = Nothing
Z:
                              txt_codigo = ""
                           End If
                           txt_codigo.SetFocus
                           txt_codigo = ""
                        Else
                           txt_codigo = ""
                           frmmensaje.lbl_mensaje = "El art�culo no tiene precio"
                           frmmensaje.Show
                       End If
                     Else
                       txt_codigo = ""
                       frmmensaje.lbl_mensaje = "El art�culo no tiene precio"
                       frmmensaje.Show
                    End If
                     rsaux10.Close
                  Else
                       txt_codigo = ""
                       frmmensaje.lbl_mensaje = "El art�culo " + rsaux9!SEGMENT1 + " no existe"
                       frmmensaje.Show
                  End If
                  rsaux8.Close
                  rsaux9.MoveNext
            Wend
         Else
            txt_codigo = ""
            frmmensaje.lbl_mensaje = "La caja no existe"
            frmmensaje.Show
         End If
         rsaux9.Close
         
      Else
         var_localizador_subinventario = " "
         var_encontro = 0
         var_cantidad_leida = 1
         var_cantidad_leida_seg_nivel = 1
         var_posible_caja = 1
         var_cantidad_leida_caja = 0
         Dim var_tela As String
         var_tela = ""
         For var_j = 1 To Len(Me.txt_codigo)
             If Mid(Me.txt_codigo, var_j, 1) = "-" Then
                var_tela = var_tela + Mid(Me.txt_codigo, var_j, 1)
             End If
         Next var_j
         var_estatus_caja = ""
         var_codigo_caja = ""
         If Mid(Me.txt_codigo, 1, 2) = "CA" Or var_tela = "---" Then
            rs.Open "SELECT * FROM XXVIA_TB_CAJAS_PROD WHERE vcha_caj_caja_id = '" + UCase(Me.txt_codigo) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               If rsaux8.State = 1 Then
                  rsaux8.Close
               End If
               rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + IIf(IsNull(rs!VCHA_ART_ARTICULO_ID), "", rs!VCHA_ART_ARTICULO_ID) + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux8.EOF Then
                  var_estatus_caja = IIf(IsNull(rs!vcha_caj_staus), "", rs!vcha_caj_staus)
                  'var_estatus_caja = "A"
                  
                  If var_estatus_caja <> "S" Then
                     var_codigo_caja = Me.txt_codigo
                     Me.txt_codigo = IIf(IsNull(rs!VCHA_ART_ARTICULO_ID), "", rs!VCHA_ART_ARTICULO_ID)
                     var_codigo_barras = IIf(IsNull(rs!vcha_Codigo_barras), "", rs!vcha_Codigo_barras)
                     If var_codigo_barras = "" Then
                        var_codigo_barras = Me.txt_codigo + "0119"
                     End If
                     var_salida_masiva = IIf(IsNull(rsaux8!attribute10), "N", rsaux8!attribute10)
                     var_cantidad_leida_caja = rs!numb_caj_cantidad
                     strconsulta = "update xxvia_Tb_cajas_prod set vcha_caj_staus = 'S' where vcha_caj_Caja_id = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_codigo_caja)
                          .Parameters.Append parametro
                     End With
                     Set rsaux9 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                  
                  Else
                     var_posible_caja = 0
                  End If
               End If
               rsaux8.Close
            End If
            rs.Close
         End If
         If var_posible_caja = 1 Then
         var_cadena = "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, NVL(a.description,'') AS localizador, B.UNIT_WEIGHT FROM (select INVENTORY_ITEM_ID, description, cross_reference from mtl_cross_references_b) A, (select inventory_item_id, DESCRIPTION, organization_id, segment1, UNIT_WEIGHT from xxvia_system_items_b) B Where a.inventory_item_id = B.inventory_item_id AND B.organization_id = " + var_unidad_organizacional + " AND CROSS_REFERENCE       = '" + Me.txt_codigo + "'"
         x = 0
         If x = 0 Then
            strconsulta = "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, NVL(a.description,'') AS localizador, B.UNIT_WEIGHT, a.cantidad FROM (select INVENTORY_ITEM_ID, description, cross_reference, nvl(attribute1,1) as cantidad from mtl_cross_references_b) A, (select inventory_item_id, DESCRIPTION, organization_id, segment1, UNIT_WEIGHT from xxvia_system_items_b) B Where a.inventory_item_id = B.inventory_item_id AND B.organization_id = ? AND CROSS_REFERENCE = ?"
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
         var_cantidad_leida = 1
         If Not rsaux8.EOF Then
            var_peso = IIf(IsNull(rsaux8!UNIT_WEIGHT), 0, rsaux8!UNIT_WEIGHT)
            var_cantidad_leida_seg_nivel = IIf(IsNull(rsaux8!cantidad), 1, rsaux8!cantidad)
            'If IIf(IsNull(rsaux8!localizador), "", rsaux8!localizador) <> "" Then
            '   var_localizador_subinventario = txt_almacen + IIf(IsNull(rsaux8!localizador), "", rsaux8!localizador)
            '   If var_localizador_subinventario <> "" Then
            '       Me.txt_codigo = rsaux8!SEGMENT1
            '   End If
            'Else
               Me.txt_codigo = rsaux8!SEGMENT1
            'End If
         End If
         rsaux8.Close
      
      
      
      
      
      
      
         
         'rsaux8.Open "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, nvl(a.description,'') as localizador FROM mtl_cross_references_v A, xxvia_system_items_b B WHERE A.inventory_item_id = B.inventory_item_id AND B.organization_id = " + var_unidad_organizacional + " AND CROSS_REFERENCE = '" + Me.txt_codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         'If Not rsaux8.EOF Then
         '   Me.txt_codigo = rsaux8!SEGMENT1
         'Else
         '   rsaux9.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
         '   If Not rsaux9.EOF Then
         '      Me.txt_codigo = rsaux9!SEGMENT1
         '   Else
         '      Me.txt_codigo = ""
         '   End If
         '   rsaux9.Close
         'End If
         'rsaux8.Close
      
         If Trim(Me.txt_codigo) <> "" Then
            rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
            
            
               
               var_unidad_medida = rsaux8!PRIMARY_UOM_CODE
               var_descripcion_articulo = rsaux8!Description
               var_inventory_item_id = rsaux8!inventory_item_id
               var_salida_masiva = IIf(IsNull(rsaux8!attribute10), "N", rsaux8!attribute10)
               '9019
               If Me.txt_cliente = "1065198" Or Me.txt_cliente = "1057429" Or Me.txt_cliente = "1066462" Then
                  var_clave_lista_precios = 9019
               End If
               var_cadena = "select * from  qp_secu_list_headers_v a, qp_list_lines_v b Where a.list_header_id = b.list_header_id and  B.product_attr_value = " + CStr(var_inventory_item_id) + " AND a.list_header_id = " + CStr(var_clave_lista_precios) + " and   product_attr_val_disp = '" + Me.txt_codigo + "' AND NVL(OPERAND,0) > 0"
               'MsgBox var_cadena
               rsaux10.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux10.EOF Then
                  If IIf(IsNull(rsaux10!OPERAND), 0, rsaux10!OPERAND) > 0 Then
                     If var_unidad_organizacional = "900" Then
                        var_salida_masiva = "Y"
                     End If
                     'If  var_unidad_organizacional = "85" Or var_unidad_organizacional = "94" Then
                     '   var_salida_masiva = "Y"
                     'End If
                     If var_clave_usuario_global = "U0000000635" Then
                        var_salida_masiva = "Y"
                     End If
                     If var_cantidad_leida_caja = 0 Then
                        If var_cantidad_leida_seg_nivel = 1 Then
                           If var_salida_masiva = "Y" Then
                              var_codigo_global = Me.txt_codigo
                              frmoracle_cantidad.Show 1
                              var_cantidad_leida = var_cantidad_global
                              Me.txt_codigo = var_codigo_global
                           Else
                              var_cantidad_leida = 1
                           End If
                           Me.txt_foco.Enabled = True
                           Me.txt_foco.SetFocus
                        Else
                           var_cantidad_leida = var_cantidad_leida_seg_nivel
                           Me.txt_foco.Enabled = True
                           Me.txt_foco.SetFocus
                        End If
                     Else
                        var_cantidad_leida = var_cantidad_leida_caja
                        Me.txt_foco.Enabled = True
                        Me.txt_foco.SetFocus
                     End If
                  Else
                     txt_codigo = ""
                     frmmensaje.lbl_mensaje = "El art�culo no tiene precio"
                     frmmensaje.Show
                     If var_codigo_caja <> "" Then
                        strconsulta = "update xxvia_Tb_cajas_prod set vcha_caj_staus = 'A' where vcha_caj_Caja_id = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_codigo_caja)
                             .Parameters.Append parametro
                        End With
                        Set rsaux9 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                     End If
                  
                  
                  End If
               Else
                  txt_codigo = ""
                  frmmensaje.lbl_mensaje = "El art�culo no se encuentra en la lista de precios del cliente"
                  frmmensaje.Show
                  If var_codigo_caja <> "" Then
                     strconsulta = "update xxvia_Tb_cajas_prod set vcha_caj_staus = 'A' where vcha_caj_Caja_id = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_codigo_caja)
                          .Parameters.Append parametro
                     End With
                     Set rsaux9 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                  End If
               
               End If
               rsaux10.Close
            Else
               txt_codigo = ""
               frmmensaje.lbl_mensaje = "Error en c�digo"
               frmmensaje.Show
            End If
            rsaux8.Close
         Else
            If var_localizador = 2 And Me.txt_codigo = "" Then
               txt_codigo = ""
               frmmensaje.lbl_mensaje = "El art�culo necesita localizador"
               frmmensaje.Show
            Else
               txt_codigo = ""
               frmmensaje.lbl_mensaje = "El art�culo no existe"
               frmmensaje.Show
            End If
         End If
         Else
            Me.txt_codigo = ""
            frmmensaje.lbl_mensaje = "La caja ya habia sido leida"
            frmmensaje.Show
         End If
      End If
   End If
End Sub

Private Sub txt_establecimiento_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci�n disponible"
End Sub

Private Sub txt_establecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      'rs.Open "SELECT hcp.site_use_id AS  VCHA_ESB_ESTABLECIMIENTO_ID, hl.address1 VCHA_ESB_NOMBRE FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'SHIP_TO' AND hcp.collector_id = " + Me.txt_agente + " AND hca.cust_account_id = " + CStr(var_clave_titular), cnnoracle_4, adOpenDynamic, adLockOptimistic
      If var_unidad_organizacional = "85" Or var_unidad_organizacional = "94" Then
         If var_clave_usuario_global = "U0000000314" Or var_clave_usuario_global = "U0000000390" Then
            rs.Open "SELECT site_use_id as VCHA_ESB_ESTABLECIMIENTO_ID, location as party_site_number, 'CANTIA' AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all WHERE LOCATION = 'E000005731-1'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         Else
            rs.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id  = ps.party_site_id AND ps.location_id  = lo.location_id  AND csu.site_use_code = 'SHIP_TO' AND cas.cust_account_id     = " + CStr(var_clave_titular), cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
      Else
         If Me.txt_cliente <> "7821" Then
            If Me.txt_cliente = "160667" Then
               If rs.State = 1 Then
                  rs.Close
               End If
               rs.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id  = ps.party_site_id AND ps.location_id  = lo.location_id  AND csu.site_use_code = 'SHIP_TO' AND  SITE_USE_ID = 326595", cnnoracle_4, adOpenDynamic, adLockOptimistic
            Else
               rs.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id  = ps.party_site_id AND ps.location_id  = lo.location_id  AND csu.site_use_code = 'SHIP_TO' AND cas.cust_account_id     = " + CStr(var_clave_titular), cnnoracle_4, adOpenDynamic, adLockOptimistic
               '''aqui
               'rs.Open "select cust_acct_site_id as VCHA_ESB_ESTABLECIMIENTO_ID, razon_social_cliente as VCHA_eSB_NOMBRE, party_site_number as party_site_number from xxvia_vw_clientes_bcp where account_number = '" + Me.txt_titular + "' and site_use_code = 'SHIP_TO'"
            End If
         Else
            rs.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id  = ps.party_site_id AND ps.location_id  = lo.location_id  AND csu.site_use_code = 'SHIP_TO' AND cas.cust_account_id     = " + CStr(var_clave_titular) + " and ps.party_site_number = 'EC000023687'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
         'rs.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'SHIP_TO' AND csu.LOCATION = ps.party_site_number and cas.cust_account_id = " + CStr(var_clave_titular), cnnoracle_4, adOpenDynamic, adLockOptimistic
      End If
      
      'MsgBox var_clave_titular
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ESB_ESTABLECIMIENTO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_esb_nombre), "", rs!vcha_esb_nombre)
            list_item.SubItems(2) = IIf(IsNull(rs!party_site_number), "", rs!party_site_number)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Establecimientos"
      VAR_TIPO_LISTA = 3
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_nombre_establecimiento.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txt_establecimiento_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_establecimiento) <> "" Then
      'rs.Open "SELECT hcp.site_use_id AS  VCHA_ESB_ESTABLECIMIENTO_ID, hl.address1 VCHA_ESB_NOMBRE FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas,hz_cust_site_uses_all hcsu,hz_locations hl, hr_operating_units hr,hz_customer_profiles hcp, ar_collectors arc,ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id=hps.party_id AND hps.party_site_id=hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'SHIP_TO' AND hcp.collector_id = " + Me.txt_agente + " AND hcp.site_use_id = " + CStr(Me.txt_establecimiento), cnnoracle_4, adOpenDynamic, adLockOptimistic
      If var_unidad_organizacional = "85" Or var_unidad_organizacional = "94" Then
         If var_clave_usuario_global = "U0000000314" Or var_clave_usuario_global = "U0000000390" Then
            rs.Open "SELECT site_use_id as VCHA_ESB_ESTABLECIMIENTO_ID, location as party_site_number, 'CANTIA' AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all WHERE LOCATION = 'E000005731-1'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         Else
            rs.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id  = ps.party_site_id AND ps.location_id  = lo.location_id  AND csu.site_use_code = 'SHIP_TO' AND cas.cust_account_id     = " + CStr(var_clave_titular) + " And csu.site_use_id = " + Me.txt_establecimiento, cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
      Else
         rs.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id  = ps.party_site_id AND ps.location_id  = lo.location_id  AND csu.site_use_code = 'SHIP_TO' AND cas.cust_account_id     = " + CStr(var_clave_titular) + " And csu.site_use_id = " + Me.txt_establecimiento, cnnoracle_4, adOpenDynamic, adLockOptimistic
         'rs.Open "select cust_acct_site_id as VCHA_ESB_ESTABLECIMIENTO_ID, razon_social_cliente as VCHA_eSB_NOMBRE, party_site_number as party_site_number from xxvia_vw_clientes_bcp where account_number = '" + Me.txt_titular + "' and site_use_code = 'SHIP_TO' AND CUST_ACCT_SITE_ID = " + Me.txt_establecimiento, cnnoracle_4, adOpenDynamic, adLockOptimistic
      End If
      
      If Not rs.EOF Then
         txt_nombre_establecimiento = rs!vcha_esb_nombre
         rs.Close
         txt_establecimiento.Enabled = False
         txt_referencia.Enabled = True
         txt_referencia.SetFocus
      Else
         rs.Close
         MsgBox "Clave de Establecimiento Incorrecta", vbOKOnly, "ATENCION"
         txt_establecimiento = ""
         txt_nombre_establecimiento = ""
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   Dim var_inserta As Boolean
   Dim var_factura As Integer
   Dim var_posible_cliente As Boolean
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   rsaux.Open "ALTER SESSION SET NLS_DATE_FORMAT = 'DD/MM/YYYY HH24:MI:SS'", cnnoracle_4, adOpenDynamic, adLockOptimistic

   rs.Open "SELECT TO_CHAR(SYSDATE,'dd/mm/yyyy hh24:mi:ss') AS FECHA FROM DUAL", cnnoracle_4, adOpenDynamic, adLockOptimistic
   VAR_FECHA_HORA = rs(0).Value
   rs.Close
   
   If Trim(txt_codigo.Text) <> "" Then
      If var_clave_movimiento = "VDI" Then
         If (var_unidad_organizacional = 85 Or var_unidad_organizacional = 94) And Me.txt_cliente = "307953" Then
            var_posible_existencia = 1
         Else
            var_cantidad_leida_VDI = 0
            For var_j = 1 To Me.lv_entradas.ListItems.Count
                  Me.lv_entradas.ListItems.Item(var_j).Selected = True
                  If Me.lv_entradas.selectedItem = Me.txt_codigo Then
                     var_cantidad_leida_VDI = var_cantidad_leida_VDI + CDbl(Me.lv_entradas.selectedItem.SubItems(2))
                  End If
            Next var_j
            
            strconsulta = "select * from Xxvia_vw_existencias_inv where organization_id = ? and subinventory_code = ? and segment1 = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_unidad_organizacional))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_almacen)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                 .Parameters.Append parametro
            End With
            Set rsaux = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            var_cantidad_leida_VDI = var_cantidad_leida_VDI + var_cantidad_leida
            If Not rsaux.EOF Then
               var_disponible = IIf(IsNull(rsaux!Disponible), 0, rsaux!Disponible)
            Else
               var_disponible = 0
            End If
            If var_cantidad_leida_VDI <= var_disponible Then
               var_posible_existencia = 1
            Else
               var_posible_existencia = 0
            End If
            rsaux.Close
         End If
      Else
         var_posible_existencia = 1
      End If

      If var_posible_existencia = 1 Then
         If var_primera_vez = True Then
            rs.Open "select * from xxvia_tb_folios_dev_clientes WHERE MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_numero_folio = rs(0).Value + 1
               Me.txt_folio = rs(0).Value + 1
               rsaux10.Open "update xxvia_tb_folios_dev_clientes set folio =  folio + 1 WHERE MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            Else
               rsaux10.Open "insert into xxvia_tb_folios_dev_clientes (folio, MOVIMIENTO) values (1,'" + var_clave_movimiento + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_numero_folio = 1
               Me.txt_folio = 1
            End If
            rs.Close
            var_primera_vez = False
         End If
         Cadena = "select * from xxvia_tb_devoluciones_clientes where numero = " + Str(var_numero_folio) + " and codigo = '" + txt_codigo + "' and inventory_item_id = " + CStr(var_inventory_item_id) + " and localizador = '" + var_localizador_subinventario + "' AND MOVIMIENTO = '" + var_clave_movimiento + "'"
         rs.Open Cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            lbl_total = CStr(CDbl(lbl_total) + var_cantidad_leida)
            var_inserta = False
            rs.Close
            If Me.txt_establecimiento = "" Then
               Me.txt_establecimiento = 0
            
            End If
            var_cadena = "insert into xxvia_tb_devoluciones_clientes (numero, organizacion, inventory_item_id, codigo, cantidad, descripcion, estatus, agente, cliente, establecimiento, titular, nombre_agente, almacen, nombre_almacen, nombre_cliente, nombre_establecimiento, referencia, usuario, maquina, fecha_inicio, unidad_medida, precio, localizador, movimiento,tipo_pedido)"
            var_cadena = var_cadena + " values (" + CStr(var_numero_folio) + "," + var_unidad_organizacional + "," + CStr(var_inventory_item_id) + ",'" + Me.txt_codigo + "'," + CStr(var_cantidad_leida) + ",'" + Replace(var_descripcion_articulo, "'", " ") + "',''," + Me.txt_agente + "," + Me.txt_cliente + "," + Me.txt_establecimiento + "," + CStr(var_clave_titular) + ",'" + Me.txt_nombre_agente + "','" + Me.txt_almacen + "','" + Me.txt_nombre_almacen + "','" + Replace(Me.txt_nombre_cliente, "'", " ") + "','" + Replace(Me.txt_nombre_establecimiento, "'", " ") + "','" + Me.txt_referencia + "','" + var_clave_usuario_global + "', '" + fun_NombrePc + "',sysdate,'" + var_unidad_medida + "',0,'" + var_localizador_subinventario + "','" + var_clave_movimiento + "'," + CStr(var_tipo_pedido) + ")"
            'MsgBox var_cadena
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            valor = Trim(txt_codigo)
         
         '1
            If var_unidad_organizacional = 90 Then
               strconsulta = "call XXVIA_TB_TRANSACCIONS_CB (?,?,?,?,?,?,?,?,?) "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_almacen)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_clave_movimiento)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, CStr(Me.txt_folio))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_codigo_barras)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, Me.txt_codigo)
                    .Parameters.Append parametro
                    If var_clave_movimiento = "VDI" Then
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_cantidad_leida * -1)
                       .Parameters.Append parametro
                    Else
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_cantidad_leida)
                       .Parameters.Append parametro
                    End If
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 500, var_nombre_usuario_global)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 0)
                    .Parameters.Append parametro
               End With
               Set rsaux17 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
            End If
         
         
         
         
            Set list_item = lv_entradas.ListItems.Add(, , Trim(txt_codigo))
            list_item.SubItems(1) = var_descripcion_articulo
            list_item.SubItems(2) = var_cantidad_leida
            list_item.SubItems(3) = var_localizador_subinventario
            var_renglon = lv_entradas.ListItems.Count
            On Error GoTo x:
            clnt.MSSoapInit var_webservice_texto
            var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MOVIMIENTO: " + var_clave_movimiento + Me.txt_folio + "  MAQUINA: " + fun_NombrePc + Chr(13) + " USUARIO: " + var_nombre_usuario_global + Chr(13) + "FECHA Y HORA: " + VAR_FECHA_HORA + Chr(13) + "C�digo: " + Me.txt_codigo + " " + var_descripcion_articulo + Chr(13) + " Cantidad: " + CStr(var_cantidad_leida))
            Set clnt = Nothing
x:
            Call ilumina_grid
            txt_codigo = ""
            
         
         Else
            var_inserta = False
            lbl_total = CStr(CDbl(lbl_total) + var_cantidad_leida)
            rs.Close
            rs.Open "update xxvia_tb_devoluciones_clientes set cantidad = cantidad +" + CStr(var_cantidad_leida) + " where numero = " + CStr(var_numero_folio) + " and inventory_item_id = " + CStr(var_inventory_item_id) + " and codigo = '" + Me.txt_codigo + "' and localizador = '" + var_localizador_subinventario + "' and movimiento = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            
            If var_unidad_organizacional = 90 Then
               strconsulta = "call XXVIA_TB_TRANSACCIONS_CB (?,?,?,?,?,?,?,?,?) "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_almacen)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_clave_movimiento)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 20, CStr(Me.txt_folio))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_codigo_barras)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, Me.txt_codigo)
                    .Parameters.Append parametro
                    If var_clave_movimiento = "VDI" Then
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_cantidad_leida * -1)
                       .Parameters.Append parametro
                    Else
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_cantidad_leida)
                       .Parameters.Append parametro
                    End If
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 500, var_nombre_usuario_global)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 0)
                    .Parameters.Append parametro
                    
               End With
               Set rsaux17 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
            End If
            
            
            
            
            valor = Me.txt_codigo
            var_j = 1
            For var_j = 1 To Me.lv_entradas.ListItems.Count
                  Me.lv_entradas.ListItems.Item(var_j).Selected = True
                  If Me.lv_entradas.selectedItem = Me.txt_codigo And Trim(Me.lv_entradas.selectedItem.SubItems(3)) = Trim(var_localizador_subinventario) Then
                     Me.lv_entradas.selectedItem.SubItems(2) = CDbl(Me.lv_entradas.selectedItem.SubItems(2)) + var_cantidad_leida
                     var_renglon = var_j
                  End If
            Next var_j
            Call ilumina_grid
            On Error GoTo Z:
            clnt.MSSoapInit var_webservice_texto
            var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MOVIMIENTO: " + var_clave_movimiento + Me.txt_folio + "  MAQUINA: " + fun_NombrePc + Chr(13) + " USUARIO: " + var_nombre_usuario_global + Chr(13) + "FECHA Y HORA: " + VAR_FECHA_HORA + Chr(13) + "C�digo: " + Me.txt_codigo + " " + var_descripcion_articulo + Chr(13) + " Cantidad: " + CStr(var_cantidad_leida))
            Set clnt = Nothing
Z:
            txt_codigo = ""
         End If
         txt_codigo.SetFocus
         txt_codigo = ""
      Else
         frmmensaje.lbl_articulo = Me.txt_codigo + " " + var_descripcion_articulo
         frmmensaje.lbl_mensaje = "La cantidad supera a la existencia disponible. Leido: " + CStr(var_cantidad_leida_VDI) + ", disponible: " + CStr(var_disponible) + "."
         frmmensaje.Show
         Me.txt_codigo = ""
      End If
   End If
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Me.txt_establecimiento.Enabled = True
      Me.txt_establecimiento.SetFocus
      Me.txt_cliente.Enabled = False
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txt_nombre_agente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci�n disponible"
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   KeyCode = 0
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "SELECT  distinct arc.collector_id as vcha_age_agente_id, arc.name as vcha_age_nombre  FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcas.org_id = " + var_empresa + " order by arc.name", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Agentes"
      VAR_TIPO_LISTA = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_agente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_almacen_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci�n disponible"
End Sub

Private Sub txt_nombre_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   KeyCode = 0
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rsaux.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_alm_almacen_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Almacenes"
      VAR_TIPO_LISTA = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txt_agente.Enabled = True Then
         txt_agente.SetFocus
      End If
   End If
End Sub

Private Sub txt_nombre_almacen_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_cliente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci�n disponible"
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   KeyCode = 0
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "  SELECT  hcp.site_use_id AS VCHA_CLI_CLAVE_ID, hl.address1 VCHA_CLI_NOMBRE FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcp.collector_id = " + Me.txt_agente, cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
      VAR_TIPO_LISTA = 4
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If txt_referencia.Enabled = True Then
         txt_referencia.SetFocus
      End If
   End If
End Sub

Private Sub txt_nombre_cliente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_establecimiento_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci�n disponible"
End Sub

Private Sub txt_nombre_establecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   KeyCode = 0
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "SELECT hcp.site_use_id AS  VCHA_ESB_ESTABLECIMIENTO_ID, hl.address1 VCHA_ESB_NOMBRE FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'SHIP_TO' AND hcp.collector_id = " + Me.txt_agente + " AND hca.cust_account_id = " + CStr(var_clave_titular)
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ESB_ESTABLECIMIENTO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_esb_nombre), "", rs!vcha_esb_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Establecimientos"
      VAR_TIPO_LISTA = 3
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_establecimiento_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_titular_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "SELECT  distinct hp.party_name as nombre_titular,  account_number as vcha_tit_titular_id, hcas.cust_account_id AS VCHA_CLI_CLAVE_ID FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors Arc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND hcp.collector_id = " + Me.txt_agente, cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tit_titular_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!NOMBRE_TITULAR)
            list_item.SubItems(2) = IIf(IsNull(rs!CUST_ACCOUNT_ID), "", rs!CUST_ACCOUNT_ID)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Titulares"
      VAR_TIPO_LISTA = 5
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Me.cmd_aceptar_pedidos.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.frm_pasar_todo.Visible = False
   End If
End Sub

Private Sub txt_referencia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_referencia) <> "" Then
         txt_codigo.Enabled = True
         txt_codigo.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_pasar_todo.Visible = False
   End If
End Sub

Private Sub txt_referencia_LostFocus()
      If Trim(txt_referencia) <> "" Then
         txt_codigo.Enabled = True
         txt_referencia.Enabled = False
         txt_codigo.SetFocus
      Else
         MsgBox "Debe de introducir una referencia", vbOKOnly, "ATENCION"
      End If
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_numero.SetFocus
   End If
End Sub

Private Sub txt_titular_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_unidad_organizacional = "85" Or var_unidad_organizacional = "94" Then
         If var_clave_usuario_global = "U0000000314" Or var_clave_usuario_global = "U0000000390" Then
            rs.Open "SELECT  distinct hp.party_name as nombre_titular,  account_number as vcha_tit_titular_id, hcas.cust_account_id AS VCHA_CLI_CLAVE_ID FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors Arc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND hcp.collector_id = " + Me.txt_agente + " AND hcas.cust_account_id = 4268 ORDER BY hp.party_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
         Else
            rs.Open "SELECT  distinct hp.party_name as nombre_titular,  account_number as vcha_tit_titular_id, hcas.cust_account_id AS VCHA_CLI_CLAVE_ID FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors Arc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND hcp.collector_id = " + Me.txt_agente + " ORDER BY hp.party_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
      Else
        If Me.txt_agente = "1016" Then
           'rs.Open "SELECT  distinct hp.party_name as nombre_titular,  account_number as vcha_tit_titular_id, hcas.cust_account_id AS VCHA_CLI_CLAVE_ID FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors Arc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND hcp.collector_id = " + Me.txt_agente + " AND account_number = 'T000001052'  ORDER BY hp.party_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
           rs.Open "SELECT  distinct hp.party_name as nombre_titular,  account_number as vcha_tit_titular_id, hcas.cust_account_id AS VCHA_CLI_CLAVE_ID FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors Arc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND hcp.collector_id = " + Me.txt_agente + " ORDER BY hp.party_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
        Else
           If rs.State = 1 Then
              rs.Close
           End If
           rs.Open "SELECT  distinct hp.party_name as nombre_titular,  account_number as vcha_tit_titular_id, hcas.cust_account_id AS VCHA_CLI_CLAVE_ID FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors Arc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND hcp.collector_id = " + Me.txt_agente + " ORDER BY hp.party_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
        End If
      End If
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tit_titular_id)
            list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_TITULAR), "", rs!NOMBRE_TITULAR)
            list_item.SubItems(2) = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Titulares"
      VAR_TIPO_LISTA = 5
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_titular_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Me.txt_cliente.Enabled = True
      Me.txt_cliente.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If

End Sub

Private Sub txt_titular_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(Me.txt_titular) <> "" Then
      var_cadena = "SELECT  distinct hp.party_name as nombre_titular,  account_number as vcha_tit_titular_id, hcas.cust_account_id AS VCHA_CLI_CLAVE_ID FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors Arc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND hcp.collector_id = " + txt_agente + " and hcas.cust_account_id = " + Me.txt_clave_titular
      rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_titular = rs!vcha_tit_titular_id
         Me.txt_nombre_titular = rs!NOMBRE_TITULAR
         var_clave_titular = rs!vcha_cli_clave_id
         'MsgBox var_clave_titular
         rs.Close
         txt_cliente.Enabled = True
         txt_cliente.SetFocus
         Me.txt_titular.Enabled = False
      Else
         rs.Close
         MsgBox "Clave de titular Incorrecta", vbOKOnly, "ATENCION"
         txt_cliente = ""
         txt_nombre_cliente = ""
      End If
   End If
End Sub
