VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_devoluciones_clientes_sello 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Devoluciones de clientes"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   360
      TabIndex        =   10
      Top             =   1200
      Width           =   7305
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1980
         Left            =   30
         TabIndex        =   67
         Top             =   390
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
         Left            =   30
         TabIndex        =   11
         Top             =   120
         Width           =   7230
      End
   End
   Begin VB.CommandButton cmd_asignar_causa_devolucion 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Picture         =   "frmoracle_devoluciones_clientes_sello.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "Asignar causa de devolución"
      Top             =   720
      Width           =   330
   End
   Begin VB.Frame Frame6 
      Height          =   1035
      Left            =   6240
      TabIndex        =   70
      Top             =   3040
      Width           =   2040
      Begin VB.TextBox txt_limite_superior 
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         TabIndex        =   72
         Top             =   585
         Width           =   1125
      End
      Begin VB.TextBox txt_limite_inferior 
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         TabIndex        =   71
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Superior:"
         Height          =   195
         Left            =   120
         TabIndex        =   74
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Inferior:"
         Height          =   195
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.Frame frm_pasar_todo 
      Height          =   90
      Left            =   5520
      TabIndex        =   2
      Top             =   2310
      Width           =   990
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
         Picture         =   "frmoracle_devoluciones_clientes_sello.frx":0102
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
         Picture         =   "frmoracle_devoluciones_clientes_sello.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   420
         Width           =   330
      End
      Begin VB.Label Label5 
         BackColor       =   &H000000C0&
         Caption         =   " Folio a Devolver"
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
      Height          =   195
      Left            =   1515
      TabIndex        =   63
      Top             =   735
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   1455
      TabIndex        =   60
      Top             =   735
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.TextBox txt_clave_titular 
      Height          =   285
      Left            =   5490
      TabIndex        =   59
      Text            =   "Text1"
      Top             =   705
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Frame Frame4 
      Height          =   1035
      Left            =   6255
      TabIndex        =   54
      Top             =   2018
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
         Height          =   240
         Left            =   60
         TabIndex        =   56
         Top             =   175
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
         Height          =   240
         Left            =   60
         TabIndex        =   55
         Top             =   555
         Width           =   1830
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3120
      Left            =   90
      TabIndex        =   43
      Top             =   4110
      Width           =   8205
      Begin VB.CommandButton Command3 
         Caption         =   "Pasar todos"
         Height          =   375
         Left            =   4320
         TabIndex        =   78
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmd_movimiento_masivo 
         Caption         =   "movimiento_masivo"
         Height          =   195
         Left            =   240
         TabIndex        =   64
         Top             =   450
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.CommandButton cmd_factura 
         Caption         =   "Factura"
         Height          =   195
         Left            =   420
         TabIndex        =   62
         Top             =   450
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.CommandButton cmd_devolucion 
         Caption         =   "Devolución"
         Height          =   195
         Left            =   330
         TabIndex        =   61
         Top             =   450
         Visible         =   0   'False
         Width           =   75
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
         TabIndex        =   48
         Top             =   405
         Width           =   2640
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   1785
         TabIndex        =   45
         Top             =   1755
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   60
            MaxLength       =   10
            TabIndex        =   46
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
            TabIndex        =   47
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
         Left            =   5790
         TabIndex        =   44
         Top             =   465
         Width           =   1890
      End
      Begin MSComctlLib.ListView lv_entradas 
         Height          =   2010
         Left            =   60
         TabIndex        =   66
         Top             =   1080
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   3545
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2478
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7232
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Enviado"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Recibido"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Localizador"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   585
         Width           =   1395
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Lectura de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   51
         Top             =   120
         Width           =   8130
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   5040
         TabIndex        =   50
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
         Left            =   4560
         TabIndex        =   49
         Top             =   420
         Width           =   3645
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   30
      TabIndex        =   42
      Top             =   570
      Width           =   8250
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   10305
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   2910
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      Height          =   960
      Index           =   0
      Left            =   6240
      TabIndex        =   19
      Top             =   1080
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
         TabIndex        =   20
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
         TabIndex        =   21
         Top             =   120
         Width           =   1980
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7920
      Picture         =   "frmoracle_devoluciones_clientes_sello.frx":0396
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Salir"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   735
      Picture         =   "frmoracle_devoluciones_clientes_sello.frx":09D0
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmoracle_devoluciones_clientes_sello.frx":0AD2
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Buscar Movimiento Alt + B"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmoracle_devoluciones_clientes_sello.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Nuevo Movimiento Alt + N"
      Top             =   720
      Width           =   330
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   1575
      TabIndex        =   12
      Top             =   1080
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         MaxLength       =   10
         TabIndex        =   13
         Top             =   495
         Width           =   2775
      End
      Begin VB.Label lbl_tipo_busqueda 
         BackColor       =   &H000000C0&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   14
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
   Begin VB.PictureBox ImageList 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   30
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   30
      TabIndex        =   53
      Top             =   975
      Width           =   8250
   End
   Begin VB.Frame Frame3 
      Height          =   3030
      Index           =   1
      Left            =   75
      TabIndex        =   22
      Top             =   1080
      Width           =   6120
      Begin VB.TextBox txt_sello 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3960
         TabIndex        =   76
         Top             =   1080
         Width           =   2085
      End
      Begin VB.TextBox txt_folio_Sello 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1290
         TabIndex        =   68
         Top             =   1080
         Width           =   1605
      End
      Begin VB.TextBox txt_nombre_titular 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1545
         Width           =   3720
      End
      Begin VB.TextBox txt_titular 
         Height          =   315
         Left            =   1290
         TabIndex        =   32
         Top             =   1545
         Width           =   1005
      End
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   1290
         TabIndex        =   34
         Top             =   1890
         Width           =   1005
      End
      Begin VB.TextBox txt_almacen 
         Height          =   315
         Left            =   1290
         TabIndex        =   31
         Top             =   420
         Width           =   1005
      End
      Begin VB.TextBox txt_establecimiento 
         Height          =   315
         Left            =   1290
         TabIndex        =   30
         Top             =   2235
         Width           =   1005
      End
      Begin VB.TextBox txt_referencia 
         Height          =   315
         Left            =   1290
         MaxLength       =   20
         TabIndex        =   29
         Top             =   2610
         Width           =   4365
      End
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   1290
         TabIndex        =   28
         Top             =   750
         Width           =   1005
      End
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   315
         Left            =   2325
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   405
         Width           =   3690
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   750
         Width           =   3720
      End
      Begin VB.TextBox txt_nombre_establecimiento 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2235
         Width           =   3720
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1890
         Width           =   3720
      End
      Begin VB.CommandButton cmd_pasar_todo 
         Height          =   330
         Left            =   5700
         Picture         =   "frmoracle_devoluciones_clientes_sello.frx":0CD6
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Pasar una factura"
         Top             =   2565
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sello:"
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
         Index           =   7
         Left            =   3240
         TabIndex        =   77
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
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
         Index           =   6
         Left            =   90
         TabIndex        =   69
         Top             =   1155
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   58
         Top             =   1605
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   40
         Top             =   1950
         Width           =   525
      End
      Begin VB.Label label 
         BackColor       =   &H000000C0&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   39
         Top             =   120
         Width           =   6045
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   495
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   37
         Top             =   2265
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Referencia:"
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   36
         Top             =   2610
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   35
         Top             =   840
         Width           =   555
      End
   End
   Begin VB.Label lblnombremovimiento 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   30
      TabIndex        =   57
      Top             =   105
      Width           =   8325
   End
End
Attribute VB_Name = "frmoracle_devoluciones_clientes_sello"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_devolucion_costales As Integer
Dim var_tipo_pedido As Integer
Dim var_localizador_subinventario As String
Dim var_localizador As Integer
Dim var_año As Integer
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
Dim var_tipo_busqueda As Integer
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
          lv_entradas.ListItems.Item(var_i).ListSubItems(3).Bold = True
          lv_entradas.ListItems.Item(var_i).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H8000&
       Else
          lv_entradas.ListItems.Item(var_i).Bold = False
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).Bold = False
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).Bold = False
          lv_entradas.ListItems.Item(var_i).ListSubItems(3).Bold = False
          lv_entradas.ListItems.Item(var_i).ForeColor = &H80000012
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000012
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000012
          lv_entradas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000012
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_entradas.ListItems.Item(var_renglon).Selected = True
      lv_entradas.selectedItem.EnsureVisible
   End If
   'If lv_entradas.ListItems.Count > 11 Then
   '   lv_entradas.ColumnHeaders(2).Width = 5050.22
   'Else
   '   lv_entradas.ColumnHeaders(2).Width = 5300.22
   'End If
   
   lv_entradas.Refresh
   
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
       rsaux.Close
    Else
       MsgBox "No se a seleccionado un almacén destino", vbOKOnly, "ATENCION"
    End If
    Me.frm_pasar_todo.Visible = False
End Sub

Private Sub cmd_asignar_causa_devolucion_Click()
      rs.Open "select * from xxvia_tb_Devoluciones_clientes where numero = " + CStr(var_numero_folio) + " and organizacion = " + var_unidad_organizacional + " AND MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_estatus_movimiento = IIf(IsNull(rs!estatus), "", rs!estatus)
      End If
      rs.Close
      If var_clave_movimiento = "DC" Or var_clave_movimiento = "DCS" Then
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
   'rsaux11.Open "select segment1, SUM(shipped_quantity) as cantidad from oe_order_headers_all oha, oe_order_lines_all ola, xxvia_system_items_b b where oha.header_id = ola.header_id and oha.ship_from_org_id = b.organization_id and ola.inventory_item_id = b.inventory_item_id and order_number = 72037 AND shipped_quantity IS NOT NULL GROUP BY SEGMENT1", cnnoracle_4, adOpenDynamic, adLockOptimistic
   
   rsaux11.Open "select segment1 AS CODIGO, sum(floa_sal_cantidad_leida) as Cantidad from xxvia_Tb_Salidas_cajas where source_header_number in (362073) and floa_sal_Cantidad_leida > 0 group by segment1 order by segment1", cnnoracle_4, adOpenDynamic, adLockOptimistic
   rsaux12.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   'rsaux11.Open "select B.SEGMENT1 as codigo, SHIPPED_QUANTITY as cantidad from wsh_deliverables_v A, XXVIA_SYSTEM_ITEMS_B B where source_header_number = 305756  AND A.INVENTORY_ITEM_ID = B.INVENTORY_ITEM_ID  AND B.ORGANIZATION_ID = 93 AND SHIPPED_QUANTITY > 0"
   var_Cadena_faltantes = ""
   Me.txt_codigo = rsaux11!codigo
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
            Me.txt_codigo = rsaux11!codigo
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
      MsgBox "Faltan los siguientes códigos " + var_Cadena_faltantes, vbOKOnly, "ATENCION"
   End If


End Sub

Private Sub cmd_factura_Click()
   Dim var_inserta As Boolean
   Dim var_factura As Integer
   Dim var_posible_cliente As Boolean
   rsaux11.Open "select codigo as segment1 , Cantidad from estampados_300716", cnn, adOpenDynamic, adLockOptimistic
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
         If var_clave_movimiento = "VDI" Then
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
   MsgBox "Surgio un error al generar los documentos electrónicos", vbOKOnly, "ATENCION"
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
      var_sello = Me.txt_folio_Sello
      rs.Open "select * from xxvia_tb_Devoluciones_clientes where numero = " + CStr(var_numero_folio) + " and organizacion = " + var_unidad_organizacional + " AND MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_estatus_movimiento = IIf(IsNull(rs!estatus), "", rs!estatus)
      End If
      rs.Close
      If var_clave_movimiento = "DCS" Then
         If var_estatus_movimiento = "I" Then
            var_numero_folio_devoluciones = CDbl(Me.txt_folio)
            var_clave_almacen_devolucion = Me.txt_almacen
            var_referencia_global_dev = Me.txt_referencia
            frmoracle_devoluciones_desgloce.Show 1
         Else
            rs.Open "select sum(cantidad) from xxvia_Tb_devoluciones_clientes where numero_referencia = " + Me.txt_folio_Sello, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_cantidad_leida = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
            Else
               var_cantidad_leida = 0
            End If
            rs.Close
            rs.Open "select sum(cantidad) from tb_Devoluciones where agente = '" + Me.txt_agente + "' and estatus = 'I' and numero = " + Me.txt_folio_Sello, cnn_devolucion_anes, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_cantidad_enviada = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
            Else
               var_cantidad_enviada = 0
            End If
            rs.Close
            If var_cantidad_leida <> var_cantidad_enviada Then
               var_si = MsgBox("Existe diferencia entre la cantidad leida y la cantidad enviada ¿Desea cerrar el movimiento?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  var_si = MsgBox("Confirmar el cerrado del movimiento", vbYesNo, "ATENCION")
               End If
            Else
               var_si = 6
            End If
            If var_si = 6 Then
            rs.Open "select codigo from xxvia_tb_devoluciones_clientes where numero = " + Me.txt_folio + " AND MOVIMIENTO = 'DCS'", cnnoracle_4, adOpenDynamic, adLockOptimistic
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
                  'Me.txt_cliente = 555741
                  'var_cadena = "SELECT  hcsu.attribute1 AS RFC,hcsu.order_type_id,hca.cust_account_id,hcp.site_use_id AS VCHA_CLI_CLAVE_ID,hl.address1 VCHA_CLI_NOMBRE FROM hz_parties hp,hz_party_sites hps,hz_cust_accounts hca,hz_cust_acct_sites_all hcas,hz_cust_site_uses_all hcsu, hz_locations hl,hr_operating_units hr,hz_customer_profiles hcp,ar_collectors arc,ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id=hps.party_id AND hps.party_site_id=hcas.party_site_id AND hca.cust_account_id=hcas.cust_account_id AND hcas.cust_acct_site_id=hcsu.cust_acct_site_id AND hps.location_id=hl.location_id AND hcas.org_id=hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcp.collector_id = " + Me.txt_agente + " AND hcp.site_use_id = " + Me.txt_cliente
                  'var_cadena = "select * from xxvia_vw_clientes_bcp where site_use_id = " + Me.txt_cliente + " and  account_number = '" + Me.txt_titular + "' and site_use_code = 'BILL_TO'"
                  var_cadena = "select * from xxvia_vw_clientes_bcp where site_use_id = " + Me.txt_cliente
                  If rsaux7.State = 1 Then
                     rsaux7.Close
                  End If
                  rsaux7.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux7.EOF Then
                     var_rfc = IIf(IsNull(rsaux7!rfc), "", rsaux7!rfc)
                  End If
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
                        rsaux.Open "SELECT L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id, SUM(NVL(L.GROSS_unit_selling_price,l.unit_selling_price)) AS PRECIO FROM RA_CUSTOMER_TRX_LINES_ALL L, RA_CUSTOMER_TRX_ALL E, ra_cust_trx_types_all TYPES Where TYPES.TYPE = 'INV' AND TYPES.cust_trx_type_id = E.cust_trx_type_id AND TYPES.org_id = E.org_id AND l.customer_trx_id = E.customer_trx_id AND L.inventory_item_id = " + CStr(rsaux9!inventory_item_id) + " AND E.sold_to_customer_id = " + CStr(rsaux9!TITULAR) + " and trx_date >= to_date('01/01/2016','DD/MM/YYYY') AND TRX_NUMBER NOT LIKE 'F%' GROUP BY L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id  ORDER BY trx_date DESC", cnnoracle_4, adOpenDynamic, adLockOptimistic
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
                              '… Establecer conexión a la base de datos con el objeto objConn.
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
                        rsaux.Open "SELECT L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id, SUM(NVL(L.GROSS_unit_selling_price,l.unit_selling_price)) AS PRECIO FROM RA_CUSTOMER_TRX_LINES_ALL L, RA_CUSTOMER_TRX_ALL E, ra_cust_trx_types_all TYPES Where TYPES.TYPE = 'INV' AND TYPES.cust_trx_type_id = E.cust_trx_type_id AND TYPES.org_id = E.org_id AND l.customer_trx_id = E.customer_trx_id AND L.inventory_item_id = " + CStr(rsaux9!inventory_item_id) + " AND E.sold_to_customer_id = " + CStr(rsaux9!TITULAR) + " AND TRX_NUMBER NOT LIKE 'F%'  GROUP BY L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id  ORDER BY trx_date desc", cnnoracle_4, adOpenDynamic, adLockOptimistic
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
                              '… Establecer conexión a la base de datos con el objeto objConn.
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
                                 'var_precio = 279
                                 rsaux10.Close
                                 VAR_DESCUENTO = 0
                                 rsaux11.Open "SELECT DISTINCT(list_header_id) as calificador FROM qp_qualifiers_v WHERE list_header_id IS NOT NULL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 While Not rsaux11.EOF
                                       If rsaux11!calificador <> 1144154 Then
                                       On Error GoTo desc:
                                       If rsaux11!calificador <> 1778169 Then
                                          rsaux10.Open "select xxvia_fn_descuento_titular(" + CStr(rsaux11!calificador) + ",'" + Me.txt_titular + "') as descuento from dual ", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          If Not rsaux10.EOF Then
                                             If rsaux10!DESCUENTO > VAR_DESCUENTO Then
                                                VAR_DESCUENTO = rsaux10!DESCUENTO
                                             End If
                                          
                                          End If
                                          rsaux10.Close
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
                        'VAR_PRECIO_ENTERO = 279
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
                                 
                                    rs.Open "select * from tb_Devoluciones where agente = '" + Me.txt_agente + "' and estatus = 'I' and numero = " + Me.txt_folio_Sello, cnn_devolucion_anes, adOpenDynamic, adLockOptimistic
                                    If Not rs.EOF Then
                                       var_tipo_devolucion = rs!tipo_Devolucion_2
                                    End If
                                    rs.Close
                                    If var_tipo_devolucion = "CONVENIO" Then
                                       var_motivo_devolucion = "133"
                                    Else
                                       If var_tipo_devolucion = "ERROR DE PEDIDO" Then
                                          var_motivo_devolucion = "222"
                                       Else
                                          If var_tipo_devolucion = "EXTEMPORANEO" Then
                                             var_motivo_devolucion = "16"
                                          Else
                                             var_motivo_devolucion = "309"
                                          End If
                                       End If
                                    End If
                                    
                                    rsaux10.Open "INSERT INTO XXVIA_TB_dEV_CLIENTES_DESGLOCE (NUMERO, ORGANIZACION, INVENTORY_ITEM_ID, CANTIDAD, CAUSA_DEVOLUCION, CONSECUTIVO, DESCRIPCION_CAUSA, ESTATUS, CODIGO, DESCRIPCION, LOCALIZADOR, MOVIMIENTO, TIPO_PEDIDO) VALUES (" + Me.txt_folio + "," + var_unidad_organizacional + "," + CStr(rsaux9!inventory_item_id) + "," + CStr(var_cantidad_LEER) + ",'" + var_motivo_devolucion + "'," + CStr(var_consecutivo) + ",'" + var_tipo_devolucion + "','','" + rsaux9!codigo + "','" + rsaux9!Descripcion + "','" + IIf(IsNull(rsaux9!localizador), "", rsaux9!localizador) + "','" + var_clave_movimiento + "'," + CStr(IIf(IsNull(rsaux9!tipo_pedido), 0, rsaux9!tipo_pedido)) + ",'" + Me.txt_limite_inferior + "','" + Me.txt_limite_superior + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 Else
                                    
                                    rs.Open "select * from tb_Devoluciones where agente = '" + Me.txt_agente + "' and estatus = 'I' and numero = " + Me.txt_folio_Sello, cnn_devolucion_anes, adOpenDynamic, adLockOptimistic
                                    If Not rs.EOF Then
                                       var_tipo_devolucion = rs!tipo_Devolucion_2
                                    End If
                                    rs.Close
                                    If var_tipo_devolucion = "CONVENIO" Then
                                       var_motivo_devolucion = "133"
                                    Else
                                       If var_tipo_devolucion = "ERROR DE PEDIDO" Then
                                          var_motivo_devolucion = "222"
                                       Else
                                          If var_tipo_devolucion = "EXTEMPORANEO" Then
                                             var_motivo_devolucion = "16"
                                          Else
                                             var_motivo_devolucion = "309"
                                          End If
                                       End If
                                    End If
                                    
                                    rsaux10.Open "INSERT INTO XXVIA_TB_dEV_CLIENTES_DESGLOCE (NUMERO, ORGANIZACION, INVENTORY_ITEM_ID, CANTIDAD, CAUSA_DEVOLUCION, CONSECUTIVO, DESCRIPCION_CAUSA, ESTATUS, CODIGO, DESCRIPCION, LOCALIZADOR, MOVIMIENTO, TIPO_PEDIDO, FOLIO_INFERIOR, FOLIO_SUPERIOR) VALUES (" + Me.txt_folio + "," + var_unidad_organizacional + "," + CStr(rsaux9!inventory_item_id) + "," + CStr(var_cantidad_LEER) + ",'" + var_motivo_devolucion + "'," + CStr(var_consecutivo) + ",'" + IIf(IsNull(var_tipo_devolucion), "", var_tipo_devolucion) + "','','" + rsaux9!codigo + "','" + rsaux9!Descripcion + "','" + IIf(IsNull(rsaux9!localizador), "", rsaux9!localizador) + "','" + var_clave_movimiento + "'," + CStr(IIf(IsNull(rsaux9!tipo_pedido), 0, rsaux9!tipo_pedido)) + ",'" + Me.txt_limite_inferior + "','" + Me.txt_limite_superior + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
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
                                    rs.Open "select * from tb_Devoluciones where agente = '" + Me.txt_agente + "' and estatus = 'I' and numero = " + Me.txt_folio_Sello, cnn_devolucion_anes, adOpenDynamic, adLockOptimistic
                                    If Not rs.EOF Then
                                       var_tipo_devolucion = rs!tipo_Devolucion_2
                                    End If
                                    rs.Close
                                    If var_tipo_devolucion = "CONVENIO" Then
                                       var_motivo_devolucion = "133"
                                    Else
                                       If var_tipo_devolucion = "ERROR DE PEDIDO" Then
                                          var_motivo_devolucion = "222"
                                       Else
                                          If var_tipo_devolucion = "EXTEMPORANEO" Then
                                             var_motivo_devolucion = "16"
                                          Else
                                             var_motivo_devolucion = "309"
                                          End If
                                       End If
                                    End If
                             
                             If var_devolucion_costales = 1 Then
                                rsaux10.Open "INSERT INTO XXVIA_TB_dEV_CLIENTES_DESGLOCE (NUMERO, ORGANIZACION, INVENTORY_ITEM_ID, CANTIDAD, CAUSA_DEVOLUCION, CONSECUTIVO, DESCRIPCION_CAUSA, ESTATUS, CODIGO, DESCRIPCION, LOCALIZADOR, MOVIMIENTO, TIPO_PEDIDO) VALUES (" + Me.txt_folio + "," + var_unidad_organizacional + "," + CStr(rsaux9!inventory_item_id) + "," + CStr(var_cantidad_LEER) + ",'" + var_motivo_devolucion + "'," + CStr(var_consecutivo) + ",'" + var_tipo_devolucion + "','','" + rsaux9!codigo + "','" + rsaux9!Descripcion + "','" + IIf(IsNull(rsaux9!localizador), "", rsaux9!localizador) + "','" + var_clave_movimiento + "'," + CStr(IIf(IsNull(rsaux9!tipo_pedido), 0, rsaux9!tipo_pedido)) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                             Else
                                rsaux10.Open "INSERT INTO XXVIA_TB_dEV_CLIENTES_DESGLOCE (NUMERO, ORGANIZACION, INVENTORY_ITEM_ID, CANTIDAD, CAUSA_DEVOLUCION, CONSECUTIVO, DESCRIPCION_CAUSA, ESTATUS, CODIGO, DESCRIPCION, LOCALIZADOR, MOVIMIENTO, TIPO_PEDIDO) VALUES (" + Me.txt_folio + "," + var_unidad_organizacional + "," + CStr(rsaux9!inventory_item_id) + "," + CStr(var_cantidad_LEER) + ",'" + var_motivo_devolucion + "'," + CStr(var_consecutivo) + ",'" + var_tipo_devolucion + "','','" + rsaux9!codigo + "','" + rsaux9!Descripcion + "','" + IIf(IsNull(rsaux9!localizador), "", rsaux9!localizador) + "','" + var_clave_movimiento + "'," + CStr(IIf(IsNull(rsaux9!tipo_pedido), 0, rsaux9!tipo_pedido)) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
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
               MsgBox "Los siguientes códigos no son retornables " + var_cadena_codigos_retornables, vbOKOnly, "ATENCION"
            End If
            End If
         End If
      Else
         If var_clave_movimiento = "VDI" Then
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
                  frmoracle_tipo_pedido.Show 1
               End If
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
                           If var_cadena_posible_existencias = "" Then
                              
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
                                 var_cadena = "INSERT INTO oe_headers_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref, creation_date, created_by, last_update_date, last_updated_by, operation_code , sold_to_org_id        , SHIP_TO_ORG_id                   ,INVOICE_TO_ORG_ID     , Order_type_ID, PRICE_LIST, org_id, ship_from_org_id)"
                                 var_cadena = var_cadena + "  VALUES (1001,'SIDVDI_" + Trim(CStr((var_numero_folio_devoluciones))) + "',SYSDATE,-1,SYSDATE, -1,'INSERT', " + CStr(rs!TITULAR) + "," + CStr(rs!establecimiento) + "," + CStr(rs!Cliente) + "," + CStr(var_clave_tipo_pedido) + ",'" + var_lista_precios + "'," + var_empresa + "," + var_unidad_organizacional + ")"
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
                                       var_cadena = var_cadena + " VALUES (1001,'SIDVDI_" + Trim(CStr(var_numero_folio_devoluciones)) + "','" + CStr(var_i) + "', " + CStr(rs!inventory_item_id) + ", " + CStr(rs!cantidad) + ",'INSERT', -1,SYSDATE, -1,SYSDATE," + CStr(rs!Precio) + "," + CStr(rs!Precio) + ",'Y', " + CStr(rs!cantidad) + ", '" + VAR_MEDIDA + "','" + IIf(IsNull(rs!localizador), "", rs!localizador) + "','" + Me.txt_almacen + "'," + var_empresa + "," + var_unidad_organizacional + ")"
                                    End If
                                    rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    rs.MoveNext
                              Wend
                              On Error GoTo salir2
                              If Me.txt_almacen = "TEX_VB1" Or Me.txt_almacen = "TEX_VB2" Or Me.txt_almacen = "TEX_VB4" Then
                                 rsaux.Open "INSERT INTO oe_actions_iface_all (order_source_ID, orig_sys_document_ref, operation_code) VALUES (1001, 'SIDVDI_" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "','BOOK_ORDER')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              Else
                                 rsaux.Open "INSERT INTO oe_actions_iface_all (order_source_ID, orig_sys_document_ref, operation_code) VALUES (1001, 'SIDVDI_" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "','BOOK_ORDER')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              'var_numero_folio_devoluciones = CDbl(Me.txt_folio)
                              rsaux.Open "CALL XXVIA_PK_INTERFACES_OM.importar_pedido('SIDVDI_" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "'," + var_empresa + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If rsaux.State = 1 Then
                                 rsaux.Close
                              End If
                              rsaux.Open "UPDATE XXVIA_TB_DEVOLUCIONES_CLIENTES A SET ESTATUS = 'I' WHERE A.NUMERO = " + CStr(var_numero_folio_devoluciones) + " AND A.ORGANIZACION = " + var_unidad_organizacional + "  AND A.MOVIMIENTO = '" + var_clave_movimiento + "'"
                              rsaux.Open "select order_number from oe_order_headers_all where orig_sys_document_ref = 'SIDVDI_" + Trim(CStr(var_numero_folio_devoluciones)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              var_pedido = rsaux(0).Value
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
                                          '… Establecer conexión a la base de datos con el objeto objConn.
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
                     rs.Close
                  Else
                     MsgBox "No se a indicado una lista de precios", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "No se a indicado un tipo de pedido", vbOKOnly, "ATENCION"
               End If
            End If
         End If
         ''''
      End If
   Else
      MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
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
   MsgBox "No se pudo generar el documento electrónico", vbOKOnly, "ATENCION"
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
End Sub


Private Sub cmd_movimiento_masivo_Click()
   Dim var_inserta As Boolean
   Dim var_factura As Integer
   Dim var_posible_cliente As Boolean
    
    rsaux12.Open "select * from muebles_cantia_301213", cnn, adOpenDynamic, adLockOptimistic
    var_primera_vez = True
    While Not rsaux12.EOF
          Me.txt_codigo = rsaux12!codigo
          var_cantidad_leida = rsaux12!cantidad
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
                         'txt_codigo.SetFocus
                         'txt_codigo = ""
                      End If
                                    
'----------------
                   Else
                      frmmensaje.lbl_mensaje = "El artículo no tiene precio " + Me.txt_codigo
                      frmmensaje.Show
                      Me.txt_codigo = ""
                   End If
                Else
                   frmmensaje.lbl_mensaje = "El artículo no se encuentra en la lista de precios del cliente " + Me.txt_codigo
                   frmmensaje.Show
                   Me.txt_codigo = ""
                End If
                rsaux10.Close
             Else
                frmmensaje.lbl_mensaje = "Error en código " + Me.txt_codigo
                frmmensaje.Show
                txt_codigo = ""
             End If
             rsaux8.Close
          End If
          rsaux12.MoveNext
    Wend
    rsaux12.Close
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_sello = ""
   Me.txt_folio_Sello.Enabled = True
   Me.txt_folio_Sello = ""
   var_devolucion_costales = 0
   lbl_total = "0"
   'Me.lbl_cantidad_enviada = "0"
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
   txt_almacen.SetFocus
   txt_nombre_almacen = ""
   txt_nombre_agente = ""
   txt_nombre_establecimiento = ""
   txt_nombre_cliente = ""
   
   
   
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

Private Sub Command3_Click()
   If Me.lv_entradas.ListItems.Count > 0 Then
      For var_xy = 1 To Me.lv_entradas.ListItems.Count
          Me.lv_entradas.ListItems.Item(var_xy).Selected = True
          var_renglon = var_xy
          Me.txt_codigo = Me.lv_entradas.selectedItem
          If Me.txt_codigo <> "" Then
             Me.lv_entradas.ListItems.Item(var_renglon).Selected = True
             var_Cantidad_posible = CDbl(Me.lv_entradas.selectedItem.SubItems(2))
             var_cantidad_leida = var_Cantidad_posible
             var_cantidad_total = CDbl(Me.lv_entradas.selectedItem.SubItems(3))
             rs.Open "SELECT TO_CHAR(SYSDATE,'dd/mm/yyyy hh24:mi:ss') AS FECHA FROM DUAL", cnnoracle_4, adOpenDynamic, adLockOptimistic
             VAR_FECHA_HORA = rs(0).Value
             rs.Close
             If var_cantidad_total + var_cantidad_leida <= var_Cantidad_posible Then
                If Trim(txt_codigo.Text) <> "" Then
                   var_posible_existencia = 1
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
                         var_clave_titular = Me.txt_clave_titular
               
                         rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                         If Not rsaux8.EOF Then
                            var_unidad_medida = rsaux8!PRIMARY_UOM_CODE
                            var_descripcion_articulo = rsaux8!Description
                            var_inventory_item_id = rsaux8!inventory_item_id
                            var_salida_masiva = IIf(IsNull(rsaux8!attribute10), "N", rsaux8!attribute10)
                            '9019
                            var_cadena = "select * from  qp_secu_list_headers_v a, qp_list_lines_v b Where a.list_header_id = b.list_header_id and  B.product_attr_value = " + CStr(var_inventory_item_id) + " AND a.list_header_id = " + CStr(var_clave_lista_precios) + " and   product_attr_val_disp = '" + Me.txt_codigo + "'"
                            'MsgBox var_cadena
                            rsaux10.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                            If Not rsaux10.EOF Then
                               If IIf(IsNull(rsaux10!OPERAND), 0, rsaux10!OPERAND) > 0 Then

                              Else
                                 txt_codigo = ""
                                 frmmensaje.lbl_mensaje = "El artículo no tiene precio"
                                 frmmensaje.Show
                              End If
                           Else
                              txt_codigo = ""
                              frmmensaje.lbl_mensaje = "El artículo no se encuentra en la lista de precios del cliente"
                              frmmensaje.Show
                           End If
                           rsaux10.Close
                        Else
                           txt_codigo = ""
                           frmmensaje.lbl_mensaje = "Error en código"
                           frmmensaje.Show
                        End If
                        rsaux8.Close
               
               
                         var_cadena = "insert into xxvia_tb_devoluciones_clientes (numero, organizacion, inventory_item_id, codigo, cantidad, descripcion, estatus, agente, cliente, establecimiento, titular, nombre_agente, almacen, nombre_almacen, nombre_cliente, nombre_establecimiento, referencia, usuario, maquina, fecha_inicio, unidad_medida, precio, localizador, movimiento,tipo_pedido, numero_referencia, FOLIO_INFERIOR, FOLIO_SUPERIOR)"
                         var_cadena = var_cadena + " values (" + CStr(var_numero_folio) + "," + var_unidad_organizacional + "," + CStr(var_inventory_item_id) + ",'" + Me.txt_codigo + "'," + CStr(var_cantidad_leida) + ",'" + Replace(var_descripcion_articulo, "'", " ") + "',''," + Me.txt_agente + "," + Me.txt_cliente + "," + Me.txt_establecimiento + "," + CStr(var_clave_titular) + ",'" + Me.txt_nombre_agente + "','" + Me.txt_almacen + "','" + Me.txt_nombre_almacen + "','" + Replace(Me.txt_nombre_cliente, "'", " ") + "','" + Replace(Me.txt_nombre_establecimiento, "'", " ") + "','" + Me.txt_referencia + "','" + var_clave_usuario_global + "', '" + fun_NombrePc + "','" + CStr(Date) + "','" + var_unidad_medida + "',0,'" + var_localizador_subinventario + "','" + var_clave_movimiento + "'," + CStr(var_tipo_pedido) + "," + Me.txt_folio_Sello + ",'" + Me.txt_limite_inferior + "','" + Me.txt_limite_superior + "')"
                         'MsgBox var_cadena
                         'Me.txt_folio_Sello = var_cadena
                         rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               
                         valor = Me.txt_codigo
                         var_j = 1
                         For var_j = 1 To Me.lv_entradas.ListItems.Count
                             Me.lv_entradas.ListItems.Item(var_j).Selected = True
                     
                             If Me.lv_entradas.selectedItem = Me.txt_codigo Then
                                Me.lv_entradas.selectedItem.SubItems(3) = CDbl(Me.lv_entradas.selectedItem.SubItems(3)) + var_cantidad_leida
                                var_renglon = var_j
                             End If
                         Next var_j
                         ''var_renglon = lv_entradas.ListItems.Count
                         'On Error GoTo x:
                         'clnt.MSSoapInit var_webservice_texto
                         'var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MOVIMIENTO: " + var_clave_movimiento + Me.txt_folio + "  MAQUINA: " + fun_NombrePc + Chr(13) + " USUARIO: " + var_nombre_usuario_global + Chr(13) + "FECHA Y HORA: " + VAR_FECHA_HORA + Chr(13) + "Código: " + Me.txt_codigo + " " + var_descripcion_articulo + " Cantidad: " + CStr(var_cantidad_leida))
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
                             If Me.lv_entradas.selectedItem = Me.txt_codigo And Trim(Me.lv_entradas.selectedItem.SubItems(4)) = Trim(var_localizador_subinventario) Then
                                Me.lv_entradas.selectedItem.SubItems(3) = CDbl(Me.lv_entradas.selectedItem.SubItems(3)) + var_cantidad_leida
                                var_renglon = var_j
                             End If
                         Next var_j
                         'On Error GoTo Z:
                         'clnt.MSSoapInit var_webservice_texto
                         'var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MOVIMIENTO: " + var_clave_movimiento + Me.txt_folio + "  MAQUINA: " + fun_NombrePc + Chr(13) + " USUARIO: " + var_nombre_usuario_global + Chr(13) + "FECHA Y HORA: " + VAR_FECHA_HORA + Chr(13) + "Código: " + Me.txt_codigo + " " + var_descripcion_articulo + " Cantidad: " + CStr(var_cantidad_leida))
                         'Set clnt = Nothing
Z:
                         Call ilumina_grid
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
             Else
                frmmensaje.lbl_articulo = Me.txt_codigo + " " + var_descripcion_articulo
                frmmensaje.lbl_mensaje = "La cantidad supera a la posible en la devolución"
                frmmensaje.Show
                Me.txt_codigo = ""
             End If
          End If
      Next var_xy
   End If
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
   parametros(0) = "admcdindustrial"
   parametros(1) = "SIDAlmacenbkp"
   If cnn_devolucion_anes.State = 1 Then
      cnn_devolucion_anes.Close
   End If
   cnn_devolucion_anes.Open "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=devolucion_anes;Data Source=SQLQUEZADA2"
   Me.chk_factura = 0
   Me.chk_factura.Visible = False
   Me.txt_movimiento = var_clave_movimiento
   Me.txt_movimiento.Visible = False
   Me.frm_pasar_todo.Visible = False
   If var_clave_usuario_global = "11" Or var_clave_usuario_global = "8" Then
      Me.cmd_pasar_todo.Visible = False
   Else
      If var_unidad_organizacional = "93" Then
         Me.cmd_pasar_todo.Visible = False
      End If
   End If
   lbl_total = "0"
   lbl_cancelado = ""
   var_año = 2005
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
   Me.txt_titular.Enabled = False
   Me.txt_nombre_titular.Enabled = False
   Me.txt_cliente.Enabled = False
   Me.txt_nombre_cliente.Enabled = False
   Me.txt_establecimiento.Enabled = False
   Me.txt_nombre_establecimiento.Enabled = False
   Me.txt_referencia.Enabled = False
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   'Me.cmd_pasar_todo.Visible = False
   If var_clave_movimiento = "DCS" Then
      Me.Command3.Visible = True
   Else
      Command3.Visible = False
   End If

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
         MsgBox "Imporsible modificar el movimiento", vbOKOnly, "ATENCION"
      Else
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
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct agente vcha_age_agente_id, nombre_agente vcha_age_nombre from tb_Devoluciones where almacen = 'ANE' order by NOMBRE_AGENTE ", cnn_devolucion_anes, adOpenDynamic, adLockOptimistic
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
      Me.txt_folio_Sello.SetFocus
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
            Me.txt_folio_Sello.Enabled = True
            Me.txt_folio_Sello.SetFocus
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
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_unidad_organizacional = "93" Then
         rs.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " AND secondary_inventory_name in ('CDI_ALMCAL','CDI_ALMPT')", cnnoracle_4, adOpenDynamic, adLockOptimistic
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
          If var_tipo_busqueda = 1 Then
            rs.Open "select * from xxvia_tb_devoluciones_clientes where REFERENCIA = '" + Me.txt_busqueda_folio + "' and organizacion = " + var_unidad_organizacional + " AND MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         Else
            rs.Open "select * from xxvia_tb_devoluciones_clientes where numero = " + Me.txt_busqueda_folio + " and organizacion = " + var_unidad_organizacional + " AND MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
         If Not rs.EOF Then
            var_clave_titular = rs!TITULAR
            var_tipo_pedido = IIf(IsNull(rs!tipo_pedido), 0, rs!tipo_pedido)
            var_estatus_movimiento = IIf(IsNull(rs!estatus), "", rs!estatus)
            var_almacen_Destino = IIf(IsNull(rs!ALMACEN), "", rs!ALMACEN)
            Me.txt_limite_inferior = IIf(IsNull(rs!FOLIO_INFERIOR), 0, rs!FOLIO_INFERIOR)
            Me.txt_limite_superior = IIf(IsNull(rs!FOLIO_SUPERIOR), 0, rs!FOLIO_SUPERIOR)
            
               If Len(Me.txt_limite_inferior) = 6 Then
                  Me.txt_limite_inferior = "00" + Me.txt_limite_inferior
               End If
               If Len(Me.txt_limite_inferior) = 7 Then
                  Me.txt_limite_inferior = "0" + Me.txt_limite_inferior
               End If
               
               If Len(Me.txt_limite_superior) = 6 Then
                  Me.txt_limite_superior = "00" + Me.txt_limite_superior
               End If
               If Len(Me.txt_limite_superior) = 7 Then
                  Me.txt_limite_superior = "0" + Me.txt_limite_superior
               End If
            
            
            
            rsaux8.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre, nvl(locator_type,0) as localizador  from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " and secondary_inventory_name = '" + rs!ALMACEN + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
               var_localizador = rsaux8!localizador
            End If
            rsaux8.Close
            txt_almacen = var_almacen_Destino
            txt_referencia = IIf(IsNull(rs!Referencia), "", rs!Referencia)
            txt_cliente = rs!Cliente
            txt_nombre_cliente = rs!nombre_cliente
            txt_agente = rs!Agente
            txt_nombre_agente = rs!NOMBRE_AGENTE
            txt_establecimiento = rs!establecimiento
            txt_nombre_establecimiento = rs!nombre_Establecimiento
            Me.txt_folio_Sello = rs!numero_referencia
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
               var_clave_lista_precios = 9016
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
            cnn_devolucion_anes.CommandTimeout = 360
            rsaux.Open "select * from tb_Devoluciones where agente = '" + Me.txt_agente + "' and estatus = 'I' and numero = " + Me.txt_folio_Sello, cnn_devolucion_anes, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               strconsulta = "SELECT ACCOUNT_NUMBER TITULAR, ACCOUNT_FULL_NAME NOMBRE_TITULAR FROM XXVIA_VW_CLIENTES_BCP WHERE CUST_ACCOUNT_ID = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, rsaux!TITULAR)
                    .Parameters.Append parametro
               End With
               Set rsaux8 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               Me.txt_titular = rsaux8!TITULAR
               Me.txt_nombre_titular = rsaux8!NOMBRE_TITULAR
               Me.txt_clave_titular = rsaux!TITULAR
               rsaux8.Close
               Me.txt_cliente = rsaux!Cliente
               Me.txt_nombre_cliente = rsaux!nombre_cliente
               var_cadena = "SELECT  hcsu.price_list_id, hcsu.order_type_id,hca.cust_account_id, hcp.site_use_id AS VCHA_CLI_CLAVE_ID,hl.address1 VCHA_CLI_NOMBRE FROM hz_parties hp,hz_party_sites hps,hz_cust_accounts hca,hz_cust_acct_sites_all hcas,hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND  hcp.site_use_id = " + Me.txt_cliente
               rsaux10.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux10.EOF Then
                  var_clave_lista_precios = rsaux10!price_list_id
               Else
                  var_clave_lista_precios = 0
               End If
               rsaux10.Close
               Me.txt_establecimiento = rsaux!establecimiento
               Me.txt_nombre_establecimiento = rsaux!nombre_Establecimiento
               Me.txt_referencia = rsaux!Referencia
               While Not rsaux.EOF
                     Set list_item = Me.lv_entradas.ListItems.Add(, , rsaux!codigo)
                     list_item.SubItems(1) = IIf(IsNull(rsaux!Descripcion), "", rsaux!Descripcion)
                     list_item.SubItems(2) = IIf(IsNull(rsaux!cantidad), "", rsaux!cantidad)
                     list_item.SubItems(3) = 0
                     rsaux.MoveNext
               Wend
               Me.txt_titular.Enabled = False
               Me.txt_nombre_titular.Enabled = False
               Me.txt_cliente.Enabled = False
               Me.txt_nombre_cliente.Enabled = False
               Me.txt_establecimiento.Enabled = False
               Me.txt_nombre_establecimiento.Enabled = False
               Me.txt_referencia.Enabled = False
            
               Me.txt_codigo.Enabled = True
               Me.txt_codigo.SetFocus
            
            Else
               MsgBox "El número de devolución no existe o no pertenece al agente seleccionado", vbOKOnly, "ATENCION"
            End If
            rsaux.Close
            
            
            Me.lbl_total = 0
            While Not rs.EOF
                  For var_j = 1 To Me.lv_entradas.ListItems.Count
                      Me.lv_entradas.ListItems.Item(var_j).Selected = True
                      If Me.lv_entradas.selectedItem = rs!codigo Then
                         Me.lv_entradas.selectedItem.SubItems(3) = CDbl(Me.lv_entradas.selectedItem.SubItems(3)) + rs!cantidad
                         Me.lbl_total = CDbl(Me.lbl_total) + rs!cantidad
                      End If
                  Next var_j
                  rs.MoveNext
            Wend
            Me.lbl_total = Format(CDbl(Me.lbl_total), "###,###0")
            Me.txt_folio_Sello.Enabled = False
            'If lv_entradas.ListItems.Count > 11 Then
            '   lv_entradas.ColumnHeaders(2).Width = 5050.22
            'Else
            '   lv_entradas.ColumnHeaders(2).Width = 5300.22
            'End If
            'rs.MoveFirst
            rs.MoveFirst
            If IIf(IsNull(rs!estatus), "", rs!estatus) = "" Then
               Me.txt_codigo.Enabled = True
               Me.txt_foco.Enabled = True
            Else
               Me.txt_codigo.Enabled = False
               Me.txt_foco.Enabled = False
            End If
         Else
            If var_tipo_busqueda = 1 Then
               MsgBox "El sello de la devolución no existe", vbOKOnly, "ATENCION"
            Else
               MsgBox "El número de movimiento no existe ", vbOKOnly, "ATENCION"
            End If
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
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If IsNumeric(txt_cantidad_eliminar) Then
         VAR_CANTIDAD_ELIMINAR = Val(txt_cantidad_eliminar)
         var_posible_eliminar = True
         If VAR_CANTIDAD_ELIMINAR <= lv_entradas.selectedItem.SubItems(3) * 1 = True Then
            If rsaux8.State = 1 Then
               rsaux8.Close
            End If
            rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + Me.lv_entradas.selectedItem + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
               var_inventory_item_id = rsaux8!inventory_item_id
               var_localizador_subinventario = Me.lv_entradas.selectedItem.SubItems(4)
               rs.Open "UPDATE xxvia_tb_devoluciones_clientes SET CANTIDAD = CANTIDAD -" + CStr(VAR_CANTIDAD_ELIMINAR) + " where numero = " + Str(var_numero_folio) + " and codigo = '" + Me.lv_entradas.selectedItem + "' and inventory_item_id = " + CStr(var_inventory_item_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
               lv_entradas.selectedItem.SubItems(3) = lv_entradas.selectedItem.SubItems(3) - Val(txt_cantidad_eliminar)
               lbl_total = CStr(CDbl(lbl_total) - Val(txt_cantidad_eliminar))
               var_renglon = lv_entradas.selectedItem.Index
               On Error GoTo x:
               rs.Open "SELECT TO_CHAR(SYSDATE,'dd/mm/yyyy hh24:mi:ss') AS FECHA FROM DUAL", cnnoracle_4, adOpenDynamic, adLockOptimistic
               VAR_FECHA_HORA = rs(0).Value
               rs.Close
               clnt.MSSoapInit var_webservice_texto
               'var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MOVIMIENTO: " + var_clave_movimiento + Me.txt_folio + "  MAQUINA: " + fun_NombrePc + Chr(13) + " USUARIO: " + var_nombre_usuario_global + Chr(13) + "FECHA Y HORA: " + VAR_FECHA_HORA + Chr(13) + "Código: " + Me.txt_codigo + " " + var_descripcion_articulo + " Cantidad: " + CStr(var_cantidad_leida))
               var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MOVIMIENTO: " + var_clave_movimiento + Me.txt_folio + "  MAQUINA: " + fun_NombrePc + Chr(13) + " USUARIO: " + var_nombre_usuario_global + Chr(13) + "FECHA Y HORA: " + VAR_FECHA_HORA + Chr(13) + "Código: " + Me.lv_entradas.selectedItem + " " + Me.lv_entradas.selectedItem.SubItems(1) + " Cantidad eliminada: " + CStr(Val(txt_cantidad_eliminar)))
               Set clnt = Nothing
x:
               Call ilumina_grid
             End If
             rsaux8.Close
         Else
            MsgBox "La cantidad a eliminar supera a la cantidad asignada a la causa de devolución seleccionada", vbOKOnly, "ATENCION"
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
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
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
               'var_cadena = "select cust_acct_site_id as VCHA_CLI_CLAVE_ID, razon_social_cliente as VCHA_CLI_NOMBRE, party_site_number as party_site_number from xxvia_vw_clientes_bcp where account_number = '" + Me.txt_titular + "' and site_use_code = 'BILL_TO'"
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
               var_clave_lista_precios = 9016
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
   txt_codigo = Trim(txt_codigo)
   If KeyAscii = 13 Then
      var_localizador_subinventario = " "
      var_encontro = 0
      var_cantidad_leida = 1
      var_cantidad_leida_seg_nivel = 1
      
      var_cantidad_leida_caja = 0
      Dim var_tela As String
      var_tela = ""
      For var_j = 1 To Len(Me.txt_codigo)
          If Mid(Me.txt_codigo, var_j, 1) = "-" Then
             var_tela = var_tela + Mid(Me.txt_codigo, var_j, 1)
          End If
      Next var_j
      If Mid(Me.txt_codigo, 1, 2) = "CA" Or var_tela = "---" Then
         rs.Open "SELECT * FROM XXVIA_TB_CAJAS_PROD WHERE vcha_caj_caja_id = '" + UCase(Me.txt_codigo) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + IIf(IsNull(rs!VCHA_ART_ARTICULO_ID), "", rs!VCHA_ART_ARTICULO_ID) + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
               Me.txt_codigo = IIf(IsNull(rs!VCHA_ART_ARTICULO_ID), "", rs!VCHA_ART_ARTICULO_ID)
               var_salida_masiva = IIf(IsNull(rsaux8!attribute10), "N", rsaux8!attribute10)
               var_cantidad_leida_caja = rs!numb_caj_cantidad
            End If
         End If
         rs.Close
      End If
      
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
      
      
      var_si = 0
      var_renglon = 0
      For var_j = 1 To Me.lv_entradas.ListItems.Count
          Me.lv_entradas.ListItems.Item(var_j).Selected = True
          If Me.lv_entradas.selectedItem = Me.txt_codigo Then
             var_si = 1
             var_renglon = var_j
          End If
      Next var_j
      
      
      
      
      If var_si = 1 Then
         If Trim(Me.txt_codigo) <> "" Then
            rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
            
            
               
               var_unidad_medida = rsaux8!PRIMARY_UOM_CODE
               var_descripcion_articulo = rsaux8!Description
               var_inventory_item_id = rsaux8!inventory_item_id
               var_salida_masiva = IIf(IsNull(rsaux8!attribute10), "N", rsaux8!attribute10)
               '9019
               var_cadena = "select * from  qp_secu_list_headers_v a, qp_list_lines_v b Where a.list_header_id = b.list_header_id and  B.product_attr_value = " + CStr(var_inventory_item_id) + " AND a.list_header_id = " + CStr(var_clave_lista_precios) + " and   product_attr_val_disp = '" + Me.txt_codigo + "'"
               'MsgBox var_cadena
               rsaux10.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux10.EOF Then
                  If IIf(IsNull(rsaux10!OPERAND), 0, rsaux10!OPERAND) > 0 Then
                     If var_unidad_organizacional = "90" Then
                        var_salida_masiva = "Y"
                     End If
                     'If var_unidad_organizacional = "85" Or var_unidad_organizacional = "94" Then
                     '   var_salida_masiva = "Y"
                     'End If
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
                     frmmensaje.lbl_mensaje = "El artículo no tiene precio"
                     frmmensaje.Show
                  End If
               Else
                  txt_codigo = ""
                  frmmensaje.lbl_mensaje = "El artículo no se encuentra en la lista de precios del cliente"
                  frmmensaje.Show
               End If
               rsaux10.Close
            Else
               txt_codigo = ""
               frmmensaje.lbl_mensaje = "Error en código"
               frmmensaje.Show
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
      Else
         txt_codigo = ""
         frmmensaje.lbl_mensaje = "El artículo no viene en la devolución"
         frmmensaje.Show
      End If
   End If
End Sub

Private Sub txt_establecimiento_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
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
   If Me.txt_codigo <> "" Then
   Me.lv_entradas.ListItems.Item(var_renglon).Selected = True
   var_Cantidad_posible = CDbl(Me.lv_entradas.selectedItem.SubItems(2))
   var_cantidad_total = CDbl(Me.lv_entradas.selectedItem.SubItems(3))
   rs.Open "SELECT TO_CHAR(SYSDATE,'dd/mm/yyyy hh24:mi:ss') AS FECHA FROM DUAL", cnnoracle_4, adOpenDynamic, adLockOptimistic
   VAR_FECHA_HORA = rs(0).Value
   rs.Close
   If var_cantidad_total + var_cantidad_leida <= var_Cantidad_posible Then
      If Trim(txt_codigo.Text) <> "" Then
         var_posible_existencia = 1
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
               var_clave_titular = Me.txt_clave_titular
               
               var_cadena = "insert into xxvia_tb_devoluciones_clientes (numero, organizacion, inventory_item_id, codigo, cantidad, descripcion, estatus, agente, cliente, establecimiento, titular, nombre_agente, almacen, nombre_almacen, nombre_cliente, nombre_establecimiento, referencia, usuario, maquina, fecha_inicio, unidad_medida, precio, localizador, movimiento,tipo_pedido, numero_referencia, FOLIO_INFERIOR, FOLIO_SUPERIOR)"
               var_cadena = var_cadena + " values (" + CStr(var_numero_folio) + "," + var_unidad_organizacional + "," + CStr(var_inventory_item_id) + ",'" + Me.txt_codigo + "'," + CStr(var_cantidad_leida) + ",'" + Replace(var_descripcion_articulo, "'", " ") + "',''," + Me.txt_agente + "," + Me.txt_cliente + "," + Me.txt_establecimiento + "," + CStr(var_clave_titular) + ",'" + Me.txt_nombre_agente + "','" + Me.txt_almacen + "','" + Me.txt_nombre_almacen + "','" + Replace(Me.txt_nombre_cliente, "'", " ") + "','" + Replace(Me.txt_nombre_establecimiento, "'", " ") + "','" + Me.txt_referencia + "','" + var_clave_usuario_global + "', '" + fun_NombrePc + "','" + CStr(Date) + "','" + var_unidad_medida + "',0,'" + var_localizador_subinventario + "','" + var_clave_movimiento + "'," + CStr(var_tipo_pedido) + "," + Me.txt_folio_Sello + ",'" + Me.txt_limite_inferior + "','" + Me.txt_limite_superior + "')"
               'MsgBox var_cadena
               'Me.txt_folio_Sello = var_cadena
               rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               
               valor = Me.txt_codigo
               var_j = 1
               For var_j = 1 To Me.lv_entradas.ListItems.Count
                     Me.lv_entradas.ListItems.Item(var_j).Selected = True
                     
                     If Me.lv_entradas.selectedItem = Me.txt_codigo Then
                        Me.lv_entradas.selectedItem.SubItems(3) = CDbl(Me.lv_entradas.selectedItem.SubItems(3)) + var_cantidad_leida
                        var_renglon = var_j
                     End If
               Next var_j
               'var_renglon = lv_entradas.ListItems.Count
               On Error GoTo x:
               clnt.MSSoapInit var_webservice_texto
               var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MOVIMIENTO: " + var_clave_movimiento + Me.txt_folio + "  MAQUINA: " + fun_NombrePc + Chr(13) + " USUARIO: " + var_nombre_usuario_global + Chr(13) + "FECHA Y HORA: " + VAR_FECHA_HORA + Chr(13) + "Código: " + Me.txt_codigo + " " + var_descripcion_articulo + " Cantidad: " + CStr(var_cantidad_leida))
               Set clnt = Nothing
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
                     If Me.lv_entradas.selectedItem = Me.txt_codigo And Trim(Me.lv_entradas.selectedItem.SubItems(4)) = Trim(var_localizador_subinventario) Then
                        Me.lv_entradas.selectedItem.SubItems(3) = CDbl(Me.lv_entradas.selectedItem.SubItems(3)) + var_cantidad_leida
                        var_renglon = var_j
                     End If
               Next var_j
               On Error GoTo Z:
               clnt.MSSoapInit var_webservice_texto
               var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MOVIMIENTO: " + var_clave_movimiento + Me.txt_folio + "  MAQUINA: " + fun_NombrePc + Chr(13) + " USUARIO: " + var_nombre_usuario_global + Chr(13) + "FECHA Y HORA: " + VAR_FECHA_HORA + Chr(13) + "Código: " + Me.txt_codigo + " " + var_descripcion_articulo + " Cantidad: " + CStr(var_cantidad_leida))
               Set clnt = Nothing
Z:
               Call ilumina_grid
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
   Else
      frmmensaje.lbl_articulo = Me.txt_codigo + " " + var_descripcion_articulo
      frmmensaje.lbl_mensaje = "La cantidad supera a la posible en la devolución"
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

Private Sub txt_folio_Sello_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_folio_Sello) Then
         strconsulta = "select * from xxvia_Tb_Devoluciones_clientes where numero_referencia = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, Me.txt_folio_Sello)
              .Parameters.Append parametro
         End With
         Set rsaux8 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         var_folio_Devolucion = 0
         If Not rsaux8.EOF Then
            var_folio_Devolucion = rsaux8!numero
         End If
         rsaux8.Close
         If var_folio_Devolucion = 0 Then
            'rs.Open "select * from tb_Devoluciones where agente = '" + Me.txt_agente + "' and estatus = 'I' and numero = " + Me.txt_folio_Sello, cnn_devolucion_anes, adOpenDynamic, adLockOptimistic
            rs.Open "select * from tb_Devoluciones where estatus = 'I' and numero = " + Me.txt_folio_Sello, cnn_devolucion_anes, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux1.Open "select distinct agente vcha_age_agente_id, nombre_agente vcha_age_nombre from tb_Devoluciones where almacen = 'ANE' and agente = '" + CStr(rs!Agente) + "' order by NOMBRE_AGENTE ", cnn_devolucion_anes, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  Me.txt_agente = rsaux1!VCHA_AGE_AGENTE_ID
                  Me.txt_nombre_agente = rsaux1!VCHA_AGE_NOMBRE
               Else
                  Me.txt_agente = ""
                  Me.txt_nombre_agente = ""
               End If
               rsaux1.Close
               Me.txt_limite_inferior = 0
               Me.txt_limite_superior = 0
               strconsulta = "SELECT ACCOUNT_NUMBER TITULAR, ACCOUNT_FULL_NAME NOMBRE_TITULAR FROM XXVIA_VW_CLIENTES_BCP WHERE CUST_ACCOUNT_ID = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, rs!TITULAR)
                    .Parameters.Append parametro
               End With
               Set rsaux8 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               Me.txt_titular = rsaux8!TITULAR
               Me.txt_nombre_titular = rsaux8!NOMBRE_TITULAR
               Me.txt_clave_titular = rs!TITULAR
               rsaux8.Close
               Me.txt_cliente = rs!Cliente
               Me.txt_nombre_cliente = rs!nombre_cliente
               var_cadena = "SELECT  hcsu.price_list_id, hcsu.order_type_id,hca.cust_account_id, hcp.site_use_id AS VCHA_CLI_CLAVE_ID,hl.address1 VCHA_CLI_NOMBRE FROM hz_parties hp,hz_party_sites hps,hz_cust_accounts hca,hz_cust_acct_sites_all hcas,hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND  hcp.site_use_id = " + Me.txt_cliente
               rsaux10.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux10.EOF Then
                  var_clave_lista_precios = rsaux10!price_list_id
               Else
                  var_clave_lista_precios = 0
               End If
               rsaux10.Close
               Me.txt_establecimiento = rs!establecimiento
               Me.txt_nombre_establecimiento = rs!nombre_Establecimiento
               Me.txt_referencia = rs!Referencia
               Me.txt_limite_inferior = IIf(IsNull(rs!LIMITE_INFERIOR), 0, rs!LIMITE_INFERIOR)
               Me.txt_limite_superior = IIf(IsNull(rs!LIMITE_SUPERIOR), 0, rs!LIMITE_SUPERIOR)
               
               If Len(Me.txt_limite_inferior) = 6 Then
                  Me.txt_limite_inferior = "00" + Me.txt_limite_inferior
               End If
               If Len(Me.txt_limite_inferior) = 7 Then
                  Me.txt_limite_inferior = "0" + Me.txt_limite_inferior
               End If
               
               If Len(Me.txt_limite_superior) = 6 Then
                  Me.txt_limite_superior = "00" + Me.txt_limite_superior
               End If
               If Len(Me.txt_limite_superior) = 7 Then
                  Me.txt_limite_superior = "0" + Me.txt_limite_superior
               End If
               
               
               While Not rs.EOF
                     Set list_item = Me.lv_entradas.ListItems.Add(, , rs!codigo)
                     list_item.SubItems(1) = IIf(IsNull(rs!Descripcion), "", rs!Descripcion)
                     list_item.SubItems(2) = IIf(IsNull(rs!cantidad), "", rs!cantidad)
                     list_item.SubItems(3) = 0
                     rs.MoveNext
               Wend
               Me.txt_titular.Enabled = False
               Me.txt_nombre_titular.Enabled = False
               Me.txt_cliente.Enabled = False
               Me.txt_nombre_cliente.Enabled = False
               Me.txt_establecimiento.Enabled = False
               Me.txt_nombre_establecimiento.Enabled = False
               Me.txt_referencia.Enabled = False
               Me.txt_agente.Enabled = False
               Me.txt_nombre_agente.Enabled = False
               Me.txt_codigo.Enabled = True
               Me.txt_codigo.SetFocus
               Me.txt_folio_Sello.Enabled = False
               
            Else
               MsgBox "El número de devolución no existe o no pertenece al agente seleccionado", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "La devolución fue hecha en el folio número " + CStr(var_folio_Devolucion), vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub txt_nombre_agente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
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
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
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
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
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
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
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

Private Sub txt_sello_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_sello <> "" Then
         strconsulta = "select * from xxvia_tb_devoluciones_clientes where movimiento = 'DCS' and agente = ? and referencia = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_agente)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_sello)
              .Parameters.Append parametro
         End With
         Set rsaux8 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         var_folio_Devolucion = 0
         If Not rsaux8.EOF Then
            var_folio_Devolucion = rsaux8!numero
         End If
         rsaux8.Close
         If var_folio_Devolucion = 0 Then
            rs.Open "select * from tb_Devoluciones where agente = '" + Me.txt_agente + "' and estatus = 'I' and referencia = '" + Me.txt_sello + "'", cnn_devolucion_anes, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               Me.txt_folio_Sello = rs!numero
               Me.txt_limite_inferior = 0
               Me.txt_limite_superior = 0
               strconsulta = "SELECT ACCOUNT_NUMBER TITULAR, ACCOUNT_FULL_NAME NOMBRE_TITULAR FROM XXVIA_VW_CLIENTES_BCP WHERE CUST_ACCOUNT_ID = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, rs!TITULAR)
                    .Parameters.Append parametro
               End With
               Set rsaux8 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               Me.txt_titular = rsaux8!TITULAR
               Me.txt_nombre_titular = rsaux8!NOMBRE_TITULAR
               Me.txt_clave_titular = rs!TITULAR
               rsaux8.Close
               Me.txt_cliente = rs!Cliente
               Me.txt_nombre_cliente = rs!nombre_cliente
               var_cadena = "SELECT  hcsu.price_list_id, hcsu.order_type_id,hca.cust_account_id, hcp.site_use_id AS VCHA_CLI_CLAVE_ID,hl.address1 VCHA_CLI_NOMBRE FROM hz_parties hp,hz_party_sites hps,hz_cust_accounts hca,hz_cust_acct_sites_all hcas,hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND  hcp.site_use_id = " + Me.txt_cliente
               rsaux10.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux10.EOF Then
                  var_clave_lista_precios = rsaux10!price_list_id
               Else
                  var_clave_lista_precios = 0
               End If
               rsaux10.Close
               Me.txt_establecimiento = rs!establecimiento
               Me.txt_nombre_establecimiento = rs!nombre_Establecimiento
               Me.txt_referencia = rs!Referencia
               Me.txt_limite_inferior = IIf(IsNull(rs!LIMITE_INFERIOR), 0, rs!LIMITE_INFERIOR)
               Me.txt_limite_superior = IIf(IsNull(rs!LIMITE_SUPERIOR), 0, rs!LIMITE_SUPERIOR)
               If Len(Me.txt_limite_inferior) = 6 Then
                  Me.txt_limite_inferior = "00" + Me.txt_limite_inferior
               End If
               If Len(Me.txt_limite_inferior) = 7 Then
                  Me.txt_limite_inferior = "0" + Me.txt_limite_inferior
               End If
               
               If Len(Me.txt_limite_superior) = 6 Then
                  Me.txt_limite_superior = "00" + Me.txt_limite_superior
               End If
               If Len(Me.txt_limite_superior) = 7 Then
                  Me.txt_limite_superior = "0" + Me.txt_limite_superior
               End If
               
               
               While Not rs.EOF
                     Set list_item = Me.lv_entradas.ListItems.Add(, , rs!codigo)
                     list_item.SubItems(1) = IIf(IsNull(rs!Descripcion), "", rs!Descripcion)
                     list_item.SubItems(2) = IIf(IsNull(rs!cantidad), "", rs!cantidad)
                     list_item.SubItems(3) = 0
                     rs.MoveNext
               Wend
               Me.txt_titular.Enabled = False
               Me.txt_nombre_titular.Enabled = False
               Me.txt_cliente.Enabled = False
               Me.txt_nombre_cliente.Enabled = False
               Me.txt_establecimiento.Enabled = False
               Me.txt_nombre_establecimiento.Enabled = False
               Me.txt_referencia.Enabled = False
               
               Me.txt_codigo.Enabled = True
               Me.txt_codigo.SetFocus
               Me.txt_folio_Sello.Enabled = False
               
            Else
               MsgBox "El número de devolución no existe o no pertenece al agente seleccionado", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "La devolución fue hecha en el folio número " + CStr(var_folio_Devolucion), vbOKOnly, "ATENCION"
         End If
      End If
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
           rs.Open "SELECT  distinct hp.party_name as nombre_titular,  account_number as vcha_tit_titular_id, hcas.cust_account_id AS VCHA_CLI_CLAVE_ID FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors Arc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND hcp.collector_id = " + Me.txt_agente + " AND account_number = 'T000001052'  ORDER BY hp.party_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
        Else
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
      rs.Open "SELECT  distinct hp.party_name as nombre_titular,  account_number as vcha_tit_titular_id, hcas.cust_account_id AS VCHA_CLI_CLAVE_ID FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors Arc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND hcp.collector_id = " + txt_agente + " and hcas.cust_account_id = " + Me.txt_clave_titular, cnnoracle_4, adOpenDynamic, adLockOptimistic
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
