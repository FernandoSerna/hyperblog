VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmclientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11670
   ControlBox      =   0   'False
   Icon            =   "frmclientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   Begin VB.CheckBox chk_promocion 
      Caption         =   "Promoción"
      Height          =   210
      Left            =   4965
      TabIndex        =   108
      Top             =   5535
      Width           =   1260
   End
   Begin VB.CommandButton cmd_establecimientos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2085
      Picture         =   "frmclientes.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   106
      ToolTipText     =   "Establecimientos"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame frm_busqueda_clientes 
      Height          =   3315
      Left            =   2535
      TabIndex        =   102
      Top             =   1530
      Width           =   6885
      Begin VB.TextBox txt_busqueda_cliente 
         Height          =   375
         Left            =   75
         TabIndex        =   105
         Top             =   465
         Width           =   6660
      End
      Begin MSComctlLib.ListView lv_busqueda_clientes 
         Height          =   2355
         Left            =   60
         TabIndex        =   103
         Top             =   870
         Width           =   6690
         _ExtentX        =   11800
         _ExtentY        =   4154
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
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         Caption         =   "Busqueda de clientes"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   104
         Top             =   120
         Width           =   6795
      End
   End
   Begin VB.CommandButton cmd_pedido 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2415
      Picture         =   "frmclientes.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   101
      ToolTipText     =   "Generar pedido "
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1755
      Picture         =   "frmclientes.frx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   87
      ToolTipText     =   "Cambiar de titular."
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame frm_colonias 
      Height          =   2400
      Left            =   6045
      TabIndex        =   84
      Top             =   945
      Width           =   5685
      Begin MSComctlLib.ListView lv_colonias 
         Height          =   1830
         Left            =   45
         TabIndex        =   85
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
         NumItems        =   10
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
            Text            =   "pais"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "nombre pais"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "estado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "nombre estado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "municipio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "nombre municipio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ciudad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "nombre ciudad"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_colonias 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   86
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1335
      TabIndex        =   81
      Top             =   1830
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   82
         Top             =   495
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
         TabIndex        =   83
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   " Dirección "
      Height          =   3600
      Left            =   6420
      TabIndex        =   70
      Top             =   540
      Width           =   5160
      Begin VB.TextBox txt_telefono 
         Height          =   315
         Left            =   900
         MaxLength       =   50
         TabIndex        =   46
         Top             =   2730
         Width           =   1815
      End
      Begin VB.TextBox txt_nombre_municipio 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         MaxLength       =   50
         TabIndex        =   80
         Top             =   1695
         Width           =   4155
      End
      Begin VB.TextBox txt_municipio 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         MaxLength       =   50
         TabIndex        =   78
         Top             =   1695
         Width           =   1005
      End
      Begin VB.TextBox txt_nombre_colonia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         MaxLength       =   50
         TabIndex        =   39
         Top             =   1005
         Width           =   4155
      End
      Begin VB.TextBox txt_nombre_ciudad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         MaxLength       =   50
         TabIndex        =   41
         Top             =   1350
         Width           =   4155
      End
      Begin VB.TextBox txt_nombre_estado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         MaxLength       =   50
         TabIndex        =   43
         Top             =   2040
         Width           =   4155
      End
      Begin VB.TextBox txt_nombre_pais 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         MaxLength       =   50
         TabIndex        =   45
         Top             =   2385
         Width           =   4155
      End
      Begin VB.TextBox txt_colonia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         MaxLength       =   50
         TabIndex        =   38
         Top             =   1005
         Width           =   1005
      End
      Begin VB.TextBox txt_domicilio 
         Height          =   315
         Left            =   900
         MaxLength       =   100
         TabIndex        =   36
         Top             =   300
         Width           =   4155
      End
      Begin VB.TextBox txt_codigo_postal 
         Height          =   315
         Left            =   900
         MaxLength       =   50
         TabIndex        =   37
         Top             =   645
         Width           =   1005
      End
      Begin VB.TextBox txt_ciudad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         MaxLength       =   50
         TabIndex        =   40
         Top             =   1350
         Width           =   1005
      End
      Begin VB.TextBox txt_estado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         MaxLength       =   50
         TabIndex        =   42
         Top             =   2040
         Width           =   1005
      End
      Begin VB.TextBox txt_pais 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         MaxLength       =   50
         TabIndex        =   44
         Top             =   2385
         Width           =   1005
      End
      Begin VB.TextBox txt_email 
         Height          =   315
         Left            =   900
         MaxLength       =   100
         TabIndex        =   47
         Top             =   3075
         Width           =   4155
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Télefono:"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   94
         Top             =   2835
         Width           =   675
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Municipio:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   79
         Top             =   1755
         Width           =   720
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "CP:"
         Height          =   195
         Index           =   22
         Left            =   150
         TabIndex        =   77
         Top             =   705
         Width           =   255
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio:"
         Height          =   195
         Index           =   21
         Left            =   135
         TabIndex        =   76
         Top             =   315
         Width           =   675
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Colonia:"
         Height          =   195
         Index           =   20
         Left            =   135
         TabIndex        =   75
         Top             =   1065
         Width           =   570
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Index           =   19
         Left            =   135
         TabIndex        =   74
         Top             =   1410
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   18
         Left            =   135
         TabIndex        =   73
         Top             =   2100
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Index           =   17
         Left            =   135
         TabIndex        =   72
         Top             =   2445
         Width           =   345
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
         Height          =   195
         Index           =   16
         Left            =   150
         TabIndex        =   71
         Top             =   3180
         Width           =   435
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   6435
      TabIndex        =   66
      Top             =   4200
      Width           =   5160
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1755
         TabIndex        =   67
         Top             =   165
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3585
         TabIndex        =   68
         Top             =   165
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ir al primero"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un Registro Atras"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un registro adelante"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ir al ultimo"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de cliente:"
         Height          =   195
         Left            =   195
         TabIndex        =   69
         Top             =   210
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2460
      Left            =   6435
      TabIndex        =   64
      Top             =   4770
      Width           =   5160
      Begin MSComctlLib.ListView lv_clientes 
         Height          =   2265
         Left            =   45
         TabIndex        =   65
         Top             =   135
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   3995
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   48
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6262
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "representante"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "fecha"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "agente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ruta"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "curp"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "rfc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "moneda"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "plazo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "tipo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "lista"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "canal"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "transporte"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "clave agrupador"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "agrupador"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "estatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "titular"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Prioridad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "email"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "pais"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "estado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "ciudad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "colonia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "domicilio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "cp"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "municipio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "Enviar facturas"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "Asigna Catalogos"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Text            =   "anterior"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Text            =   "Pedido"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Text            =   "referencia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   32
            Text            =   "franquisia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   33
            Text            =   "Telefono"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   34
            Text            =   "trazabilidad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   35
            Text            =   "no activo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   36
            Text            =   "clave unificada"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   37
            Text            =   "unificador"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   38
            Text            =   "venta al publico en general"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   39
            Text            =   "Promocion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   40
            Text            =   "NOMBRE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   41
            Text            =   "APELLIDO PATERNO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   42
            Text            =   "APELLIDO MATERNO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   43
            Text            =   "NUMERO "
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   44
            Text            =   "CLAVE TEL PAIS"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   45
            Text            =   "CLAVE TEL ESTADO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   46
            Text            =   "calle"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   47
            Text            =   "numero_interno"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmclientes.frx":0FC0
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmclientes.frx":10C2
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   765
      Picture         =   "frmclientes.frx":11C4
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1095
      Picture         =   "frmclientes.frx":1296
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1425
      Picture         =   "frmclientes.frx":1398
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11220
      Picture         =   "frmclientes.frx":149A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   60
      TabIndex        =   54
      Top             =   285
      Width           =   11475
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos del cliente "
      Height          =   6690
      Left            =   60
      TabIndex        =   48
      Top             =   540
      Width           =   6315
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5865
         Picture         =   "frmclientes.frx":1AD4
         Style           =   1  'Graphical
         TabIndex        =   109
         ToolTipText     =   "Generar pedido "
         Top             =   660
         Width           =   330
      End
      Begin VB.CheckBox chk_venta_publico_general 
         Caption         =   "Venta al publico general"
         Height          =   195
         Left            =   3060
         TabIndex        =   107
         Top             =   4283
         Width           =   2790
      End
      Begin VB.TextBox txt_clave_unificada 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1485
         TabIndex        =   99
         Top             =   6300
         Width           =   1680
      End
      Begin VB.TextBox txt_unificador 
         Enabled         =   0   'False
         Height          =   330
         Left            =   4230
         TabIndex        =   98
         Top             =   6300
         Width           =   1995
      End
      Begin VB.CheckBox chk_activo 
         Caption         =   "No activo"
         Height          =   345
         Left            =   4455
         TabIndex        =   96
         Top             =   5310
         Width           =   1305
      End
      Begin VB.CheckBox chk_trazabilidad 
         Caption         =   "Trazabilidad"
         Height          =   210
         Left            =   3660
         TabIndex        =   95
         Top             =   4980
         Width           =   1260
      End
      Begin VB.CheckBox chk_franquicia 
         Caption         =   "Franquicia"
         Height          =   210
         Left            =   2505
         TabIndex        =   93
         Top             =   4980
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.TextBox txt_referencia 
         Height          =   330
         Left            =   4230
         TabIndex        =   92
         Top             =   5955
         Width           =   1995
      End
      Begin VB.CheckBox chk_pedido 
         Caption         =   "Cliente para pedido de tienda"
         Height          =   210
         Left            =   3510
         TabIndex        =   90
         Top             =   5700
         Width           =   2685
      End
      Begin VB.TextBox txt_anterior 
         Height          =   330
         Left            =   1485
         TabIndex        =   88
         Top             =   5955
         Width           =   1680
      End
      Begin MSComCtl2.MonthView mes 
         Height          =   2370
         Index           =   0
         Left            =   2085
         TabIndex        =   12
         Top             =   750
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   71958529
         CurrentDate     =   37581
      End
      Begin VB.TextBox txt_nombre_ruta 
         Height          =   315
         Left            =   2445
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1716
         Width           =   3750
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   315
         Left            =   2445
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1362
         Width           =   3750
      End
      Begin VB.TextBox txt_nombre_tipo_cliente 
         Height          =   315
         Left            =   2445
         MaxLength       =   50
         TabIndex        =   24
         Top             =   3120
         Width           =   3750
      End
      Begin VB.TextBox txt_nombre_plazo 
         Height          =   315
         Left            =   2445
         MaxLength       =   50
         TabIndex        =   22
         Top             =   2775
         Width           =   3750
      End
      Begin VB.TextBox txt_nombre_transporte 
         Height          =   315
         Left            =   2445
         MaxLength       =   50
         TabIndex        =   28
         Top             =   3840
         Width           =   3750
      End
      Begin VB.TextBox txt_nombre_lista_precios 
         Height          =   315
         Left            =   2445
         MaxLength       =   50
         TabIndex        =   26
         Top             =   3480
         Width           =   3750
      End
      Begin VB.TextBox txt_nombre_agrupador 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2460
         MaxLength       =   50
         TabIndex        =   31
         Top             =   4575
         Width           =   3750
      End
      Begin VB.TextBox txt_nombre_moneda 
         Height          =   315
         Left            =   2445
         MaxLength       =   50
         TabIndex        =   20
         Top             =   2415
         Width           =   3750
      End
      Begin VB.CheckBox chk_asignacion_catalogos 
         Caption         =   "Asignación Catálogos"
         Height          =   210
         Left            =   1470
         TabIndex        =   35
         Top             =   5700
         Width           =   1890
      End
      Begin VB.CheckBox chk_enviar_facturas 
         Caption         =   "Enviar Facturas"
         Height          =   210
         Left            =   2745
         TabIndex        =   34
         Top             =   5400
         Width           =   1515
      End
      Begin VB.CheckBox chk_usa_agrupador 
         Caption         =   "Usar agrupador"
         Height          =   210
         Left            =   1470
         TabIndex        =   29
         Top             =   4275
         Width           =   1410
      End
      Begin VB.CheckBox chk_estatus 
         Caption         =   "Estatus"
         Height          =   210
         Left            =   1470
         TabIndex        =   33
         Top             =   5385
         Width           =   855
      End
      Begin VB.CommandButton cmdfecha 
         Height          =   285
         Index           =   0
         Left            =   2760
         Picture         =   "frmclientes.frx":1BD6
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Seleccione la fecha"
         Top             =   1035
         Width           =   315
      End
      Begin VB.TextBox txt_moneda 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   19
         Top             =   2415
         Width           =   960
      End
      Begin VB.TextBox txt_agrupador 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   30
         Top             =   4575
         Width           =   975
      End
      Begin VB.TextBox txt_lista_precios 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   25
         Top             =   3480
         Width           =   960
      End
      Begin VB.TextBox txt_transporte 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   27
         Top             =   3840
         Width           =   960
      End
      Begin VB.TextBox txt_rfc 
         Height          =   315
         Left            =   4125
         MaxLength       =   50
         TabIndex        =   18
         Top             =   2085
         Width           =   2085
      End
      Begin VB.TextBox txt_plazo 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   21
         Top             =   2775
         Width           =   960
      End
      Begin VB.TextBox txt_tipo_cliente 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   23
         Top             =   3120
         Width           =   960
      End
      Begin VB.TextBox txt_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1470
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         Top             =   300
         Width           =   1140
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2625
         MaxLength       =   100
         TabIndex        =   8
         Top             =   300
         Width           =   3570
      End
      Begin VB.TextBox txt_representante 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   9
         Top             =   654
         Width           =   4350
      End
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1350
         Width           =   960
      End
      Begin VB.TextBox txt_ruta 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1716
         Width           =   960
      End
      Begin VB.TextBox txt_curp 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   17
         Top             =   2070
         Width           =   2070
      End
      Begin VB.TextBox txt_fecha_captura 
         Height          =   315
         Left            =   1470
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1020
         Width           =   1230
      End
      Begin VB.TextBox txt_prioridad 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   32
         Top             =   4935
         Width           =   960
      End
      Begin VB.Label Label5 
         Caption         =   "Unificador:"
         Height          =   240
         Left            =   3360
         TabIndex        =   100
         Top             =   6345
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Clave unificada:"
         Height          =   195
         Left            =   120
         TabIndex        =   97
         Top             =   6368
         Width           =   1140
      End
      Begin VB.Label Label3 
         Caption         =   "Referencia:"
         Height          =   240
         Left            =   3375
         TabIndex        =   91
         Top             =   6000
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "Anterior:"
         Height          =   300
         Left            =   150
         TabIndex        =   89
         Top             =   5970
         Width           =   930
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Prioridad:"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   63
         Top             =   4995
         Width           =   660
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Fam. agrupadores:"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   62
         Top             =   4635
         Width           =   1320
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Transporte:"
         Height          =   195
         Index           =   13
         Left            =   135
         TabIndex        =   61
         Top             =   3900
         Width           =   810
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         Height          =   195
         Index           =   15
         Left            =   135
         TabIndex        =   60
         Top             =   2475
         Width           =   630
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "CURP:"
         Height          =   195
         Index           =   14
         Left            =   135
         TabIndex        =   59
         Top             =   2130
         Width           =   495
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Plazo:"
         Height          =   195
         Index           =   12
         Left            =   135
         TabIndex        =   58
         Top             =   2835
         Width           =   435
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Lista de precios:"
         Height          =   195
         Index           =   10
         Left            =   135
         TabIndex        =   57
         Top             =   3540
         Width           =   1155
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   56
         Top             =   3180
         Width           =   360
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "RFC:"
         Height          =   195
         Index           =   8
         Left            =   3615
         TabIndex        =   55
         Top             =   2130
         Width           =   360
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   53
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Representante:"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   52
         Top             =   714
         Width           =   1095
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Fecha captura:"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   51
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   50
         Top             =   1776
         Width           =   390
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   49
         Top             =   1422
         Width           =   555
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7245
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   7680
      Top             =   -15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientes.frx":1CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientes.frx":25B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4065
      Top             =   -180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientes.frx":2E8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientes.frx":3766
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientes.frx":4040
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientes.frx":45DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientes.frx":4EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientes.frx":5792
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientes.frx":606C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientes.frx":617E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientes.frx":6290
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientes.frx":63A2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmclientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_hubo_cambios As Boolean
Dim var_guardar_cambios As Boolean
Dim bitacora As Boolean
Dim numero_items_clientes As Integer
Dim var_tipo_lista As Integer
Dim var_barra As String




Private Sub chk_activo_Click()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub chk_activo_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub chk_asignacion_catalogos_Click()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub chk_asignacion_catalogos_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub chk_asignacion_catalogos_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub chk_enviar_facturas_Click()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub chk_enviar_facturas_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub chk_enviar_facturas_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub chk_estatus_Click()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub chk_estatus_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub chk_estatus_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub


Private Sub chk_franquicia_Click()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub chk_franquicia_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub chk_pedido_Click()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub chk_pedido_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub chk_promocion_Click()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub chk_trazabilidad_Click()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub chk_trazabilidad_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub chk_usa_agrupador_Click()
   var_hubo_cambios = True
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
   If chk_usa_agrupador = 1 Then
      txt_agrupador.Enabled = True
   Else
      txt_agrupador.Enabled = False
   End If
End Sub

Private Sub chk_usa_agrupador_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub chk_venta_publico_general_Click()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub cmd_deshacer_Click()
   Call pro_textos
End Sub

Private Sub cmd_deshacer_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub cmd_eliminar_Click()
   var_opcion_seguridad = 2
   var_acepta_seguridad = 1
   If var_global_permiso3 = 1 Then
      var_acepta_seguridad = 2
      If var_global_permiso4 = 1 Then
         frmpasswords2.Show 1
      Else
         frmpasswords.Show 1
      End If
   End If
   If var_acepta_seguridad = 1 Then
      'rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0", cnn_distribucion, adOpenDynamic, adLockOptimistic
      'While Not rsaux5.EOF
      '      var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
      '      If Trim(var_conexion_importacion) <> "" Then
      '         If cnn.State = 1 Then
      '            cnn.Close
      '         End If
      '         cnn.Open var_conexion_importacion
               Call pro_elimina_clientes
      '      End If
      '      rsaux5.MoveNext
      'Wend
      'rsaux5.Close
      
      MsgBox "Se Elimino Correctamente el Registro", vbInformation
      lv_clientes.ListItems.Remove (lv_clientes.selectedItem.Index)
      Call pro_limpiatextos(Me)
      txt_registros = lv_clientes.ListItems.Count
      lv_clientes.selectedItem.Selected = True
      pro_textos
      
      
      
      rs.Open "select * from tb_clientes where vcha_tit_titular_id = '" + vartitular + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If rs.BOF Then
         cmd_guardar.Enabled = False
         cmd_deshacer.Enabled = False
         cmd_eliminar.Enabled = False
      Else
         cmd_guardar.Enabled = True
         cmd_deshacer.Enabled = True
         cmd_eliminar.Enabled = True
      End If
      rs.Close
   Else
      MsgBox "Imposible realizar la acción solicitada", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_establecimientos_Click()
   var_opcion_seguridad = 2
   var_cliente_pedido_internet = Me.txt_cliente
   var_activa_forma_establecimientos = "frmclientes"
   frmclientes.Enabled = False
   frmestablecimientos.Show
End Sub

Private Sub cmd_eliminar_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub cmd_guardar_Click()
Dim var_posible As Boolean
Dim numero_establecimiento As Double
   var_posible = True
   If var_posible = True Then
      If txt_nombre_cliente = "" Or txt_representante = "" Or txt_fecha_captura = "" Or txt_agente = "" Or txt_ruta = "" Or txt_moneda = "" Or txt_plazo = "" Or txt_tipo_cliente = "" Or txt_lista_precios = "" Then
         MsgBox "Información incompleta", vbOKOnly, "ATENCION"
      Else
         var_resultado = InStr(1, var_menus, Me.Caption)
         var_inicio = var_resultado + Len(Me.Caption) + 3
         var_opcion_seguridad = 2
         var_acepta_seguridad = 1
         If var_global_permiso3 = 1 Then
            var_acepta_seguridad = 2
            If var_global_permiso4 = 1 Then
               frmpasswords2.Show 1
            Else
               frmpasswords.Show
            End If
         End If
         If var_acepta_seguridad = 1 Then
            rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0 and vcha_emp_empresa_id = '02' ORDER BY INTE_EMP_ORDEN_CONEXION", cnn_distribucion, adOpenDynamic, adLockOptimistic
            While Not rsaux5.EOF
                  var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
                  If Trim(var_conexion_importacion) <> "" Then
                     If cnn_importacion.State = 1 Then
                        cnn_importacion.Close
                     End If
                     cnn_importacion.Open var_conexion_importacion
                     Call pro_guardar_clientes
                  End If
                  rsaux5.MoveNext
            Wend
            rsaux5.Close
            
            If (UCase(parametros(0)) = "sqlquezada2" Or UCase(parametros(0)) = "DBPRUEBAS") And var_empresa = "31" Then
               rs.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
               If rs.EOF Then
                  rsaux.Open "select * from tb_clientes WHERE VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                  var_cadena = "INSERT INTO "
                  var_cadena = var_cadena + "TB_CLIENTES  (VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_CLI_REPRESENTANTE, DTIM_CLI_FECHA_CAPTURA, VCHA_AGE_AGENTE_ID, VCHA_RUT_RUTA_ID, VCHA_CLI_CURP, VCHA_CLI_RFC, VCHA_MON_MONEDA_ID, VCHA_PLA_PLAZO_ID, VCHA_TCL_TIPO_CLIENTE_ID, VCHA_LIS_LISTA_ID, VCHA_TRA_TRANSPORTE_ID, VCHA_FAG_FAMILIA_AGRUPADOR_ID, INTE_CLI_AGRUPADOR, INTE_CLI_ESTATUS, VCHA_TIT_TITULAR_ID, CHAR_PRI_PRIORIDAD_ID, VCHA_CLI_EMAIL, VCHA_PAI_PAIS_ID, VCHA_EST_ESTADO_ID, VCHA_MUN_MUNICIPIO_ID, VCHA_CIU_CIUDAD_ID, VCHA_CLI_COLONIA, VCHA_CLI_DIRECCION, VCHA_CLI_CP, INTE_CLI_ENVIO_FACTURA, INTE_CLI_ASIGNACION_CATALOGOS, VCHA_CLI_CLAVE_ANTERIOR_ID, VCHA_EMP_EMPRESA_ID, INTE_CLI_CLIENTE_PEDIDO_TIENDA, INTE_CLI_PERSONA_FISICA, NUM_INTER_TRANC_TYPE, NUM_INTER_UPLOADED, DATE_INTER_DATE, VCHA_CLI_REFERENCIA, TEXTILERA, VCHA_CLI_TIENDA, VCHA_CLI_CLAVE_TIENDA, DTIM_INT_FECHA, INTE_INT_INTERFACE, INTE_CLI_FRANQUICIA, VCHA_CLI_TELEFONO, INTE_CLI_TRAZABILIDAD , Referencia, VCHA_SRU_SUBRUTA_ID,"
                  var_cadena = var_cadena + " INTE_CLI_ACTIVO, VCHA_CLI_CLAVE_UNIFICADA_ID, INTE_CLI_UNIFICADOR)  "
                  var_cadena = var_cadena + " values ('" + IIf(IsNull(rsaux!vcha_cli_clave_id), "", rsaux!vcha_cli_clave_id) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_NOMBRE), "", rsaux!VCHA_CLI_NOMBRE) + "', '" + IIf(IsNull(rsaux!vcha_cli_representante), "", rsaux!vcha_cli_representante) + "', "
                  var_cadena = var_cadena + " " + Format(IIf(IsNull(rsaux!dtim_cli_fecha_Captura), Date, rsaux!dtim_cli_fecha_Captura), "Short Date") + ", '" + IIf(IsNull(rsaux!VCHA_AGE_AGENTE_ID), "", rsaux!VCHA_AGE_AGENTE_ID) + "', '" + IIf(IsNull(rsaux!vcha_rut_ruta_id), "", rsaux!vcha_rut_ruta_id) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_CURP), "", rsaux!VCHA_CLI_CURP) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_RFC), "", rsaux!VCHA_CLI_RFC) + "', '" + IIf(IsNull(rsaux!vcha_mon_moneda_id), "", rsaux!vcha_mon_moneda_id) + "', '" + IIf(IsNull(rsaux!VCHA_PLA_PLAZO_ID), "", rsaux!VCHA_PLA_PLAZO_ID) + "', '" + IIf(IsNull(rsaux!VCHA_TCL_TIPO_CLIENTE_ID), "", rsaux!VCHA_TCL_TIPO_CLIENTE_ID) + "', '" + IIf(IsNull(rsaux!vcha_LIS_LISTA_iD), "", rsaux!vcha_LIS_LISTA_iD) + "', '" + IIf(IsNull(rsaux!VCHA_TRA_TRANSPORTE_ID), "", rsaux!VCHA_TRA_TRANSPORTE_ID) + "','"
                  var_cadena = var_cadena + IIf(IsNull(rsaux!VCHA_FAG_FAMILIA_AGRUPADOR_ID), "", rsaux!VCHA_FAG_FAMILIA_AGRUPADOR_ID) + "' , " + CStr(IIf(IsNull(rsaux!INTE_CLI_AGRUPADOR), 0, rsaux!INTE_CLI_AGRUPADOR)) + ", " + CStr(IIf(IsNull(rsaux!INTE_CLI_ESTATUS), 0, rsaux!INTE_CLI_ESTATUS)) + ", '" + IIf(IsNull(rsaux!vcha_tit_titular_id), "", rsaux!vcha_tit_titular_id) + "', '" + IIf(IsNull(rsaux!CHAR_PRI_PRIORIDAD_ID), "", rsaux!CHAR_PRI_PRIORIDAD_ID) + "', '" + IIf(IsNull(rsaux!vcha_cli_email), "", rsaux!vcha_cli_email) + "', '"
                  var_cadena = var_cadena + IIf(IsNull(rsaux!VCHA_PAI_PAIS_ID), "", rsaux!VCHA_PAI_PAIS_ID) + "', '" + IIf(IsNull(rsaux!VCHA_EST_ESTADO_ID), "", rsaux!VCHA_EST_ESTADO_ID) + "', '" + IIf(IsNull(rsaux!VCHA_MUN_MUNICIPIO_ID), "", rsaux!VCHA_MUN_MUNICIPIO_ID) + "', '"
                  var_cadena = var_cadena + IIf(IsNull(rsaux!VCHA_CIU_CIUDAD_ID), "", rsaux!VCHA_CIU_CIUDAD_ID) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_COLONIA), "", rsaux!VCHA_CLI_COLONIA) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_DIRECCION), "", rsaux!VCHA_CLI_DIRECCION) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_CP), "", rsaux!VCHA_CLI_CP) + "',"
                  var_cadena = var_cadena + CStr(IIf(IsNull(rsaux!INTE_CLI_ENVIO_FACTURA), 0, rsaux!INTE_CLI_ENVIO_FACTURA)) + ", " + CStr(IIf(IsNull(rsaux!INTE_CLI_ASIGNACION_CATALOGOS), 0, rsaux!INTE_CLI_ASIGNACION_CATALOGOS)) + ", '"
                  var_cadena = var_cadena + IIf(IsNull(rsaux!vcha_cli_clave_anterior_id), "", rsaux!vcha_cli_clave_anterior_id) + "' , '" + IIf(IsNull(rsaux!VCHA_EMP_EMPRESA_ID), "", rsaux!VCHA_EMP_EMPRESA_ID) + "', " + CStr(IIf(IsNull(rsaux!INTE_CLI_CLIENTE_PEDIDO_TIENDA), 0, rsaux!INTE_CLI_CLIENTE_PEDIDO_TIENDA)) + ", " + CStr(IIf(IsNull(rsaux!inte_cli_persona_fisica), 0, rsaux!inte_cli_persona_fisica)) + ", " + CStr(IIf(IsNull(rsaux!num_inter_tranc_type), 0, rsaux!num_inter_tranc_type)) + ", " + CStr(IIf(IsNull(rsaux!NUM_INTER_UPLOADED), "", rsaux!NUM_INTER_UPLOADED)) + ", " + Format(IIf(IsNull(rsaux!date_inter_date), Date, rsaux!date_inter_date), "Short Date") + ", '" + IIf(IsNull(rsaux!VCHA_CLI_REFERENCIA), "", rsaux!VCHA_CLI_REFERENCIA) + "', '" + IIf(IsNull(rsaux!TEXTILERA), "", rsaux!TEXTILERA) + "', '" + IIf(IsNull(rsaux!vcha_cli_tienda), "", rsaux!vcha_cli_tienda) + "', '" + IIf(IsNull(rsaux!vcha_cli_clave_tienda), "", rsaux!vcha_cli_clave_tienda) + "', "
                  var_cadena = var_cadena + Format(IIf(IsNull(rsaux!DTIM_INT_FECHA), Date, rsaux!DTIM_INT_FECHA), "Short Date") + ", " + CStr(IIf(IsNull(rsaux!INTE_INT_INTERFACE), 0, rsaux!INTE_INT_INTERFACE)) + ", " + CStr(IIf(IsNull(rsaux!INTE_CLI_FRANQUICIA), 0, rsaux!INTE_CLI_FRANQUICIA)) + ", '" + IIf(IsNull(rsaux!vcha_cli_telefono), "", rsaux!vcha_cli_telefono) + "', " + CStr(IIf(IsNull(rsaux!INTE_CLI_TRAZABILIDAD), 0, rsaux!INTE_CLI_TRAZABILIDAD)) + ", '" + IIf(IsNull(rsaux!Referencia), "", rsaux!Referencia) + "', '" + IIf(IsNull(rsaux!vcha_sru_subruta_id), "", rsaux!vcha_sru_subruta_id) + "',"
                  var_cadena = var_cadena + CStr(IIf(IsNull(rsaux!inte_cli_activo), "", rsaux!inte_cli_activo)) + ", '" + IIf(IsNull(rsaux!vcha_cli_clave_unificada_id), "", rsaux!vcha_cli_clave_unificada_id) + "', " + CStr(IIf(IsNull(rsaux!inte_cli_unificador), 0, rsaux!inte_cli_unificador)) + ")"
                  rsaux.Close
                  rsaux2.Open var_cadena, cnn_distribucion, adOpenDynamic, adLockOptimistic
               Else
                  rsaux2.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                  var_cadena = "UPDATE TB_CLIENTES SET VCHA_CLI_NOMBRE = '" + rsaux2!VCHA_CLI_NOMBRE + "', VCHA_CLI_REPRESENTANTE = '" + IIf(IsNull(rsaux2!vcha_cli_representante), "", rsaux2!vcha_cli_representante) + "', DTIM_CLI_FECHA_CAPTURA = " + CStr(IIf(IsNull(rsaux2!dtim_cli_fecha_Captura), "", rsaux2!dtim_cli_fecha_Captura)) + ", VCHA_AGE_AGENTE_ID = '" + IIf(IsNull(rsaux2!VCHA_AGE_AGENTE_ID), "", rsaux2!VCHA_AGE_AGENTE_ID) + "', VCHA_RUT_RUTA_ID = '" + IIf(IsNull(rsaux2!vcha_rut_ruta_id), "", rsaux2!vcha_rut_ruta_id) + "', VCHA_CLI_CURP = '" + IIf(IsNull(rsaux2!VCHA_CLI_CURP), "", rsaux2!VCHA_CLI_CURP) + "', VCHA_CLI_RFC = '" + IIf(IsNull(rsaux2!VCHA_CLI_RFC), "", rsaux2!VCHA_CLI_RFC) + "', VCHA_MON_MONEDA_ID = '" + IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id) + "', VCHA_PLA_PLAZO_ID = '" + IIf(IsNull(rsaux2!VCHA_PLA_PLAZO_ID), "", rsaux2!VCHA_PLA_PLAZO_ID) + "', VCHA_TCL_TIPO_CLIENTE_ID = '"
                  var_cadena = var_cadena + IIf(IsNull(rsaux2!VCHA_TCL_TIPO_CLIENTE_ID), "", rsaux2!VCHA_TCL_TIPO_CLIENTE_ID) + "',VCHA_LIS_LISTA_ID = '" + IIf(IsNull(rsaux2!vcha_LIS_LISTA_iD), "", rsaux2!vcha_LIS_LISTA_iD) + "', VCHA_TRA_TRANSPORTE_ID = '" + IIf(IsNull(rsaux2!VCHA_TRA_TRANSPORTE_ID), "", rsaux2!VCHA_TRA_TRANSPORTE_ID) + "', VCHA_FAG_FAMILIA_AGRUPADOR_ID = '" + IIf(IsNull(rsaux2!VCHA_FAG_FAMILIA_AGRUPADOR_ID), "", rsaux2!VCHA_FAG_FAMILIA_AGRUPADOR_ID) + "', INTE_CLI_AGRUPADOR = " + CStr(IIf(IsNull(rsaux2!INTE_CLI_AGRUPADOR), 0, rsaux2!INTE_CLI_AGRUPADOR)) + ", INTE_CLI_ESTATUS = " + CStr(IIf(IsNull(rsaux2!INTE_CLI_ESTATUS), 0, rsaux2!INTE_CLI_ESTATUS)) + ", VCHA_TIT_TITULAR_ID = '" + IIf(IsNull(rsaux2!vcha_tit_titular_id), "", rsaux2!vcha_tit_titular_id) + "', CHAR_PRI_PRIORIDAD_ID = '" + IIf(IsNull(rsaux2!CHAR_PRI_PRIORIDAD_ID), "", rsaux2!CHAR_PRI_PRIORIDAD_ID) + "', VCHA_CLI_EMAIL = '" + IIf(IsNull(rsaux2!vcha_cli_email), "", rsaux2!vcha_cli_email) + "', "
                  var_cadena = var_cadena + " VCHA_PAI_PAIS_ID = '" + IIf(IsNull(rsaux2!VCHA_PAI_PAIS_ID), "", rsaux2!VCHA_PAI_PAIS_ID) + "', VCHA_EST_ESTADO_ID = '" + IIf(IsNull(rsaux2!VCHA_EST_ESTADO_ID), "", rsaux2!VCHA_EST_ESTADO_ID) + "', VCHA_MUN_MUNICIPIO_ID = '" + IIf(IsNull(rsaux2!VCHA_MUN_MUNICIPIO_ID), "", rsaux2!VCHA_MUN_MUNICIPIO_ID) + "', VCHA_CIU_CIUDAD_ID = '" + IIf(IsNull(rsaux2!VCHA_CIU_CIUDAD_ID), "", rsaux2!VCHA_CIU_CIUDAD_ID) + "', VCHA_cLI_COLONIA = '" + IIf(IsNull(rsaux2!VCHA_CLI_COLONIA), "", rsaux2!VCHA_CLI_COLONIA) + "', VCHA_CLI_DIRECCION = '" + IIf(IsNull(rsaux2!VCHA_CLI_DIRECCION), "", rsaux2!VCHA_CLI_DIRECCION) + "', VCHA_CLI_CP ='" + IIf(IsNull(rsaux2!VCHA_CLI_CP), "", rsaux2!VCHA_CLI_CP) + "',INTE_CLI_ENVIO_FACTURA= " + CStr(IIf(IsNull(rsaux2!INTE_CLI_ENVIO_FACTURA), 0, rsaux2!INTE_CLI_ENVIO_FACTURA)) + ", INTE_CLI_ASIGNACION_CATALOGOS = " + CStr(IIf(IsNull(rsaux2!INTE_CLI_ASIGNACION_CATALOGOS), 0, rsaux2!INTE_CLI_ASIGNACION_CATALOGOS)) + ", VCHA_CLI_CLAVE_ANTERIOR_ID = '"
                  var_cadena = var_cadena + IIf(IsNull(rsaux2!vcha_cli_clave_anterior_id), "", rsaux2!vcha_cli_clave_anterior_id) + "  ', VCHA_EMP_EMPRESA_ID = '" + IIf(IsNull(rsaux2!VCHA_EMP_EMPRESA_ID), "", rsaux2!VCHA_EMP_EMPRESA_ID) + "', INTE_CLI_CLIENTE_PEDIDO_TIENDA = " + CStr(IIf(IsNull(rsaux2!INTE_CLI_CLIENTE_PEDIDO_TIENDA), 0, rsaux2!INTE_CLI_CLIENTE_PEDIDO_TIENDA)) + ", INTE_CLI_PERSONA_FISICA = " + CStr(IIf(IsNull(rsaux2!inte_cli_persona_fisica), 0, rsaux2!inte_cli_persona_fisica)) + ", TEXTILERA = '" + IIf(IsNull(rsaux2!TEXTILERA), "", rsaux2!TEXTILERA) + "', VCHA_CLI_REFERENCIA = '" + IIf(IsNull(rsaux2!VCHA_CLI_REFERENCIA), "", rsaux2!VCHA_CLI_REFERENCIA) + "', VCHA_CLI_TIENDA = '" + IIf(IsNull(rsaux2!vcha_cli_tienda), "", rsaux2!vcha_cli_tienda) + "', VCHA_CLI_CLAVE_TIENDA = '" + IIf(IsNull(rsaux2!vcha_cli_clave_tienda), "", rsaux2!vcha_cli_clave_tienda) + "', INTE_CLI_FRANQUICIA = " + CStr(IIf(IsNull(rsaux2!INTE_CLI_FRANQUICIA), 0, rsaux2!INTE_CLI_FRANQUICIA)) + ",  "
                  var_cadena = var_cadena + " VCHA_CLI_TELEFONO = '" + IIf(IsNull(rsaux2!vcha_cli_telefono), "", rsaux2!vcha_cli_telefono) + "', INTE_CLI_TRAZABILIDAD = " + CStr(IIf(IsNull(rsaux2!INTE_CLI_TRAZABILIDAD), 0, rsaux2!INTE_CLI_TRAZABILIDAD)) + " WHERE VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "'"
                  rsaux3.Open var_cadena, cnn_distribucion, adOpenDynamic, adLockOptimistic
                  rsaux2.Close
               End If
               rs.Close
               'cnn_distribucion.BeginTrans
               'cnn.BeginTrans
               'rs.Open "SELECT * FROM TB_DETALLE_ESTABLECIMIENTOS WHERE VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
               'If rs.EOF Then
               '   rsaux1.Open "SELECT MAX(VCHA_ESB_ESTABLECIMIENTO_ID) FROM TB_ESTABLECIMIENTOS", cnn_distribucion, adOpenDynamic, adLockOptimistic
               '   numero_establecimiento = CDbl(Mid(IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value), 2, 90)) + 1
               '   rsaux1.Close
               '   clave_establecimiento = Trim(CStr(numero_establecimiento))
               '   If Len(Trim(clave_establecimiento)) = 1 Then
               '      clave_establecimiento = "E00000000" + Trim(clave_establecimiento)
               '   End If
               '   If Len(Trim(clave_establecimiento)) = 2 Then
               '      clave_establecimiento = "E0000000" + Trim(clave_establecimiento)
               '   End If
               '   If Len(Trim(clave_establecimiento)) = 3 Then
               '      clave_establecimiento = "E000000" + Trim(clave_establecimiento)
               '   End If
               '   If Len(Trim(clave_establecimiento)) = 4 Then
               '      clave_establecimiento = "E00000" + Trim(clave_establecimiento)
               '   End If
               '   If Len(Trim(clave_establecimiento)) = 5 Then
               '      clave_establecimiento = "E0000" + Trim(clave_establecimiento)
               '   End If
               '   If Len(Trim(clave_establecimiento)) = 6 Then
               '      clave_establecimiento = "E000" + Trim(clave_establecimiento)
               '   End If
               '   If Len(Trim(clave_establecimiento)) = 7 Then
               '      clave_establecimiento = "E00" + Trim(clave_establecimiento)
               '   End If
               '   If Len(Trim(clave_establecimiento)) = 8 Then
               '      clave_establecimiento = "E0" + Trim(clave_establecimiento)
               '   End If
               '   If Len(Trim(clave_establecimiento)) = 9 Then
               '      clave_establecimiento = "E" + Trim(clave_establecimiento)
               '   End If
               '
               '   var_cadena = "INSERT INTO TB_ESTABLECIMIENTOS (VCHA_TIT_TITULAR_ID, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_ESB_NOMBRE, VCHA_PAI_PAIS_ID, VCHA_EST_ESTADO_ID, VCHA_CIU_CIUDAD_ID, VCHA_COL_COLONIA_ID, VCHA_ESB_DOMICILIO, VCHA_ESB_TELEFONO, CHAR_ESB_FACTURA_CATALOGOS, VCHA_MUN_MUNICIPIO_ID, VCHA_ESB_CP, vcha_emp_empresa_id)"
               '   var_cadena = var_cadena + " Values ('" + vartitular + "','" + clave_establecimiento + "', '" + Me.txt_nombre_cliente + "','" + Me.txt_pais + "', '" + Me.txt_estado + "','" + Me.txt_ciudad + "', '" + Me.txt_colonia + "','" + Me.txt_domicilio + "', '" + Me.txt_telefono + "',0,'" + Me.txt_municipio + "','" + Me.txt_codigo_postal + "','" + var_empresa + "')"
               '   rsaux1.Open var_cadena, cnn_distribucion, adOpenDynamic, adLockOptimistic
               '   rsaux1.Open var_cadena, cnn_distribucion, adOpenDynamic, adLockOptimistic
               '   rsaux1.Open "insert into tb_detalle_establecimientos (vcha_cli_clave_id, vcha_esb_establecimiento_id) values ('" + Me.txt_cliente + "','" + clave_establecimiento + "')", cnn_importacion, adOpenDynamic, adLockOptimistic
               '   rsaux1.Open "insert into tb_detalle_establecimientos (vcha_cli_clave_id, vcha_esb_establecimiento_id) values ('" + Me.txt_cliente + "','" + clave_establecimiento + "')", cnn_distribucion, adOpenDynamic, adLockOptimistic
               'Else
               '   var_cadena = "Update tb_establecimientos set VCHA_TIT_TITULAR_ID = '" + vartitular + "', VCHA_ESB_NOMBRE = '" + Me.txt_nombre_cliente + "', VCHA_PAI_PAIS_ID = '" + Me.txt_pais + "', VCHA_EST_ESTADO_ID = '" + Me.txt_estado + "', VCHA_CIU_CIUDAD_ID = '" + Me.txt_ciudad + "', VCHA_COL_COLONIA_ID = '" + Me.txt_colonia + "', VCHA_ESB_DOMICILIO = '" + Me.txt_domicilio + "', VCHA_ESB_TELEFONO = '" + Me.txt_telefono + "', CHAR_ESB_FACTURA_CATALOGOS = '', VCHA_MUN_MUNICIPIO_ID = '" + Me.txt_municipio + "', VCHA_ESB_CP = '" + Me.txt_codigo_postal + "' where VCHA_ESB_ESTABLECIMIENTO_ID = '" + rs!vcha_ESB_ESTABLECIMIENTO_id + "'"
               '   rsaux1.Open var_cadena, cnn_distribucion, adOpenDynamic, adLockOptimistic
               '   rsaux1.Open var_cadena, cnn_distribucion, adOpenDynamic, adLockOptimistic
               'End If
               'rs.Close
               'cnn.CommitTrans
               'cnn_distribucion.CommitTrans
            End If
            pro_actualiza_ListView
            txt_cliente.Enabled = False
            txt_registros = lv_clientes.ListItems.Count
            MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
            var_modifica_registro_cliente = True
            var_hubo_cambios = False
            rs.Open "select * from tb_clientes where vcha_tit_titular_id = '" + vartitular + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
            If rs.EOF Then
               var_guardar_cambios = False
               cmd_guardar.Enabled = False
               cmd_deshacer.Enabled = False
               cmd_eliminar.Enabled = False
            Else
               cmd_guardar.Enabled = True
               cmd_deshacer.Enabled = True
               cmd_eliminar.Enabled = True
               var_guardar_cambios = False
            End If
            rs.Close
         Else
            MsgBox "Imposible realizar la acción solicitada", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "Clave de cliente ya existe", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_guardar_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub cmd_imprimir_Click()
   Set reporte = appl.OpenReport(App.Path + "\rep_catalogo_clientes.rpt")
   frmvistasprevias.cr.ReportSource = reporte
   For ntablas = 1 To reporte.Database.Tables.Count
       reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   Next ntablas
   frmvistasprevias.cr.ViewReport
   frmvistasprevias.Caption = "Catálogo de Clientes"
   frmvistasprevias.Show
   Set reporte = Nothing
End Sub

Private Sub cmd_imprimir_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
   txt_fecha_captura = Date
   var_guardar_cambios = True
   txt_nombre_cliente.Enabled = True
   txt_representante.Enabled = True
   txt_fecha_captura.Enabled = True
   txt_agente.Enabled = True
   txt_ruta.Enabled = True
   txt_curp.Enabled = True
   txt_rfc.Enabled = True
   txt_moneda.Enabled = True
   txt_plazo.Enabled = True
   txt_tipo_cliente.Enabled = True
   txt_lista_precios.Enabled = True
   txt_transporte.Enabled = True
   txt_agrupador.Enabled = True
   chk_usa_agrupador.Enabled = True
   chk_estatus.Enabled = True
   txt_prioridad.Enabled = True
   txt_email.Enabled = True
   txt_domicilio.Enabled = True
   txt_codigo_postal.Enabled = True
   chk_enviar_facturas.Enabled = True
   chk_asignacion_catalogos.Enabled = True
   txt_cliente.Enabled = False
   txt_nombre_cliente.SetFocus: var_modifica_registro_cliente = False
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
   txt_fecha_captura = Date
   var_guardar_cambios = True
   txt_nombre_cliente.Enabled = True
   txt_representante.Enabled = True
   txt_fecha_captura.Enabled = True
   txt_agente.Enabled = True
   txt_ruta.Enabled = True
   txt_curp.Enabled = True
   txt_rfc.Enabled = True
   txt_moneda.Enabled = True
   txt_plazo.Enabled = True
   txt_tipo_cliente.Enabled = True
   txt_lista_precios.Enabled = True
   txt_transporte.Enabled = True
   txt_agrupador.Enabled = True
   chk_usa_agrupador.Enabled = True
   chk_estatus.Enabled = True
   txt_prioridad.Enabled = True
   txt_email.Enabled = True
   txt_domicilio.Enabled = True
   txt_codigo_postal.Enabled = True
   chk_enviar_facturas.Enabled = True
   chk_asignacion_catalogos.Enabled = True
   If (UCase(parametros(0)) = "DBPRUEBAS" Or UCase(parametros(0)) = "sqlquezada2") And var_empresa = "31" Then
      Me.txt_agente = "00260"
      rs.Open "SELECT * FROM TB_AGENTES WHERE VCHA_aGE_AGENTE_ID = '00260'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      End If
      rs.Close
      Me.txt_ruta = "0125"
      rs.Open "SELECT * FROM TB_RUTAS WHERE VCHA_RUT_RUTA_ID = '0125'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
      End If
      rs.Close
      Me.txt_moneda = "1"
      rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + Me.txt_moneda + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_moneda = IIf(IsNull(rs!vcha_mon_nombre), "", rs!vcha_mon_nombre)
      End If
      rs.Close
      Me.txt_plazo = "4"
      rs.Open "select * from tb_plazos where vcha_pla_plazo_id = '" + Me.txt_plazo + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_plazo = IIf(IsNull(rs!vcha_pla_nombre), "", rs!vcha_pla_nombre)
      End If
      rs.Close
      Me.txt_tipo_cliente = "IN"
      rs.Open "select * from TB_TIPOSCLIENTES where VCHA_TCL_TIPO_CLIENTE_ID = '" + txt_tipo_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_tipo_cliente = IIf(IsNull(rs!VCHA_TCL_nombre), "", rs!VCHA_TCL_nombre)
      End If
      rs.Close
      Me.txt_lista_precios = "24"
      rs.Open "select * from TB_LISTADEPRECIOS where VCHA_LIS_LISTA_ID = '" + txt_lista_precios + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_lista_precios = IIf(IsNull(rs!VCHA_lIS_NOMBRE), "", rs!VCHA_lIS_NOMBRE)
      End If
      rs.Close
   End If
End Sub

Private Sub cmd_nuevo_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub cmd_pedido_Click()
   If Me.txt_cliente <> "" Then
      var_cliente_pedido_internet = Me.txt_cliente
      frmgenerapedido.Show
   End If
End Sub

Private Sub cmd_pedido_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_cliente = False Then
      var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
      If var_si <> 6 Then
         GoTo salir:
      End If
   Else
      If var_hubo_cambios = True Then
         var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
         If var_si <> 6 Then
            GoTo salir:
         End If
      End If
   End If
   Unload Me
   Exit Sub
salir:
End Sub

Private Sub cmd_salir_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub cmdfecha_Click(Index As Integer)
   If IsDate(Me.txt_fecha_captura) Then
      mes(0).Value = Date
   Else
      mes(0).Value = CDate(Me.txt_fecha_captura)
   End If
   mes(0).Visible = True
   mes(0).SetFocus
End Sub

Private Sub cmdfecha_GotFocus(Index As Integer)
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub Command1_Click()
   If Trim(Me.txt_cliente) <> "" Then
      frmcambiar_titular.txt_tipo = 1
      frmcambiar_titular.txt_clave_cliente = Me.txt_cliente
      frmcambiar_titular.Show 1
      Call pro_limpiatextos(Me)
      lv_clientes.ListItems.Clear
      Call pro_llena_listview1
   Else
      MsgBox "No se a seleccionado un cliente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command1_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub Command2_Click()
   If Me.lv_clientes.ListItems.Count > 0 Then
      var_tipo_datos_adicionales = 0
      var_hubo_cambios = True
      var_nombre_cliente_ad = lv_clientes.selectedItem.SubItems(40)
      var_paterno_cliente_ad = lv_clientes.selectedItem.SubItems(41)
      var_materno_cliente_ad = lv_clientes.selectedItem.SubItems(42)
      var_numero_cliente_ad = lv_clientes.selectedItem.SubItems(43)
      var_clave_tel_pais_ad = lv_clientes.selectedItem.SubItems(44)
      var_clave_tel_estado_ad = lv_clientes.selectedItem.SubItems(45)
      var_calle_cliente_ad = lv_clientes.selectedItem.SubItems(46)
      var_numero_interno_cliente_ad = Me.lv_clientes.selectedItem.SubItems(47)
      frmdatos_adisionales.Show 1
      Me.lv_clientes.selectedItem.SubItems(40) = var_nombre_cliente_ad
      lv_clientes.selectedItem.SubItems(41) = var_paterno_cliente_ad
      lv_clientes.selectedItem.SubItems(42) = var_materno_cliente_ad
      lv_clientes.selectedItem.SubItems(43) = var_numero_cliente_ad
      lv_clientes.selectedItem.SubItems(44) = var_clave_tel_pais_ad
      lv_clientes.selectedItem.SubItems(45) = var_clave_tel_estado_ad
      lv_clientes.selectedItem.SubItems(46) = var_calle_cliente_ad
      lv_clientes.selectedItem.SubItems(47) = var_numero_interno_cliente_ad
   Else
      MsgBox "Debe de dar primero de alta el cliente", vbOKOnly, "ATENCION"
   End If
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 71 Then
      cmd_guardar_Click
   End If
   If Shift = 4 And KeyCode = 68 Then
      cmd_deshacer_Click
   End If
   If Shift = 4 And KeyCode = 69 Then
      cmd_eliminar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 0
   var_barra = ""
   frm_colonias.Visible = False
   frm_lista.Visible = False
   mes(0).Visible = False
   var_guardar_cambios = False
   Me.frm_busqueda_clientes.Visible = False
   If var_tipo_filtrado_cliente = 2 Then
      vartitular = frmclientes2.lv_titulares.selectedItem
   Else
      vartitular = frmlistatitulares.lv_listatitulares.selectedItem
   End If
   var_modifica_registro_cliente = True
   lv_clientes.SmallIcons = ImageList1
   Call pro_encabezadosView(Me, lv_clientes, False)
   rs.Open "select * from tb_clientes where vcha_tit_titular_id =  '" & vartitular & "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
   If rs.BOF Then
      rs.Close
      cmd_guardar.Enabled = False
      cmd_deshacer.Enabled = False
      cmd_eliminar.Enabled = False
   Else
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
      cmd_eliminar.Enabled = True
      rs.Close
      Call pro_llena_listview1
      Call pro_textos
      'lv_clientes.SetFocus
   End If
   If (UCase(parametros(0)) = "DBPRUEBAS" Or UCase(parametros(0)) = "sqlquezada2") And var_empresa = "31" Then
      Me.Command1.Enabled = False
      Me.cmd_pedido.Enabled = True
   Else
      Me.cmd_pedido.Enabled = False
   End If
   pro_textos
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Frmmenu2.StatusBar1.Panels(1) = ""
   Call activa_forma(var_activa_forma_clientes)
End Sub

Private Sub Label8_Click()

End Sub

Private Sub lv_busqueda_clientes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_busqueda_clientes.ListItems.Count > 0 Then
         valor = Me.lv_busqueda_clientes.selectedItem
         Set itmfound = Me.lv_clientes.findItem(valor, lvwText, , lvwPartial)
         itmfound.EnsureVisible
         itmfound.Selected = True
         Me.lv_clientes.SetFocus
         Me.frm_busqueda_clientes.Visible = False
      Else
         MsgBox "No se selecciono ningún cliente", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      Me.txt_nombre_cliente.SetFocus
   End If
End Sub

Private Sub lv_clientes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_clientes, ColumnHeader)
End Sub

Private Sub lv_clientes_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Call pro_textos
End Sub

Private Sub lv_clientes_ItemClick(ByVal item As MSComctlLib.ListItem)
   Set lv_clientes.selectedItem = item
   Call pro_textos
   var_modifica_registro_cliente = True
   txt_cliente.Enabled = False
End Sub

Private Sub lv_colonias_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_colonias, ColumnHeader)
End Sub

Private Sub lv_colonias_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_colonias.ListItems.Count > 0 Then
         txt_colonia = lv_colonias.selectedItem
         txt_nombre_colonia = lv_colonias.selectedItem.SubItems(1)
         txt_pais = lv_colonias.selectedItem.SubItems(2)
         txt_nombre_pais = lv_colonias.selectedItem.SubItems(3)
         txt_estado = lv_colonias.selectedItem.SubItems(4)
         txt_nombre_estado = lv_colonias.selectedItem.SubItems(5)
         txt_municipio = lv_colonias.selectedItem.SubItems(6)
         txt_nombre_municipio = lv_colonias.selectedItem.SubItems(7)
         txt_ciudad = lv_colonias.selectedItem.SubItems(8)
         txt_nombre_ciudad = lv_colonias.selectedItem.SubItems(9)
      Else
         txt_colonia = ""
         txt_nombre_colonia = ""
         txt_pais = ""
         txt_nombre_pais = ""
         txt_estado = ""
         txt_nombre_estado = ""
         txt_municipio = ""
         txt_nombre_municipio = ""
         txt_ciudad = ""
         txt_nombre_ciudad = ""
      End If
      frm_colonias.Visible = False
      Me.txt_email.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_colonias.Visible = False
   End If
End Sub

Private Sub lv_colonias_LostFocus()
   frm_colonias.Visible = False
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_agente = lv_lista.selectedItem
            txt_nombre_agente = lv_lista.selectedItem.SubItems(1)
         Else
            txt_agente = ""
            txt_nombre_agente = ""
         End If
         txt_agente.SetFocus
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_ruta = lv_lista.selectedItem
            txt_nombre_ruta = lv_lista.selectedItem.SubItems(1)
         Else
            txt_ruta = ""
            txt_nombre_ruta = ""
         End If
         txt_ruta.SetFocus
      End If
      If var_tipo_lista = 3 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_moneda = lv_lista.selectedItem
            txt_nombre_moneda = lv_lista.selectedItem.SubItems(1)
         Else
            txt_moneda = ""
            txt_nombre_moneda = ""
         End If
         txt_moneda.SetFocus
      End If
      If var_tipo_lista = 4 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_plazo = lv_lista.selectedItem
            txt_nombre_plazo = lv_lista.selectedItem.SubItems(1)
         Else
            txt_plazo = ""
            txt_nombre_plazo = ""
         End If
         txt_plazo.SetFocus
      End If
      If var_tipo_lista = 5 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_tipo_cliente = lv_lista.selectedItem
            txt_nombre_tipo_cliente = lv_lista.selectedItem.SubItems(1)
         Else
            txt_tipo_cliente = ""
            txt_nombre_tipo_cliente = ""
         End If
         txt_tipo_cliente.SetFocus
      End If
      If var_tipo_lista = 6 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_lista_precios = lv_lista.selectedItem
            txt_nombre_lista_precios = lv_lista.selectedItem.SubItems(1)
         Else
            txt_lista_precios = ""
            txt_nombre_lista_precios = ""
         End If
         txt_lista_precios.SetFocus
      End If
      If var_tipo_lista = 7 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_canal_venta = lv_lista.selectedItem
            txt_nombre_canal_venta = lv_lista.selectedItem.SubItems(1)
         Else
            txt_canal_venta = ""
            txt_nombre_canal_venta = ""
         End If
         txt_canal_venta.SetFocus
      End If
      If var_tipo_lista = 8 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_transporte = lv_lista.selectedItem
            txt_nombre_transporte = lv_lista.selectedItem.SubItems(1)
         Else
           txt_transporte = ""
           txt_nombre_transporte = ""
         End If
         txt_transporte.SetFocus
      End If
      If var_tipo_lista = 9 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_agrupador = lv_lista.selectedItem
            txt_nombre_agrupador = lv_lista.selectedItem.SubItems(1)
         Else
            txt_agrupador = ""
            txt_nombre_agrupador = ""
         End If
         txt_agrupador.SetFocus
      End If
      If var_tipo_lista = 10 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_ciudad = lv_lista.selectedItem
            txt_nombre_ciudad = lv_lista.selectedItem.SubItems(1)
         Else
            txt_ciudad = ""
            txt_nombre_ciudad = ""
         End If
         txt_ciudad.SetFocus
      End If
      If var_tipo_lista = 11 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_municipio = lv_lista.selectedItem
            txt_nombre_municipio = lv_lista.selectedItem.SubItems(1)
         Else
            txt_municipio = ""
            txt_nombre_municipio = ""
         End If
         txt_municipio.SetFocus
      End If
      If var_tipo_lista = 12 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_estado = lv_lista.selectedItem
            txt_nombre_estado = lv_lista.selectedItem.SubItems(1)
         Else
            txt_estado = ""
            txt_nombre_estado = ""
         End If
         txt_estado.SetFocus
      End If
      If var_tipo_lista = 13 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_pais = lv_lista.selectedItem
            txt_nombre_pais = lv_lista.selectedItem.SubItems(1)
         Else
            txt_pais = ""
            txt_nombre_pais = ""
         End If
         txt_pais.SetFocus
      End If
      If var_tipo_lista = 14 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_colonia = lv_lista.selectedItem
            txt_nombre_colonia = lv_lista.selectedItem.SubItems(1)
         Else
            txt_colonia = ""
            txt_nombre_colonia = ""
         End If
         txt_colonia.SetFocus
      End If
      If var_tipo_lista = 16 Then
         If lv_lista.ListItems.Count > 0 Then
            rs.Open "select * from vw_colonias where vcha_col_cp = '" + Me.txt_codigo_postal + "' AND VCHA_PAI_PAIS_ID = '" + lv_lista.selectedItem + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
            If rs.RecordCount = 1 Then
               txt_colonia = IIf(IsNull(rs!VCHA_COL_COLONIA_ID), "", rs!VCHA_COL_COLONIA_ID)
               txt_nombre_colonia = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
               txt_pais = IIf(IsNull(rs!VCHA_PAI_PAIS_ID), "", rs!VCHA_PAI_PAIS_ID)
               txt_nombre_pais = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
               txt_estado = IIf(IsNull(rs!VCHA_EST_ESTADO_ID), "", rs!VCHA_EST_ESTADO_ID)
               txt_nombre_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
               txt_municipio = IIf(IsNull(rs!VCHA_MUN_MUNICIPIO_ID), "", rs!VCHA_MUN_MUNICIPIO_ID)
               txt_nombre_municipio = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
               txt_ciudad = IIf(IsNull(rs!VCHA_CIU_CIUDAD_ID), "", rs!VCHA_CIU_CIUDAD_ID)
               txt_nombre_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
               rs.Close
               txt_email.SetFocus
               frm_lista.Visible = False
            Else
               'rs.MoveFirst
               If Not rs.EOF Then
                  lv_colonias.ListItems.Clear
                  While Not rs.EOF
                        Set list_item = lv_colonias.ListItems.Add(, , rs!VCHA_COL_COLONIA_ID)
                        list_item.SubItems(1) = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
                        list_item.SubItems(2) = IIf(IsNull(rs!VCHA_PAI_PAIS_ID), "", rs!VCHA_PAI_PAIS_ID)
                        list_item.SubItems(3) = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
                        list_item.SubItems(4) = IIf(IsNull(rs!VCHA_EST_ESTADO_ID), "", rs!VCHA_EST_ESTADO_ID)
                        list_item.SubItems(5) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                        list_item.SubItems(6) = IIf(IsNull(rs!VCHA_MUN_MUNICIPIO_ID), "", rs!VCHA_MUN_MUNICIPIO_ID)
                        list_item.SubItems(7) = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
                        list_item.SubItems(8) = IIf(IsNull(rs!VCHA_CIU_CIUDAD_ID), "", rs!VCHA_CIU_CIUDAD_ID)
                        list_item.SubItems(9) = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                        rs.MoveNext
                 Wend
                  lbl_colonias = "COLONIAS DEL C.P. " + txt_codigo_postal
                  var_n = lv_colonias.ListItems.Count
                  If var_n > 6 Then
                     lv_colonias.ColumnHeaders(2).Width = 4270.71
                  Else
                     lv_colonias.ColumnHeaders(2).Width = 4499.71
                  End If
                  frm_colonias.Visible = True
                  lv_colonias.SetFocus
               Else
                  MsgBox "Código postal incorrecto", vbOKOnly, "ATENCION"
               End If
               rs.Close
            End If
         Else
            txt_colonia = ""
            txt_nombre_colonia = ""
            txt_pais = ""
            txt_nombre_pais = ""
            txt_estado = ""
            txt_nombre_estado = ""
            txt_municipio = ""
            txt_nombre_municipio = ""
            txt_ciudad = ""
            txt_nombre_ciudad = ""
         End If
      End If
      frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      If var_tipo_lista = 1 Then
         txt_agente.SetFocus
      End If
      If var_tipo_lista = 2 Then
         txt_ruta.SetFocus
      End If
      If var_tipo_lista = 3 Then
         txt_moneda.SetFocus
      End If
      If var_tipo_lista = 4 Then
         txt_plazo.SetFocus
      End If
      If var_tipo_lista = 5 Then
         txt_tipo_cliente.SetFocus
      End If
      If var_tipo_lista = 6 Then
         txt_lista_precios.SetFocus
      End If
      If var_tipo_lista = 7 Then
         txt_canal_venta.SetFocus
      End If
      If var_tipo_lista = 8 Then
         txt_transporte.SetFocus
      End If
      If var_tipo_lista = 9 Then
         txt_agrupador.SetFocus
      End If
      If var_tipo_lista = 10 Then
         txt_ciudad.SetFocus
      End If
      If var_tipo_lista = 11 Then
         txt_municipio.SetFocus
      End If
      If var_tipo_lista = 12 Then
         txt_estado.SetFocus
      End If
      If var_tipo_lista = 13 Then
         txt_pais.SetFocus
      End If
      If var_tipo_lista = 14 Then
         txt_colonia.SetFocus
      End If
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_agente_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_agente_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_agrupador_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_agrupador_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub


Private Sub txt_anterior_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_anterior_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub txt_anterior_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txt_domicilio.SetFocus
    End If
End Sub

Private Sub txt_buscar_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub txt_buscar_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.frm_busqueda_clientes.Visible = True
      Me.txt_busqueda_cliente = ""
      Me.lv_busqueda_clientes.ListItems.Clear
      Me.txt_busqueda_cliente.SetFocus
   End If
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_clientes, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_busqueda_cliente_Change()
   Me.lv_busqueda_clientes.ListItems.Clear
End Sub

Private Sub txt_busqueda_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.lv_busqueda_clientes.ListItems.Clear
      rs.Open "SELECT * FROM VW_CLIENTES WHERE VCHA_CLI_NOMBRE LIKE '%" + Me.txt_busqueda_cliente + "%' and vcha_tit_titular_id = '" + vartitular + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = Me.lv_busqueda_clientes.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      If Me.lv_busqueda_clientes.ListItems.Count > 0 Then
         Me.lv_busqueda_clientes.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.txt_nombre_cliente.SetFocus
   End If
End Sub

Private Sub txt_ciudad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_ciudades where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_ciu_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CIU_CIUDAD_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CIUDADES DE " + txt_nombre_estado
      var_tipo_lista = 10
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
   If KeyCode = 117 Then
      frmciudades.Show
   End If
End Sub

Private Sub txt_clave_unificada_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub txt_cliente_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub txt_codigo_postal_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_codigo_postal_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_codigo_postal_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_activa_forma_direcciones = Me.Name
      frmclientes.Enabled = False
      frmdirecciones.Show
      If var_aceptar_direccion = True Then
         txt_pais = var_dir_pais
         txt_nombre_pais = var_dir_nombre_pais
         txt_estado = var_dir_estado
         txt_nombre_estado = var_dir_nombre_estado
         txt_municipio = var_dir_municipio
         txt_nombre_municipio = var_dir_nombre_municipio
         txt_ciudad = var_dir_ciudad
         txt_nombre_ciudad = var_dir_nombre_ciudad
         txt_colonia = var_dir_colonia
         txt_nombre_colonia = var_dir_nombre_colonia
         txt_codigo_postal = var_dir_codigo_postal
      End If
   End If
End Sub

Private Sub txt_codigo_postal_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_colonia_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_colonias where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_col_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_COL_COLONIA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "COLONIAS DE " + txt_nombre_estado
      var_tipo_lista = 14
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
   If KeyCode = 117 Then
      frmciudades.Show
   End If
End Sub

Private Sub txt_curp_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_curp_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_domicilio_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_domicilio_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_email_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_email_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_estado_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_estados where vcha_pai_pais_id = '" + txt_pais + "' order by vcha_est_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_EST_ESTADO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ESTADOS DE " + txt_nombre_pais
      var_tipo_lista = 12
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
   If KeyCode = 117 Then
      frmestados.Show
   End If
End Sub

Private Sub txt_fecha_captura_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_fecha_captura_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_lista_precios_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_lista_precios_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_moneda_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_moneda_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_municipio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_municipios where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_mun_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_MUN_MUNICIPIO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MUNICIPIOS DE" + txt_nombre_estado
      var_tipo_lista = 11
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
   If KeyCode = 117 Then
      frmmunicipios.Show
   End If
End Sub

Private Sub txt_municipio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If

End Sub

Private Sub txt_nombre_agente_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_agente_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_agente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_agrupador_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_agrupador_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_agrupador_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub



Private Sub txt_nombre_ciudad_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_ciudad_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub txt_nombre_ciudad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_ciudades where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_ciu_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CIU_CIUDAD_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CIUDADES DE " + txt_nombre_estado
      var_tipo_lista = 10
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
   If KeyCode = 117 Then
      frmciudades.Show
   End If
End Sub

Private Sub txt_nombre_cliente_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_cliente_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.frm_busqueda_clientes.Visible = True
      Me.txt_busqueda_cliente = ""
      Me.lv_busqueda_clientes.ListItems.Clear
      Me.txt_busqueda_cliente.SetFocus
   End If
End Sub

Private Sub txt_nombre_colonia_Change()
   Me.frm_busqueda_clientes.Visible = False
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_colonia_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_colonias where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_col_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_COL_COLONIA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_COLONIA_NOMBRE), "", rs!vcha_col_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "COLONIAS DE " + txt_nombre_estado
      var_tipo_lista = 14
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
   If KeyCode = 117 Then
      frmciudades.Show
   End If
End Sub

Private Sub txt_nombre_estado_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_estado_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub txt_nombre_estado_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_estados where vcha_pai_pais_id = '" + txt_pais + "' order by vcha_est_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_EST_ESTADO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ESTADOS DE " + txt_nombre_pais
      var_tipo_lista = 12
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
   If KeyCode = 117 Then
      frmestados.Show
   End If
End Sub

Private Sub txt_nombre_lista_precios_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_lista_precios_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_lista_precios_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_moneda_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_moneda_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_moneda_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_municipio_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_municipio_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub txt_nombre_municipio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_municipios where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_mun_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_MUN_MUNICIPIO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MUNICIPIOS DE" + txt_nombre_estado
      var_tipo_lista = 3
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
   If KeyCode = 117 Then
      frmmunicipios.Show
   End If
End Sub

Private Sub txt_nombre_pais_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_pais_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub txt_nombre_pais_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_paises order by vcha_pai_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_PAI_PAIS_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PAISES"
      var_tipo_lista = 13
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
   If KeyCode = 117 Then
      var_catalogo_articulos = True
      frmpaises.Show
   End If
End Sub

Private Sub txt_nombre_plazo_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_plazo_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_plazo_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_ruta_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_ruta_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_ruta_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_tipo_cliente_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_tipo_cliente_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_tipo_cliente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_transporte_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_transporte_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_transporte_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_pais_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_paises order by vcha_pai_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_PAI_PAIS_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PAISES"
      var_tipo_lista = 13
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
   If KeyCode = 117 Then
      var_catalogo_articulos = True
      frmpaises.Show
   End If
End Sub

Private Sub mes_DblClick(Index As Integer)
   txt_fecha_captura = mes(0).Value
   mes(0).Visible = False
   txt_agente.SetFocus
End Sub

Private Sub mes_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 27 Then
      mes(0).Visible = False
      txt_agente.SetFocus
   End If
End Sub

Private Sub mes_LostFocus(Index As Integer)
   mes(0).Visible = False
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_clientes.SetFocus
      Call pro_avanzar(Me, lv_clientes, Button)
      lv_clientes.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_clientes.ListItems(1).Selected = True
      lv_clientes.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_clientes = lv_clientes.ListItems.Count
      lv_clientes.ListItems(numero_items_clientes).Selected = True
      lv_clientes.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_clientes()
   Dim ok As Boolean
   Set TB_CLIENTES = New TB_CLIENTES
   Set TB_BITACORA_CLIENTES = New TB_BITACORA_CLIENTES
   If txt_nombre_cliente <> "" Then
      If var_hubo_cambios Then
         rs.Open "select * from tb_clientes where VCHA_CLI_CLAVE_ID = '" + txt_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
         var_cliente_regreso = txt_cliente
         ok = TB_CLIENTES.Anadir(txt_cliente, txt_nombre_cliente, txt_representante, txt_fecha_captura, txt_agente, txt_ruta, txt_curp, txt_rfc, txt_moneda, txt_plazo, txt_tipo_cliente, txt_lista_precios, txt_transporte, txt_agrupador, chk_usa_agrupador, chk_estatus, vartitular, txt_prioridad, txt_email, txt_pais, txt_estado, txt_ciudad, txt_colonia, txt_domicilio, txt_codigo_postal, txt_municipio, chk_enviar_facturas, chk_asignacion_catalogos, txt_anterior)
         If Trim(var_cliente_regreso) <> "" Then
            txt_cliente = var_cliente_regreso
         End If
         If Trim(Me.txt_anterior) = "" Then
            Me.txt_anterior = txt_cliente
         End If
         If lv_clientes.ListItems.Count > 0 Then
            var_nombre_cliente_ad = lv_clientes.selectedItem.SubItems(40)
            var_paterno_cliente_ad = lv_clientes.selectedItem.SubItems(41)
            var_materno_cliente_ad = lv_clientes.selectedItem.SubItems(42)
            var_numero_cliente_ad = lv_clientes.selectedItem.SubItems(43)
            var_clave_tel_pais_ad = lv_clientes.selectedItem.SubItems(44)
            var_clave_tel_estado_ad = lv_clientes.selectedItem.SubItems(45)
            var_calle_cliente_ad = lv_clientes.selectedItem.SubItems(46)
            var_numero_interno_cliente_ad = lv_clientes.selectedItem.SubItems(47)
         Else
            var_nombre_cliente_ad = ""
            var_paterno_cliente_ad = ""
            var_materno_cliente_ad = ""
            var_numero_cliente_ad = ""
            var_clave_tel_pais_ad = ""
            var_clave_tel_estado_ad = ""
            var_calle_cliente_ad = ""
            var_numero_interno_cliente_ad = ""
      End If
         var_cadena = "UPDATE TB_CLIENTES SET vcha_cli_numero_interno = '" + var_numero_interno_cliente_ad + "', inte_cli_clave_tel_estado = '" + var_clave_tel_estado_ad + "', inte_cli_clave_Tel_pais = '" + var_clave_tel_pais_ad + "', VCHA_CLI_CALLE = '" + var_calle_cliente_ad + "', INTE_CLI_NUMERO = '" + CStr(var_numero_cliente_ad) + "', VCHA_CLI_MATERNO = '" + var_materno_cliente_ad + "', vcha_cli_paterno  = '" + var_paterno_cliente_ad + "', vcha_cli_nombre_2 = '" + var_nombre_cliente_ad + "', VCHA_CLI_CLAVE_ANTERIOR_ID = '" + Me.txt_anterior + "', VCHA_EMP_EMPRESA_ID = '" + var_empresa + "', inte_cli_cliente_pedido_tienda = " + CStr(Me.chk_pedido) + ", inte_cli_franquicia = " + CStr(Me.chk_franquicia) + ", vcha_cli_telefono = '" + Me.txt_telefono + "', inte_cli_Trazabilidad = " + CStr(Me.chk_trazabilidad) + ", vcha_cli_referencia = '" + Me.txt_referencia + "', inte_cli_activo = " + CStr(Me.chk_activo) + ", INTE_CLI_PUBLICO_GENERAL = " + CStr(Me.chk_venta_publico_general)
         var_cadena = var_cadena + ", inte_cli_promocion = " + CStr(Me.chk_promocion) + " where vcha_cli_clave_id = '" + txt_cliente + "'"
         rsaux.Open var_cadena, cnn_distribucion, adOpenDynamic, adLockOptimistic
         If var_modifica_registro_cliente = False Then
            rsaux.Open "EXEC SP_CALCULO_REFERENCIA_CLIENTES '" + txt_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
         End If
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + txt_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            txt_referencia = IIf(IsNull(rsaux!VCHA_CLI_REFERENCIA), "", rsaux!VCHA_CLI_REFERENCIA)
         End If
         rsaux.Close
         If ok Then
            bitacora = True
            If var_modifica_registro_cliente = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_CLIENTES.Anadir(txt_cliente, "VCHA_TAL_NOMBRE", "", txt_nombre_cliente, var_clave_usuario_global, fun_NombrePc, Date)
               rsaux.Open "UPDATE TB_CLIENTES SET NUM_INTER_TRANC_TYPE = 1, NUM_INTER_UPLOADED = 1 where vcha_cli_clave_id = '" + txt_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
            Else
               var_operacion_bitacora = "M"
               rsaux.Open "UPDATE TB_CLIENTES SET NUM_INTER_TRANC_TYPE = 2, NUM_INTER_UPLOADED = 2, VCHA_CLI_REFERENCIA = '" + Me.txt_referencia + "' where vcha_cli_clave_id = '" + txt_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  If rs!vcha_cli_clave_id <> txt_cliente Then
                        bitacora = TB_BITACORA_CLIENTES.Anadir(txt_cliente, "VCHA_CLI_CLIENTE_ID", rs!vcha_cli_clave_id, txt_cliente, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!VCHA_CLI_NOMBRE <> txt_nombre_cliente Then
                     bitacora = TB_BITACORA_CLIENTES.Anadir(txt_cliente, "VCHA_CLI_NOMBRE", rs!VCHA_CLI_NOMBRE, txt_nombre_cliente, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!vcha_cli_representante <> txt_representante Then
                     bitacora = TB_BITACORA_CLIENTES.Anadir(txt_cliente, "VCHA_CLI_REPRESENTANTE", rs!vcha_cli_representante, txt_representante, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!dtim_cli_fecha_Captura <> txt_fecha_captura Then
                     bitacora = TB_BITACORA_CLIENTES.Anadir(txt_cliente, "VCHA_CLI_FECHA_FACTURA", rs!dtim_cli_fecha_Captura, txt_fecha_captura, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!VCHA_AGE_AGENTE_ID <> txt_agente Then
                     bitacora = TB_BITACORA_CLIENTES.Anadir(txt_cliente, "VCHA_VEN_VENDEDOR_ID", rs!VCHA_AGE_AGENTE_ID, txt_agente, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!vcha_rut_ruta_id <> txt_ruta Then
                     bitacora = TB_BITACORA_CLIENTES.Anadir(txt_cliente, "VCHA_RUT_RUTA_ID", rs!vcha_rut_ruta_id, txt_ruta, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!VCHA_CLI_CURP <> txt_curp Then
                     bitacora = TB_BITACORA_CLIENTES.Anadir(txt_cliente, "VCHA_CLI_CURP", rs!VCHA_CLI_CURP, txt_curp, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!VCHA_CLI_RFC <> txt_rfc Then
                     bitacora = TB_BITACORA_CLIENTES.Anadir(txt_cliente, "VCHA_CLI_RFC", rs!VCHA_CLI_RFC, txt_rfc, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!vcha_mon_moneda_id <> txt_moneda Then
                     bitacora = TB_BITACORA_CLIENTES.Anadir(txt_cliente, "VCHA_MON_MONEDA_ID", rs!vcha_mon_moneda_id, txt_moneda, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!VCHA_PLA_PLAZO_ID <> txt_plazo Then
                     bitacora = TB_BITACORA_CLIENTES.Anadir(txt_cliente, "VCHA_PLA_PLAZO_ID", rs!VCHA_PLA_PLAZO_ID, txt_plazo, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!VCHA_TCL_TIPO_CLIENTE_ID <> txt_tipo_cliente Then
                     bitacora = TB_BITACORA_CLIENTES.Anadir(txt_cliente, "VCHA_TCL_TIPO_CLIENTE_ID", rs!VCHA_TCL_TIPO_CLIENTE_ID, txt_tipo_cliente, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!vcha_LIS_LISTA_iD <> txt_lista_precios Then
                     bitacora = TB_BITACORA_CLIENTES.Anadir(txt_cliente, "VCHA_LIS_LISTA_ID", rs!vcha_LIS_LISTA_iD, txt_lista_precios, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!VCHA_TRA_TRANSPORTE_ID <> txt_transporte Then
                     bitacora = TB_BITACORA_CLIENTES.Anadir(txt_cliente, "VCHA_TRN_TRANSPORTE_ID", rs!VCHA_TRA_TRANSPORTE_ID, txt_transporte, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!VCHA_FAG_FAMILIA_AGRUPADOR_ID <> txt_agrupador Then
                     bitacora = TB_BITACORA_CLIENTES.Anadir(txt_cliente, "VCHA_FAG_FAMILIA_AGRUPADOR_ID", rs!VCHA_FAG_FAMILIA_AGRUPADOR_ID, txt_agrupador, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!VCHA_CLI_REFERENCIA <> Me.txt_referencia Then
                     bitacora = TB_BITACORA_CLIENTES.Anadir(txt_cliente, "VCHA_CLI_REFEERENCIA", rs!VCHA_CLI_REFERENCIA, Me.txt_referencia, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
               End If
            End If
            rs.Close
         Else
            MsgBox "No se puede grabar registro: " + TB_CLIENTES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
  Set TB_CLIENTES = Nothing
End Sub

Sub pro_elimina_clientes()
Dim var_llave_usuarios As String

Set TB_CLIENTES = New TB_CLIENTES
Set TB_BITACORA_CLIENTES = New TB_BITACORA_CLIENTES
On Error GoTo salir
  
    ok = True
        If txt_cliente <> "" And txt_nombre_cliente <> "" Then
            If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
                ok = TB_CLIENTES.Eliminar(txt_cliente)
            Else
                GoTo salir:
            End If
            If ok Then
                numero_items_clientes = numero_items_clientes - 1
                var_operacion_bitacora = "E"
                bitacora = TB_BITACORA_CLIENTES.Anadir(txt_cliente, "VCHA_CLI_NOMBRE", "", txt_nombre_cliente, var_clave_usuario_global, fun_NombrePc, Date)
           Else
                MsgBox "No se puede grabar registro: " + TB_CLIENTES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
salir:
Set TB_CLIENTES = Nothing

End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   numero_items_clientes = 0
   
   rs.Open "select * from VW_clientes where vcha_tit_titular_id = '" & vartitular & "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
      Set list_item = lv_clientes.ListItems.Add(, , rs!vcha_cli_clave_id)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
      list_item.SubItems(2) = IIf(IsNull(rs!vcha_cli_representante), "", rs!vcha_cli_representante)
      list_item.SubItems(3) = IIf(IsNull(rs!dtim_cli_fecha_Captura), "", rs!dtim_cli_fecha_Captura)
      list_item.SubItems(4) = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
      list_item.SubItems(5) = IIf(IsNull(rs!vcha_rut_ruta_id), "", rs!vcha_rut_ruta_id)
      list_item.SubItems(6) = IIf(IsNull(rs!VCHA_CLI_CURP), "", rs!VCHA_CLI_CURP)
      list_item.SubItems(7) = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
      list_item.SubItems(8) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
      list_item.SubItems(9) = IIf(IsNull(rs!VCHA_PLA_PLAZO_ID), "", rs!VCHA_PLA_PLAZO_ID)
      list_item.SubItems(10) = IIf(IsNull(rs!VCHA_TCL_TIPO_CLIENTE_ID), "", rs!VCHA_TCL_TIPO_CLIENTE_ID)
      list_item.SubItems(11) = IIf(IsNull(rs!vcha_LIS_LISTA_iD), "", rs!vcha_LIS_LISTA_iD)
      list_item.SubItems(12) = ""
      list_item.SubItems(13) = IIf(IsNull(rs!VCHA_TRA_TRANSPORTE_ID), "", rs!VCHA_TRA_TRANSPORTE_ID)
      list_item.SubItems(14) = IIf(IsNull(rs!VCHA_FAG_FAMILIA_AGRUPADOR_ID), "", rs!VCHA_FAG_FAMILIA_AGRUPADOR_ID)
      list_item.SubItems(15) = IIf(IsNull(rs!INTE_CLI_AGRUPADOR), 0, rs!INTE_CLI_AGRUPADOR)
      list_item.SubItems(16) = IIf(IsNull(rs!INTE_CLI_ESTATUS), 0, rs!INTE_CLI_ESTATUS)
      list_item.SubItems(17) = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
      list_item.SubItems(18) = IIf(IsNull(rs!CHAR_PRI_PRIORIDAD_ID), "", rs!CHAR_PRI_PRIORIDAD_ID)
      list_item.SubItems(19) = IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
      list_item.SubItems(20) = IIf(IsNull(rs!VCHA_PAI_PAIS_ID), "", rs!VCHA_PAI_PAIS_ID)
      list_item.SubItems(21) = IIf(IsNull(rs!VCHA_EST_ESTADO_ID), "", rs!VCHA_EST_ESTADO_ID)
      list_item.SubItems(22) = IIf(IsNull(rs!VCHA_CIU_CIUDAD_ID), "", rs!VCHA_CIU_CIUDAD_ID)
      list_item.SubItems(23) = IIf(IsNull(rs!VCHA_CLI_COLONIA), "", rs!VCHA_CLI_COLONIA)
      list_item.SubItems(24) = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION)
      list_item.SubItems(25) = IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
      list_item.SubItems(26) = IIf(IsNull(rs!VCHA_MUN_MUNICIPIO_ID), "", rs!VCHA_MUN_MUNICIPIO_ID)
      list_item.SubItems(27) = IIf(IsNull(rs!INTE_CLI_ENVIO_FACTURA), 0, rs!INTE_CLI_ENVIO_FACTURA)
      list_item.SubItems(28) = IIf(IsNull(rs!INTE_CLI_ASIGNACION_CATALOGOS), 0, rs!INTE_CLI_ASIGNACION_CATALOGOS)
      list_item.SubItems(29) = IIf(IsNull(rs!vcha_cli_clave_anterior_id), 0, rs!vcha_cli_clave_anterior_id)
      list_item.SubItems(30) = IIf(IsNull(rs!INTE_CLI_CLIENTE_PEDIDO_TIENDA), 0, rs!INTE_CLI_CLIENTE_PEDIDO_TIENDA)
      list_item.SubItems(31) = IIf(IsNull(rs!VCHA_CLI_REFERENCIA), "", rs!VCHA_CLI_REFERENCIA)
      list_item.SubItems(32) = IIf(IsNull(rs!INTE_CLI_FRANQUICIA), 0, rs!INTE_CLI_FRANQUICIA)
      list_item.SubItems(33) = IIf(IsNull(rs!vcha_cli_telefono), "", rs!vcha_cli_telefono)
      list_item.SubItems(34) = IIf(IsNull(rs!INTE_CLI_TRAZABILIDAD), "", rs!INTE_CLI_TRAZABILIDAD)
      list_item.SubItems(35) = IIf(IsNull(rs!inte_cli_activo), 0, rs!inte_cli_activo)
      list_item.SubItems(36) = IIf(IsNull(rs!vcha_cli_clave_unificada_id), "", rs!vcha_cli_clave_unificada_id)
      list_item.SubItems(37) = IIf(IsNull(rs!inte_cli_unificador), 0, rs!inte_cli_unificador)
      list_item.SubItems(38) = IIf(IsNull(rs!INTE_CLI_PUBLICO_GENERAL), 0, rs!INTE_CLI_PUBLICO_GENERAL)
      list_item.SubItems(39) = IIf(IsNull(rs!inte_cli_promocion), 0, rs!inte_cli_promocion)
      list_item.SubItems(40) = IIf(IsNull(rs!vcha_cli_nombre_2), "", rs!vcha_cli_nombre_2)
      list_item.SubItems(41) = IIf(IsNull(rs!vcha_cli_paterno), "", rs!vcha_cli_paterno)
      list_item.SubItems(42) = IIf(IsNull(rs!VCHA_CLI_MATERNO), "", rs!VCHA_CLI_MATERNO)
      list_item.SubItems(43) = IIf(IsNull(rs!INTE_CLI_NUMERO), "", rs!INTE_CLI_NUMERO)
      list_item.SubItems(44) = IIf(IsNull(rs!INTE_CLI_CLAVE_TEL_PAIS), "", rs!INTE_CLI_CLAVE_TEL_PAIS)
      list_item.SubItems(45) = IIf(IsNull(rs!INTE_CLI_CLAVE_TEL_ESTADO), "", rs!INTE_CLI_CLAVE_TEL_ESTADO)
      list_item.SubItems(46) = IIf(IsNull(rs!VCHA_CLI_CALLE), "", rs!VCHA_CLI_CALLE)
      list_item.SubItems(47) = IIf(IsNull(rs!VCHA_CLI_NUMERO_INTERNO), "", rs!VCHA_CLI_NUMERO_INTERNO)
      numero_items_clientes = numero_items_clientes + 1
      rs.MoveNext:
   Wend
   rs.Close
End Sub


Sub pro_textos()
'On Error GoTo err0:
Dim var_n As Double
   var_n = lv_clientes.ListItems.Count
   If var_n > 0 Then
      txt_cliente = lv_clientes.selectedItem
      txt_nombre_cliente = lv_clientes.selectedItem.SubItems(1)
      txt_representante = lv_clientes.selectedItem.SubItems(2)
      txt_fecha_captura = lv_clientes.selectedItem.SubItems(3)
      txt_agente = lv_clientes.selectedItem.SubItems(4)
      txt_ruta = lv_clientes.selectedItem.SubItems(5)
      txt_curp = lv_clientes.selectedItem.SubItems(6)
      txt_rfc = lv_clientes.selectedItem.SubItems(7)
      txt_moneda = lv_clientes.selectedItem.SubItems(8)
      txt_plazo = lv_clientes.selectedItem.SubItems(9)
      txt_tipo_cliente = lv_clientes.selectedItem.SubItems(10)
      txt_lista_precios = lv_clientes.selectedItem.SubItems(11)
      txt_canal_venta = lv_clientes.selectedItem.SubItems(12)
      txt_transporte = lv_clientes.selectedItem.SubItems(13)
      txt_agrupador = lv_clientes.selectedItem.SubItems(14)
      chk_usa_agrupador = lv_clientes.selectedItem.SubItems(15)
      chk_estatus = lv_clientes.selectedItem.SubItems(16)
      txt_titular = lv_clientes.selectedItem.SubItems(17)
      txt_prioridad = lv_clientes.selectedItem.SubItems(18)
      txt_email = lv_clientes.selectedItem.SubItems(19)
      txt_pais = lv_clientes.selectedItem.SubItems(20)
      txt_estado = lv_clientes.selectedItem.SubItems(21)
      txt_ciudad = lv_clientes.selectedItem.SubItems(22)
      txt_colonia = lv_clientes.selectedItem.SubItems(23)
      txt_domicilio = lv_clientes.selectedItem.SubItems(24)
      txt_codigo_postal = lv_clientes.selectedItem.SubItems(25)
      txt_municipio = lv_clientes.selectedItem.SubItems(26)
      chk_enviar_facturas = lv_clientes.selectedItem.SubItems(27)
      chk_asignacion_catalogos = lv_clientes.selectedItem.SubItems(28)
      txt_anterior = lv_clientes.selectedItem.SubItems(29)
      Me.chk_pedido = lv_clientes.selectedItem.SubItems(30)
      txt_referencia = lv_clientes.selectedItem.SubItems(31)
      Me.chk_franquicia = lv_clientes.selectedItem.SubItems(32)
      Me.txt_telefono = lv_clientes.selectedItem.SubItems(33)
      Me.chk_activo = lv_clientes.selectedItem.SubItems(35)
      Me.txt_clave_unificada = lv_clientes.selectedItem.SubItems(36)
      Me.txt_unificador = lv_clientes.selectedItem.SubItems(37)
      Me.chk_venta_publico_general = Me.lv_clientes.selectedItem.SubItems(38)
      Me.chk_promocion = Me.lv_clientes.selectedItem.SubItems(39)
      var_nombre_cliente_ad = Me.lv_clientes.selectedItem.SubItems(40)
      var_paterno_cliente_ad = Me.lv_clientes.selectedItem.SubItems(41)
      var_materno_cliente_ad = Me.lv_clientes.selectedItem.SubItems(42)
      var_numero_cliente_ad = Me.lv_clientes.selectedItem.SubItems(43)
      var_clave_tel_pais_ad = Me.lv_clientes.selectedItem.SubItems(44)
      var_clave_tel_estado_ad = Me.lv_clientes.selectedItem.SubItems(45)
      var_calle_cliente_ad = Me.lv_clientes.selectedItem.SubItems(46)
      var_numero_interno_cliente_ad = Me.lv_clientes.selectedItem.SubItems(47)
      Me.cmd_guardar.Enabled = True
      Me.cmd_deshacer.Enabled = True
      cmd_eliminar.Enabled = True
      If chk_usa_agrupador.Value = 1 Then
         txt_agrupador.Enabled = True
      Else
         txt_agrupador.Enabled = False
      End If
           
      rs.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + txt_agente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      Else
         txt_nombre_agente = ""
      End If
      rs.Close
        
      rs.Open "select * from tb_rutas where vcha_rut_ruta_id = '" + txt_ruta + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
      Else
         txt_nombre_ruta = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_MONEDAS WHERE VCHA_MON_MONEDA_ID = '" + txt_moneda + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_moneda = IIf(IsNull(rs!vcha_mon_nombre), "", rs!vcha_mon_nombre)
      Else
         txt_nombre_moneda = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_PLAZOS WHERE VCHA_PLA_PLAZO_ID = '" + txt_plazo + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_plazo = IIf(IsNull(rs!vcha_pla_nombre), "", rs!vcha_pla_nombre)
      Else
         txt_nombre_plazo = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_TIPOSCLIENTES WHERE VCHA_TCL_TIPO_CLIENTE_ID = '" + txt_tipo_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_tipo_cliente = IIf(IsNull(rs!VCHA_TCL_nombre), "", rs!VCHA_TCL_nombre)
      Else
         txt_nombre_tipo_cliente = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_LISTADEPRECIOS WHERE VCHA_LIS_LISTA_ID = '" + txt_lista_precios + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_lista_precios = IIf(IsNull(rs!VCHA_lIS_NOMBRE), "", rs!VCHA_lIS_NOMBRE)
      Else
         txt_nombre_lista_precios = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_TRANSPORTES WHERE VCHA_TRN_TRANSPORTE_ID = '" + txt_transporte + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_transporte = IIf(IsNull(rs!vcha_trn_nombre), "", rs!vcha_trn_nombre)
      Else
         txt_nombre_transporte = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_FAMILIA_AGRUPADORES WHERE VCHA_FAG_FAMILIA_AGRUPADOR_ID ='" + txt_agrupador + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_agrupador = IIf(IsNull(rs!vcha_fag_nombre), "", rs!vcha_fag_nombre)
      Else
         txt_nombre_agrupador = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_PAISES WHERE VCHA_PAI_PAIS_ID = '" + txt_pais + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_pais = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
      Else
         txt_nombre_pais = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_ESTADOS WHERE VCHA_EST_ESTADO_ID = '" + txt_estado + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
      Else
         txt_nombre_estado = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_MUNICIPIOS WHERE VCHA_MUN_MUNICIPIO_ID = '" + txt_municipio + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_municipio = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
      Else
         txt_nombre_municipio = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_CIUDADES WHERE VCHA_CIU_CIUDAD_ID = '" + txt_ciudad + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
      Else
         txt_nombre_ciudad = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_COLONIAS WHERE VCHA_COL_COLONIA_ID = '" + txt_colonia + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_colonia = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
      Else
         txt_nombre_colonia = ""
      End If
      rs.Close
   Else
      txt_cliente.Enabled = False
      txt_nombre_cliente.Enabled = False
      txt_representante.Enabled = False
      txt_fecha_captura.Enabled = False
      txt_agente.Enabled = False
      txt_ruta.Enabled = False
      txt_curp.Enabled = False
      txt_rfc.Enabled = False
      txt_moneda.Enabled = False
      txt_plazo.Enabled = False
      txt_tipo_cliente.Enabled = False
      txt_lista_precios.Enabled = False
      txt_transporte.Enabled = False
      txt_agrupador.Enabled = False
      chk_usa_agrupador.Enabled = False
      chk_estatus.Enabled = False
      txt_prioridad.Enabled = False
      txt_email.Enabled = False
      txt_domicilio.Enabled = False
      txt_codigo_postal.Enabled = False
      chk_enviar_facturas.Enabled = False
      chk_asignacion_catalogos.Enabled = False
      cmd_guardar.Enabled = False
      cmd_eliminar.Enabled = False
      cmd_deshacer.Enabled = False
   End If
   var_hubo_cambios = False
   var_modifica_registro_cliente = True
   var_numero_renglones = lv_clientes.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_clientes.ColumnHeaders(2).Width = 3320
   Else
      lv_clientes.ColumnHeaders(2).Width = 3569
   End If
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_cliente = False Then
        Set list_item = lv_clientes.ListItems.Add(, , txt_cliente)
        list_item.SubItems(1) = txt_nombre_cliente
        list_item.SubItems(2) = txt_representante
        list_item.SubItems(3) = txt_fecha_captura
        list_item.SubItems(4) = txt_agente
        list_item.SubItems(5) = txt_ruta
        list_item.SubItems(6) = txt_curp
        list_item.SubItems(7) = txt_rfc
        list_item.SubItems(8) = txt_moneda
        list_item.SubItems(9) = txt_plazo
        list_item.SubItems(10) = txt_tipo_cliente
        list_item.SubItems(11) = txt_lista_precios
        list_item.SubItems(12) = ""
        list_item.SubItems(13) = txt_transporte
        list_item.SubItems(14) = txt_agrupador
        list_item.SubItems(15) = chk_usa_agrupador
        list_item.SubItems(16) = chk_estatus
        list_item.SubItems(17) = txt_titular
        list_item.SubItems(18) = txt_prioridad
        list_item.SubItems(19) = txt_email
        list_item.SubItems(20) = txt_pais
        list_item.SubItems(21) = txt_estado
        list_item.SubItems(22) = txt_ciudad
        list_item.SubItems(23) = txt_colonia
        list_item.SubItems(24) = txt_domicilio
        list_item.SubItems(25) = txt_codigo_postal
        list_item.SubItems(26) = txt_municipio
        list_item.SubItems(27) = chk_enviar_facturas
        list_item.SubItems(28) = chk_asignacion_catalogos
        list_item.SubItems(29) = Me.txt_anterior
        list_item.SubItems(30) = Me.chk_pedido
        list_item.SubItems(31) = txt_referencia
        list_item.SubItems(32) = Me.chk_franquicia
        list_item.SubItems(33) = Me.txt_telefono
        list_item.SubItems(34) = Me.chk_trazabilidad
        list_item.SubItems(35) = Me.chk_activo
        list_item.SubItems(36) = Me.txt_clave_unificada
        list_item.SubItems(37) = Me.txt_unificador
        list_item.SubItems(38) = Me.chk_venta_publico_general
        list_item.SubItems(39) = Me.chk_promocion
        list_item.SubItems(40) = var_nombre_cliente_ad
        list_item.SubItems(41) = var_paterno_cliente_ad
        list_item.SubItems(42) = var_materno_cliente_ad
        list_item.SubItems(43) = var_numero_cliente_ad
        list_item.SubItems(44) = var_clave_tel_pais_ad
        list_item.SubItems(45) = var_clave_tel_estado_ad
        list_item.SubItems(46) = var_calle_cliente_ad
        list_item.SubItems(47) = var_numero_interno_cliente_ad
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_clientes = numero_items_clientes + 1
    Else
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).Checked = False
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index) = txt_cliente
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(1) = txt_nombre_cliente
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(2) = txt_representante
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(3) = txt_fecha_captura
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(4) = txt_agente
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(5) = txt_ruta
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(6) = txt_curp
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(7) = txt_rfc
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(8) = txt_moneda
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(9) = txt_plazo
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(10) = txt_tipo_cliente
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(11) = txt_lista_precios
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(12) = ""
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(13) = txt_transporte
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(14) = txt_agrupador
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(15) = chk_usa_agrupador
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(16) = chk_estatus
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(17) = txt_titular
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(18) = txt_prioridad
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(19) = txt_email
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(20) = txt_pais
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(21) = txt_estado
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(22) = txt_ciudad
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(23) = txt_colonia
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(24) = txt_domicilio
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(25) = txt_codigo_postal
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(26) = txt_municipio
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(27) = chk_enviar_facturas
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(28) = chk_asignacion_catalogos
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(29) = txt_anterior
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(30) = Me.chk_pedido
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(31) = txt_referencia
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(32) = Me.chk_franquicia
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(33) = Me.txt_telefono
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(34) = Me.chk_trazabilidad
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(35) = Me.chk_activo
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(36) = Me.txt_clave_unificada
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(37) = Me.txt_unificador
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(38) = Me.chk_venta_publico_general
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(39) = Me.chk_promocion
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(40) = var_nombre_cliente_ad
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(41) = var_paterno_cliente_ad
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(42) = var_materno_cliente_ad
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(43) = var_numero_cliente_ad
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(44) = var_clave_tel_pais_ad
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(45) = var_clave_tel_estado_ad
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(46) = var_calle_cliente_ad
        lv_clientes.ListItems.item(lv_clientes.selectedItem.Index).ListSubItems(47) = var_numero_interno_cliente_ad
    End If
    lv_clientes.SetFocus
End Sub




Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 2220
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes where vcha_emp_empresa_id = '" + var_empresa + "' or vcha_age_agente_id = '00100' order by vcha_age_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 1
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
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_agentes = Me.Name
      frmagentes.Show
   End If
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_agente) <> "" Then
      If Me.txt_agente = "00100" Then
         rs.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_agente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      Else
         rs.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_agente + "' and vcha_Emp_empresa_id = '" + var_empresa + "' ", cnn_distribucion, adOpenDynamic, adLockOptimistic
      End If
      If Not rs.EOF Then
         txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      Else
         MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
         txt_agente = ""
         txt_nombre_agente = ""
      End If
      rs.Close
   Else
      txt_nombre_agente = ""
   End If
End Sub

Private Sub txt_agrupador_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 3060
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_FAMILIA_AGRUPADORES order by VCHA_FAG_NOMBRE", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_FAG_FAMILIA_AGRUPADOR_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_fag_nombre), "", rs!vcha_fag_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "FAMILIA DE AGRUPADORES"
      var_tipo_lista = 9
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
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_familia_agrupadores = Me.Name
      frmfamilia_agrupadores.Show
   End If
End Sub

Private Sub txt_agrupador_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agrupador_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_agrupador) <> "" Then
      rs.Open "select * from TB_FAMILIA_AGRUPADORES where VCHA_FAG_FAMILIA_AGRUPADOR_ID = '" + txt_agrupador + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_agrupador = IIf(IsNull(rs!vcha_fag_nombre), "", rs!vcha_fag_nombre)
      Else
         MsgBox "Clave de familia de agrupadores incorrecta", vbOKOnly, "ATENCION"
         txt_agrupador = ""
         txt_nombre_agrupador = ""
      End If
      rs.Close
   Else
      txt_nombre_agrupador = ""
   End If
End Sub



Private Sub txt_ciudad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ciudad_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_ciudad) <> "" Then
      rs.Open "select * from TB_CIUDADES where VCHA_CIU_CIUDAD_ID = '" + txt_ciudad + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
      Else
         MsgBox "Clave de ciudad incorrecta", vbOKOnly, "ATENCION"
         txt_ciudad = ""
         txt_nombre_ciudad = ""
      End If
      rs.Close
   Else
      txt_nombre_ciudad = ""
   End If
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
   
End Sub

Private Sub txt_codigo_postal_KeyPress(KeyAscii As Integer)
   Dim var_n As Integer
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   If KeyAscii = 13 Then
      If Trim(txt_codigo_postal) <> "" Then
         rs.Open "select distinct vcha_pai_pais_id from tb_colonias where vcha_col_cp = '" + txt_codigo_postal + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
         Dim var_ren As Integer
         var_ren = rs.RecordCount
         rs.Close
         If var_ren > 1 Then
            frm_lista.Left = 5850
            frm_lista.Top = 1230
            lv_lista.ListItems.Clear
            rsaux.Open "select DISTINCT VCHA_PAI_PAIS_ID, VCHA_PAI_NOMBRE from vw_colonias order by vcha_pai_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  Set list_item = lv_lista.ListItems.Add(, , rsaux!VCHA_PAI_PAIS_ID)
                  list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_pai_nombre), "", rsaux!vcha_pai_nombre)
                  rsaux.MoveNext
            Wend
            rsaux.Close
            lbl_lista = "SELECCIONE EL PAIS"
            var_tipo_lista = 16
             var_n = lv_lista.ListItems.Count
            If var_n > 6 Then
               lv_lista.ColumnHeaders(2).Width = 4270.71
            Else
               lv_lista.ColumnHeaders(2).Width = 4499.71
            End If
            frm_lista.Visible = True
            lv_lista.SetFocus
         Else
            rs.Open "select * from vw_colonias where vcha_col_cp = '" + txt_codigo_postal + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               If rs.RecordCount = 1 Then
                  txt_colonia = IIf(IsNull(rs!VCHA_COL_COLONIA_ID), "", rs!VCHA_COL_COLONIA_ID)
                  txt_nombre_colonia = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
                  txt_pais = IIf(IsNull(rs!VCHA_PAI_PAIS_ID), "", rs!VCHA_PAI_PAIS_ID)
                  txt_nombre_pais = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
                  txt_estado = IIf(IsNull(rs!VCHA_EST_ESTADO_ID), "", rs!VCHA_EST_ESTADO_ID)
                  txt_nombre_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                  txt_municipio = IIf(IsNull(rs!VCHA_MUN_MUNICIPIO_ID), "", rs!VCHA_MUN_MUNICIPIO_ID)
                  txt_nombre_municipio = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
                  txt_ciudad = IIf(IsNull(rs!VCHA_CIU_CIUDAD_ID), "", rs!VCHA_CIU_CIUDAD_ID)
                  txt_nombre_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                  txt_email.SetFocus
               Else
                  lv_colonias.ListItems.Clear
                  While Not rs.EOF
                        Set list_item = lv_colonias.ListItems.Add(, , rs!VCHA_COL_COLONIA_ID)
                        list_item.SubItems(1) = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
                        list_item.SubItems(2) = IIf(IsNull(rs!VCHA_PAI_PAIS_ID), "", rs!VCHA_PAI_PAIS_ID)
                        list_item.SubItems(3) = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
                        list_item.SubItems(4) = IIf(IsNull(rs!VCHA_EST_ESTADO_ID), "", rs!VCHA_EST_ESTADO_ID)
                        list_item.SubItems(5) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                        list_item.SubItems(6) = IIf(IsNull(rs!VCHA_MUN_MUNICIPIO_ID), "", rs!VCHA_MUN_MUNICIPIO_ID)
                        list_item.SubItems(7) = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
                        list_item.SubItems(8) = IIf(IsNull(rs!VCHA_CIU_CIUDAD_ID), "", rs!VCHA_CIU_CIUDAD_ID)
                        list_item.SubItems(9) = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                        rs.MoveNext
                  Wend
                  lbl_colonias = "COLONIAS DEL C.P. " + txt_codigo_postal
                  var_n = lv_colonias.ListItems.Count
                  If var_n > 6 Then
                     lv_colonias.ColumnHeaders(2).Width = 4270.71
                  Else
                     lv_colonias.ColumnHeaders(2).Width = 4499.71
                  End If
                  frm_colonias.Visible = True
                  lv_colonias.SetFocus
               End If
            Else
               MsgBox "Código postal incorrecto", vbOKOnly, "ATENCION"
            End If
            rs.Close
         End If
      Else
         txt_telefono.SetFocus
      End If
   End If
End Sub

Private Sub txt_colonia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_colonia_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_colonia) <> "" Then
      rs.Open "select * from TB_COLONIAS where VCHA_COL_COLONIA_ID = '" + txt_colonia + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_colonia = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
      Else
         MsgBox "Clave de colonias incorrecta", vbOKOnly, "ATENCION"
         txt_colonia = ""
         txt_nombre_colonia = ""
      End If
      rs.Close
   Else
      txt_nombre_colonia = ""
   End If
End Sub

Private Sub txt_curp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Or KeyAscii = 32 Or KeyAscii = 45 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_domicilio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_email_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   var_hubo_cambios = True
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub

Private Sub txt_estado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_estado_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_estado) <> "" Then
      rs.Open "select * from TB_ESTADOS where VCHA_EST_ESTADO_ID = '" + txt_estado + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
      Else
         MsgBox "Clave de estado incorrecta", vbOKOnly, "ATENCION"
         txt_estado = ""
         txt_nombre_estado = ""
      End If
      rs.Close
   Else
      txt_nombre_estado = ""
   End If
End Sub

Private Sub txt_fecha_captura_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_lista_precios_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 4608
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_LISTADEPRECIOS order by VCHA_LIS_NOMBRE", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_LIS_LISTA_iD)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_lIS_NOMBRE), "", rs!VCHA_lIS_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "LISTA DE PRECIOS"
      var_tipo_lista = 6
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
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_listadeprecios = Me.Name
      frmlistadeprecios.Show
   End If
End Sub

Private Sub txt_lista_precios_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_lista_precios_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_lista_precios) <> "" Then
      rs.Open "select * from TB_LISTADEPRECIOS where VCHA_LIS_LISTA_ID = '" + txt_lista_precios + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_lista_precios = IIf(IsNull(rs!VCHA_lIS_NOMBRE), "", rs!VCHA_lIS_NOMBRE)
      Else
         MsgBox "Clave de lista precios incorrecta", vbOKOnly, "ATENCION"
         txt_lista_precios = ""
         txt_nombre_lista_precios = ""
      End If
      rs.Close
   Else
      txt_nombre_lista_precios = ""
   End If
End Sub

Private Sub txt_moneda_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 3645
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_MONEDAS order by vcha_mon_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_mon_moneda_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_mon_nombre), "", rs!vcha_mon_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MONEDAS"
      var_tipo_lista = 3
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
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_monedas = Me.Name
      frmmonedas.Show
   End If
End Sub

Private Sub txt_moneda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_moneda_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_moneda) <> "" Then
      rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + txt_moneda + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_moneda = IIf(IsNull(rs!vcha_mon_nombre), "", rs!vcha_mon_nombre)
      Else
         MsgBox "Clave de moneda incorrecta", vbOKOnly, "ATENCION"
         txt_moneda = ""
         txt_nombre_moneda = ""
      End If
      rs.Close
   Else
      txt_nombre_moneda = ""
   End If
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 2220
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_age_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 1
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
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_agentes = Me.Name
      frmagentes.Show
   End If
End Sub

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_agrupador_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 3060
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_FAMILIA_AGRUPADORES order by VCHA_FAG_NOMBRE", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_FAG_FAMILIA_AGRUPADOR_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_fag_nombre), "", rs!vcha_fag_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "FAMILIA DE AGRUPADORES"
      var_tipo_lista = 9
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
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_familia_agrupadores = Me.Name
      frmfamilia_agrupadores.Show
   End If
End Sub

Private Sub txt_nombre_agrupador_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub



Private Sub txt_nombre_ciudad_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_colonia_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_estado_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_lista_precios_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 4608
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_LISTADEPRECIOS order by VCHA_LIS_NOMBRE", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_LIS_LISTA_iD)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_lIS_NOMBRE), "", rs!VCHA_lIS_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "LISTA DE PRECIOS"
      var_tipo_lista = 6
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
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_listadeprecios = Me.Name
      frmlistadeprecios.Show
   End If
End Sub

Private Sub txt_nombre_lista_precios_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_moneda_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 3645
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_MONEDAS order by vcha_mon_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_mon_moneda_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_mon_nombre), "", rs!vcha_mon_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MONEDAS"
      var_tipo_lista = 3
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
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_monedas = Me.Name
      frmmonedas.Show
   End If
End Sub

Private Sub txt_nombre_moneda_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_pais_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_plazo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 4005
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_plazos order by vcha_pla_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_PLA_PLAZO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_pla_nombre), "", rs!vcha_pla_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PLAZOS"
      var_tipo_lista = 4
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
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_plazos = Me.Name
      frmplazos.Show
   End If
End Sub

Private Sub txt_nombre_plazo_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_ruta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 2550
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_rutas order by vcha_rut_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_rut_ruta_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "RUTAS"
      var_tipo_lista = 2
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
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_rutas = Me.Name
      frmrutas.Show
   End If
End Sub

Private Sub txt_nombre_ruta_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_tipo_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 4320
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TIPOSCLIENTES order by VCHA_TCL_NOMBRE", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TCL_TIPO_CLIENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TCL_nombre), "", rs!VCHA_TCL_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TIPOS DE CLIENTES"
      var_tipo_lista = 5
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
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_tiposclientes = Me.Name
      frmtiposclientes.Show
   End If
End Sub

Private Sub txt_nombre_tipo_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_transporte_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 2280
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TRANSPORTES order by VCHA_TRN_NOMBRE", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TRN_TRANSPORTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_trn_nombre), "", rs!vcha_trn_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TRANSPORTES"
      var_tipo_lista = 8
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
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_transportes = Me.Name
      frmtransportes.Show
   End If
End Sub

Private Sub txt_nombre_transporte_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_pais_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_pais_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_pais) <> "" Then
      rs.Open "select * from TB_PAISES where VCHA_PAI_PAIS_ID = '" + txt_pais + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_pais = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
      Else
         MsgBox "Clave de pais incorrecta", vbOKOnly, "ATENCION"
         txt_pais = ""
         txt_nombre_pais = ""
      End If
      rs.Close
   Else
      txt_nombre_pais = ""
   End If
End Sub

Private Sub txt_plazo_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_plazo_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_plazo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 4005
      lv_lista.ListItems.Clear
      If UCase(parametros(0)) = "ADMCDINDUSTRIAL" Then
         rs.Open "select * from tb_plazos WHERE VCHA_PLA_PLAZO_ID = '4' order by vcha_pla_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      Else
         rs.Open "select * from tb_plazos order by vcha_pla_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      End If
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_PLA_PLAZO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_pla_nombre), "", rs!vcha_pla_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PLAZOS"
      var_tipo_lista = 4
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
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_plazos = Me.Name
      frmplazos.Show
   End If
End Sub

Private Sub txt_plazo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Else
      If UCase(parametros(0)) = "ADMCDINDUSTRIAL" Then
         'MsgBox CStr(KeyAscii)
         If KeyAscii <> 13 Then
            If KeyAscii <> 8 Then
               If KeyAscii <> 27 Then
                  KeyAscii = Asc(UCase(Chr(52)))
               End If
            End If
         Else
            Call pro_enfoque(13)
         End If
      Else
         KeyAscii = Asc(UCase(Chr(KeyAscii)))
      End If
   End If
   var_hubo_cambios = True
   
End Sub

Private Sub txt_plazo_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_plazo) <> "" Then
      rs.Open "select * from tb_plazos where vcha_pla_plazo_id = '" + txt_plazo + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_plazo = IIf(IsNull(rs!vcha_pla_nombre), "", rs!vcha_pla_nombre)
      Else
         MsgBox "Clave de plazo incorrecta", vbOKOnly, "ATENCION"
         txt_plazo = ""
         txt_nombre_plazo = ""
      End If
      rs.Close
   Else
      txt_nombre_plazo = ""
   End If
End Sub

Private Sub txt_prioridad_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_prioridad_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_prioridad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_referencia_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_referencia_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub txt_representante_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_representante_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_representante_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_rfc_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_rfc_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_rfc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Or KeyAscii = 32 Or KeyAscii = 45 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ruta_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_ruta_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_ruta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 2550
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_rutas order by vcha_rut_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_rut_ruta_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "RUTAS"
      var_tipo_lista = 2
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
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_rutas = Me.Name
      frmrutas.Show
   End If
End Sub

Private Sub txt_ruta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ruta_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_ruta) <> "" Then
      rs.Open "select * from tb_rutas where vcha_rut_ruta_id = '" + txt_ruta + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
      Else
         MsgBox "Clave de ruta incorrecta", vbOKOnly, "ATENCION"
         txt_ruta = ""
         txt_nombre_ruta = ""
      End If
      rs.Close
   Else
      txt_nombre_ruta = ""
   End If
End Sub

Private Sub txt_telefono_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub

Private Sub txt_telefono_KeyPress(KeyAscii As Integer)
  var_hubo_cambios = True
  Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_tipo_cliente_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_tipo_cliente_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_tipo_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 4320
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TIPOSCLIENTES order by VCHA_TCL_NOMBRE", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TCL_TIPO_CLIENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TCL_nombre), "", rs!VCHA_TCL_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TIPOS DE CLIENTES"
      var_tipo_lista = 5
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
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_tiposclientes = Me.Name
      frmtiposclientes.Show
   End If
End Sub

Private Sub txt_tipo_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_tipo_cliente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_tipo_cliente) <> "" Then
      rs.Open "select * from TB_TIPOSCLIENTES where VCHA_TCL_TIPO_CLIENTE_ID = '" + txt_tipo_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_tipo_cliente = IIf(IsNull(rs!VCHA_TCL_nombre), "", rs!VCHA_TCL_nombre)
      Else
         MsgBox "Clave de tipo cliente incorrecta", vbOKOnly, "ATENCION"
         txt_tipo_cliente = ""
         txt_nombre_tipo_cliente = ""
      End If
      rs.Close
   Else
      txt_nombre_tipo_cliente = ""
   End If
End Sub

Private Sub txt_transporte_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_transporte_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_transporte_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 2280
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TRANSPORTES order by VCHA_TRN_NOMBRE", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TRN_TRANSPORTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_trn_nombre), "", rs!vcha_trn_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TRANSPORTES"
      var_tipo_lista = 8
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
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_transportes = Me.Name
      frmtransportes.Show
   End If
End Sub

Private Sub txt_transporte_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_transporte_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_transporte) <> "" Then
      rs.Open "select * from TB_TRANSPORTES where VCHA_TRN_TRANSPORTE_ID = '" + txt_transporte + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_transporte = IIf(IsNull(rs!vcha_trn_nombre), "", rs!vcha_trn_nombre)
      Else
         MsgBox "Clave de transporte incorrecta", vbOKOnly, "ATENCION"
         txt_transporte = ""
         txt_nombre_transporte = ""
      End If
      rs.Close
   Else
      txt_nombre_transporte = ""
   End If
End Sub

Private Sub txt_unificador_GotFocus()
   Me.frm_busqueda_clientes.Visible = False
End Sub
