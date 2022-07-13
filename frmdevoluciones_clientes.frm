VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmdevoluciones_clientes 
   Caption         =   "Form1"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   Icon            =   "frmdevoluciones_clientes.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   11865
   Begin VB.Frame frm_datos 
      Caption         =   " Cliente "
      Height          =   3720
      Left            =   60
      TabIndex        =   22
      Top             =   1815
      Width           =   5055
      Begin VB.TextBox txt_cp 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   34
         Top             =   3300
         Width           =   1425
      End
      Begin VB.TextBox txt_pais 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   33
         Top             =   2955
         Width           =   3120
      End
      Begin VB.TextBox txt_estado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1035
         TabIndex        =   32
         Top             =   2625
         Width           =   3915
      End
      Begin VB.TextBox txt_ciudad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1035
         TabIndex        =   31
         Top             =   2295
         Width           =   3915
      End
      Begin VB.TextBox txt_direccion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1035
         TabIndex        =   30
         Top             =   1635
         Width           =   3915
      End
      Begin VB.TextBox txt_colonia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1035
         TabIndex        =   29
         Top             =   1965
         Width           =   3915
      End
      Begin VB.TextBox txt_descuentos 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1035
         TabIndex        =   28
         Top             =   1290
         Width           =   2280
      End
      Begin VB.TextBox txt_rfc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1035
         TabIndex        =   27
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cmb_clientes 
         Height          =   315
         Left            =   2085
         TabIndex        =   26
         Top             =   615
         Width           =   2880
      End
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   1035
         TabIndex        =   25
         Top             =   615
         Width           =   1035
      End
      Begin VB.TextBox txt_establecimiento 
         Height          =   315
         Left            =   1035
         TabIndex        =   24
         Top             =   270
         Width           =   1035
      End
      Begin VB.ComboBox cmb_establecimientos 
         Height          =   315
         Left            =   2085
         TabIndex        =   23
         Top             =   270
         Width           =   2880
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "CP:"
         Height          =   195
         Left            =   90
         TabIndex        =   44
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Left            =   90
         TabIndex        =   43
         Top             =   3015
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   90
         TabIndex        =   42
         Top             =   2685
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Left            =   90
         TabIndex        =   41
         Top             =   2355
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
         Height          =   195
         Left            =   90
         TabIndex        =   40
         Top             =   1695
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Colonia:"
         Height          =   195
         Left            =   90
         TabIndex        =   39
         Top             =   2025
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Descuentos:"
         Height          =   195
         Left            =   90
         TabIndex        =   38
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "RFC:"
         Height          =   195
         Left            =   90
         TabIndex        =   37
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   36
         Top             =   645
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Establecim.:"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   35
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Movimiento "
      Height          =   1155
      Left            =   5205
      TabIndex        =   6
      Top             =   600
      Width           =   6540
      Begin VB.TextBox txt_folio 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1005
         TabIndex        =   9
         Top             =   315
         Width           =   1650
      End
      Begin VB.TextBox txt_codigo 
         Height          =   315
         Left            =   1005
         TabIndex        =   8
         Top             =   750
         Width           =   1650
      End
      Begin VB.TextBox txt_cantidad 
         Height          =   315
         Left            =   4260
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   3555
         TabIndex        =   12
         Top             =   795
         Width           =   675
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   345
         Width           =   600
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   810
         Width           =   600
      End
   End
   Begin VB.Frame frm_eliminar 
      Height          =   840
      Left            =   6990
      TabIndex        =   3
      Top             =   3060
      Width           =   2910
      Begin VB.TextBox txt_cantidad_eliminar 
         Height          =   330
         Left            =   90
         TabIndex        =   4
         Top             =   390
         Width           =   2745
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad a eliminar"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   5
         Top             =   15
         Width           =   2895
      End
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   555
      TabIndex        =   0
      Top             =   315
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
         Width           =   3075
      End
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   3075
      Top             =   60
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
            Picture         =   "frmdevoluciones_clientes.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones_clientes.frx":11A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   135
      TabIndex        =   45
      Top             =   30
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo Movimiento"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar Movimiento"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Factura"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir de Esta Ventana"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   0
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones_clientes.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones_clientes.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones_clientes.frx":2C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones_clientes.frx":31CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones_clientes.frx":3AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones_clientes.frx":4384
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones_clientes.frx":4C5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones_clientes.frx":4D70
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones_clientes.frx":4E82
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones_clientes.frx":4F94
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones_clientes.frx":50A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones_clientes.frx":51B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones_clientes.frx":533A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3990
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones_clientes.frx":544C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   " Salida a Vistas "
      Height          =   1155
      Left            =   75
      TabIndex        =   15
      Top             =   600
      Width           =   5040
      Begin VB.TextBox txt_numero_salida 
         Height          =   375
         Left            =   1035
         TabIndex        =   18
         Top             =   285
         Width           =   1695
      End
      Begin VB.TextBox txt_nombre_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2085
         TabIndex        =   17
         Top             =   720
         Width           =   2880
      End
      Begin VB.TextBox txt_clave_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1035
         TabIndex        =   16
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   90
         TabIndex        =   20
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   780
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Artículos "
      Height          =   3720
      Index           =   0
      Left            =   5205
      TabIndex        =   13
      Top             =   1815
      Width           =   6585
      Begin MSComctlLib.ListView lv_existencias 
         Height          =   3465
         Left            =   45
         TabIndex        =   14
         Top             =   210
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   6112
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
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Disponibles"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Facturar"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Faltan"
            Object.Width           =   1676
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   0
      TabIndex        =   21
      Top             =   360
      Width           =   11835
   End
End
Attribute VB_Name = "frmdevoluciones_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    If var_activa_menu = True Then
       Frmmenu2.Enabled = True
    End If

End Sub

