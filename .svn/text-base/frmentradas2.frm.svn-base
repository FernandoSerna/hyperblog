VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmentradas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   3150
      Index           =   1
      Left            =   105
      TabIndex        =   8
      Top             =   2145
      Width           =   4140
      Begin VB.ComboBox cmb_almacen_destino 
         Height          =   315
         Left            =   810
         TabIndex        =   27
         Top             =   450
         Width           =   3255
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   2280
         TabIndex        =   21
         Top             =   1185
         Width           =   1800
      End
      Begin VB.Frame frm_devoluciones 
         BorderStyle     =   0  'None
         Height          =   2280
         Left            =   30
         TabIndex        =   45
         Top             =   645
         Width           =   4065
         Begin VB.TextBox txt_dev_almacen_origen 
            Height          =   330
            Left            =   780
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   165
            Width           =   3225
         End
         Begin VB.Label lbl_titulo_devolueciones 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   150
            TabIndex        =   48
            Top             =   570
            Width           =   2025
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Origen:"
            Height          =   195
            Left            =   210
            TabIndex        =   47
            Top             =   195
            Width           =   510
         End
      End
      Begin VB.Frame frm_notas_envio 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   2145
         Left            =   90
         TabIndex        =   28
         Top             =   780
         Width           =   4005
         Begin VB.ComboBox cmb_almacen_origen 
            Height          =   315
            Left            =   720
            TabIndex        =   37
            Top             =   30
            Width           =   3270
         End
         Begin VB.TextBox txt_chofer 
            Height          =   330
            Left            =   585
            TabIndex        =   30
            Top             =   1770
            Width           =   3375
         End
         Begin VB.Label lbl_titulo_notas_envio 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   375
            TabIndex        =   32
            Top             =   450
            Width           =   1710
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Origen:"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   31
            Top             =   60
            Width           =   510
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Chofer:"
            Height          =   195
            Left            =   45
            TabIndex        =   29
            Top             =   1830
            Width           =   510
         End
      End
      Begin VB.Frame frm_ordenes_compra 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   2145
         Left            =   60
         TabIndex        =   33
         Top             =   810
         Width           =   3990
         Begin VB.TextBox txt_factura 
            Enabled         =   0   'False
            Height          =   330
            Left            =   960
            TabIndex        =   42
            Top             =   750
            Width           =   1875
         End
         Begin VB.TextBox txt_proveedor 
            Enabled         =   0   'False
            Height          =   330
            Left            =   975
            TabIndex        =   36
            Top             =   15
            Width           =   3000
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Factura:"
            Height          =   195
            Index           =   3
            Left            =   345
            TabIndex        =   41
            Top             =   795
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor:"
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   35
            Top             =   60
            Width           =   780
         End
         Begin VB.Label lbl_titulo_orden_compra 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   105
            TabIndex        =   34
            Top             =   420
            Width           =   1980
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   26
         Top             =   480
         Width           =   585
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   9
         Top             =   120
         Width           =   4065
      End
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   840
      TabIndex        =   23
      Top             =   1125
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   120
         Width           =   3075
      End
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   12405
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1785
      Width           =   2100
   End
   Begin MSComDlg.CommonDialog cmdentradas 
      Left            =   3105
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Busqueda de archivo"
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Index           =   1
      Left            =   11325
      TabIndex        =   17
      Top             =   720
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   1965
      Index           =   4
      Left            =   2265
      TabIndex        =   13
      Top             =   5310
      Width           =   1995
      Begin VB.Label lbl_recibidos 
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
         Height          =   525
         Left            =   105
         TabIndex        =   19
         Top             =   885
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad Recibida"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   5
         Left            =   30
         TabIndex        =   14
         Top             =   120
         Width           =   1920
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1965
      Index           =   3
      Left            =   105
      TabIndex        =   11
      Top             =   5295
      Width           =   1995
      Begin VB.Label lbl_enviados 
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
         Height          =   525
         Left            =   135
         TabIndex        =   18
         Top             =   885
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad Enviada"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   4
         Left            =   30
         TabIndex        =   12
         Top             =   120
         Width           =   1920
      End
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Index           =   0
      Left            =   90
      TabIndex        =   6
      Top             =   1200
      Width           =   4155
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
         TabIndex        =   10
         Top             =   405
         Width           =   4050
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   7
         Top             =   120
         Width           =   4080
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6075
      Left            =   4305
      TabIndex        =   1
      Top             =   1200
      Width           =   7515
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
         Left            =   5115
         TabIndex        =   44
         Top             =   555
         Width           =   1890
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   1785
         TabIndex        =   38
         Top             =   1755
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   75
            TabIndex        =   40
            Top             =   375
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   "Cantidad a eliminar"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   39
            Top             =   15
            Width           =   2895
         End
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
         TabIndex        =   3
         Top             =   495
         Width           =   2640
      End
      Begin MSComctlLib.ListView lv_entradas 
         Height          =   4875
         Left            =   45
         TabIndex        =   20
         Top             =   1140
         Width           =   7380
         _ExtentX        =   13018
         _ExtentY        =   8599
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Env."
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Rec."
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Mov."
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Faltan"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Costo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "precio"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   4410
         TabIndex        =   43
         Top             =   675
         Width           =   675
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   4
         Top             =   120
         Width           =   7440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   675
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   570
      Width           =   11745
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   60
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
            Picture         =   "frmentradas2.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas2.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas2.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas2.frx":1750
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas2.frx":202C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas2.frx":2906
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas2.frx":31E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas2.frx":32F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas2.frx":3404
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas2.frx":3516
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas2.frx":3628
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas2.frx":373A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Index           =   0
      Left            =   150
      TabIndex        =   16
      Top             =   690
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo Movimiento"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cargar desde archivo"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar Movimiento"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Movimiento"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   60
      TabIndex        =   15
      Top             =   975
      Width           =   11745
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
      Left            =   90
      TabIndex        =   5
      Top             =   75
      Width           =   11445
   End
End
Attribute VB_Name = "frmentradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_primera_vez As Boolean
Dim var_numero_folio As Integer
Dim VAR_TABLA_NOMBRE_ORIGEN As String
Dim VAR_RUTA_TABLA_ORIGEN As String
Dim VAR_CAMPO_CODIGO_ORIGEN As String
Dim VAR_CAMPO_DESCRIPCION_ORIGEN As String
Dim VAR_CAMPO_COSTO_ORIGEN As String
Dim VAR_CAMPO_CANTIDAD_ORIGEN As String
Dim VAR_CAMPO_CANTIDAD_ENTRADA As String
Dim VAR_TABLA_DESTINO As String
Dim VAR_CAMPO_CODIGO_DESTINO As String
Dim VAR_CAMPO_DESCRIPCION_DESTINO As String
Dim VAR_CAMPO_COSTO_DESTINO As String
Dim VAR_CAMPO_CANTIDAD_DESTINO  As String
Dim VAR_CAMPO_NUMERO  As String
Dim var_cantidad_enviada As Double
Dim var_cantidad_recibida As Double
Dim var_almacen_destino As String
Dim var_almacen_origen As String
Dim var_proveedor As String
Dim var_estatus_movimiento As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_modifica As Boolean
Dim var_factura As String
Dim var_cantidad_leida As Double
Dim var_tabla As ADODB.Connection
Dim var_ruta As String



Private Sub cmb_almacen_destino_Click()
   var_almacen_destino = Obtener_llave(cnn, rsaux, "TB_almacenes", "VCHA_ALM_NOMBRE", cmb_almacen_destino, 2, "T")
End Sub

Private Sub cmb_almacen_destino_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
         cmb_almacen_destino.Enabled = False
         cmb_almacen_origen.Enabled = True
         cmb_almacen_origen.SetFocus
      End If
      If VAR_TABLA_DESTINO = "TB_ORDENES_COMPRA" Then
         txt_numero.Enabled = True
         txt_numero.SetFocus
      End If
      If VAR_TABLA_DESTINO = "TB_DEVOLUCIONES" Then
         txt_numero.Enabled = True
         txt_numero.SetFocus
      End If
   Else
      KeyAscii = 0
   End If
   
End Sub

Private Sub cmb_almacen_destino_LostFocus()
   If Len(Trim(cmb_almacen_destino)) = 0 Then
      cmb_almacen_destino.Enabled = True
      cmb_almacen_destino.SetFocus
      cmb_almacen_origen.Enabled = False
      txt_numero.Enabled = False
   Else
      cmb_almacen_destino.Enabled = False
   End If
End Sub

Private Sub cmb_almacen_origen_Click()
   var_almacen_origen = Obtener_llave(cnn, rsaux, "TB_ALMACENES", "VCHA_ALM_NOMBRE", cmb_almacen_origen, 2, "T")
End Sub

Private Sub cmb_almacen_origen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_numero.Enabled = True
      txt_numero.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub cmb_almacen_origen_LostFocus()
   If var_almacen_origen <> "" Then
      cmb_almacen_origen.Enabled = False
      txt_numero.Enabled = True
   Else
      cmb_almacen_origen.SetFocus
   End If
End Sub

Private Sub Form_Load()
   var_cantidad_leida = 1#
   var_estatus_movimiento = ""
   var_almacen_destino = ""
   var_almacen_origen = ""
   var_proveedor = ""
   var_factura = ""
   frm_notas_envio.Visible = False
   frm_ordenes_compra.Visible = False
   frm_eliminar.Visible = False
   frm_devoluciones.Visible = False
   var_modifica = False
   txt_cantidad.Visible = False
   lbl_cantidad.Visible = False
   rs.Open "select * from tb_referencias where vcha_ref_referencia_id = '" + var_clave_referencia + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not IsNull(rs(3).Value) Then
      VAR_TABLA_NOMBRE_ORIGEN = rs(3).Value
   End If
   If Not IsNull(rs(4).Value) Then
      VAR_RUTA_TABLA_ORIGEN = rs(4).Value
   End If
   If Not IsNull(rs(5).Value) Then
      VAR_CAMPO_CODIGO_ORIGEN = rs(5).Value
   End If
   If Not IsNull(rs(6).Value) Then
      VAR_CAMPO_DESCRIPCION_ORIGEN = rs(6).Value
   End If
   If Not IsNull(rs(7).Value) Then
      VAR_CAMPO_COSTO_ORIGEN = rs(7).Value
   End If
   If Not IsNull(rs(8).Value) Then
      VAR_CAMPO_CANTIDAD_ORIGEN = rs(8).Value
   End If
   VAR_CAMPO_CANTIDAD_ENTRADA = rs(9).Value
   VAR_TABLA_DESTINO = rs(10).Value
   VAR_CAMPO_CODIGO_DESTINO = rs(11).Value
   If Not IsNull(rs(12).Value) Then
      VAR_CAMPO_DESCRIPCION_DESTINO = rs(12).Value
   End If
   VAR_CAMPO_COSTO_DESTINO = rs(13).Value
   VAR_CAMPO_CANTIDAD_DESTINO = rs(14).Value
   VAR_CAMPO_NUMERO = rs(15).Value
   rs.Close
   If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
      frm_notas_envio.Visible = True
      lv_entradas.ColumnHeaders(3).Text = "Env."
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_2 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_almacen_origen.hwnd, rs, 3)
         rs.Close
      Else
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_almacen_origen.hwnd, rs, 2)
         rs.Close
      End If
      cmb_almacen_origen.Enabled = False
   End If
   If VAR_TABLA_DESTINO = "TB_ORDENES_COMPRA" Then
      frm_notas_envio.Visible = False
      frm_devoluciones.Visible = False
      frm_ordenes_compra.Visible = True
      lv_entradas.ColumnHeaders(3).Text = "Comp."
   End If
   If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
      frm_notas_envio.Visible = True
      frm_devoluciones.Visible = False
      frm_ordenes_compra.Visible = False
      lv_entradas.ColumnHeaders(3).Text = "Env."
   End If
   If VAR_TABLA_DESTINO = "TB_DEVOLUCIONES" Then
      frm_notas_envio.Visible = False
      frm_devoluciones.Visible = True
      frm_ordenes_compra.Visible = False
      lv_entradas.ColumnHeaders(3).Text = "Dev."
   End If
   txt_numero.Enabled = False
   txt_codigo.Enabled = False
   var_primera_vez = True
   frm_busqueda.Visible = False
   If var_tipo_permiso = 1 Then
      rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      Call RecsetToCombo(cmb_almacen_destino.hwnd, rs, 3)
      rs.Close
   Else
      rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      Call RecsetToCombo(cmb_almacen_destino.hwnd, rs, 2)
      rs.Close
   End If
End Sub

Private Sub lv_entradas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         MsgBox "Imporsible modificar el movimiento", vbOKOnly, "ATENCION"
      Else
         frm_eliminar.Visible = True
         txt_cantidad_eliminar.SetFocus
      End If
   End If
End Sub

Private Sub Toolbar1_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_ENTRADAS_INSERTA = New TB_ENTRADAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   If Index = 0 Then
      Select Case Button.Index
         Case 1
            If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
               frm_notas_envio.Visible = True
               If var_tipo_permiso = 1 Then
                  rs.Open "select * from vw_almacen_permiso_2 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
                  Call RecsetToCombo(cmb_almacen_origen.hwnd, rs, 3)
                  rs.Close
               Else
                  rs.Open "select * from tb_almacenes order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
                  Call RecsetToCombo(cmb_almacen_origen.hwnd, rs, 3)
                  rs.Close
               End If
               cmb_almacen_origen.Enabled = False
            Else
               frm_notas_envio.Visible = False
            End If
            txt_numero.Enabled = False
            txt_codigo.Enabled = False
            var_primera_vez = True
            frm_busqueda.Visible = False
            If var_tipo_permiso = 1 Then
               rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
               Call RecsetToCombo(cmb_almacen_destino.hwnd, rs, 3)
               rs.Close
            Else
               rs.Open "select * from tb_almacenes order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
               Call RecsetToCombo(cmb_almacen_destino.hwnd, rs, 3)
               rs.Close
            End If
            lv_entradas.ListItems.Clear
            var_cantidad_enviada = 0
            var_cantidad_recibida = 0
            var_numero_folio = 0
            var_factura = ""
            txt_factura = ""
            txt_proveedor = ""
            txt_numero = ""
            lbl_recibidos = ""
            lbl_enviados = ""
            txt_folio = ""
            txt_codigo = ""
            cmb_almacen_destino.Enabled = True
            var_estatus_movimiento = ""
         Case 2
            cmdentradas.CancelError = True
            On Error GoTo ErrHandler
            cmdentradas.Flags = cdlOFNHideReadOnly
            cmdentradas.Filter = "Archivos compatibles (" + VAR_TABLA_NOMBRE_ORIGEN + ")|" + VAR_TABLA_NOMBRE_ORIGEN
            cmdentradas.InitDir = VAR_RUTA_TABLA_ORIGEN
            cmdentradas.FilterIndex = 2
            cmdentradas.ShowOpen
            'cmdentradas.FileName
            Exit Sub
            Exit Sub
         Case 3
            frm_busqueda.Visible = True
            txt_busqueda_folio.SetFocus
         Case 4
            If var_numero_folio > 0 Then
               If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
                  If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_ENTRADAS_PRODUCCION.rpt")
                     reporte.RecordSelectionFormula = "{VW_NOTAS_ENVIO.INTE_NEN_NUMERO} = " + txt_numero + " AND {VW_NOTAS_ENVIO.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_origen + "' and {VW_NOTAS_ENVIO.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_NOTAS_ENVIO.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "'"
                     frmvistasprevias.cr.ReportSource = reporte
                     frmvistasprevias.cr.ViewReport
                     frmvistasprevias.Caption = "Reporte de Movimientos"
                     frmvistasprevias.Show 1
                     Set reporte = Nothing
                  End If
                  If VAR_TABLA_DESTINO = "TB_ORDENES_COMPRA" Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_ordenes_compra_movimientos.rpt")
                     reporte.RecordSelectionFormula = "{VW_ORDENES_COMPRA_MOVIMIENTOS2.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_ORDENES_COMPRA_MOVIMIENTOS2.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_destino + "' AND {VW_ORDENES_COMPRA_MOVIMIENTOS2.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and {VW_ORDENES_COMPRA_MOVIMIENTOS2.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "'"
                     frmvistasprevias.cr.ReportSource = reporte
                     frmvistasprevias.cr.ViewReport
                     frmvistasprevias.Caption = "Reporte de Movimientos"
                     frmvistasprevias.Show 1
                     Set reporte = Nothing
                  End If
               Else
                  var_si = MsgBox("¿Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
                  If var_si = 1 Then
                     Cadena = "select * from tb_temporal_entradas where vcha_alm_almacen_id = " + var_almacen_destino + " and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
                     rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     While Not rs.EOF
                         var_inserta = False
                         var_inserta = TB_ENTRADAS_INSERTA.Anadir(rs(0).Value, rs(1).Value, rs(2).Value, rs(3).Value, rs(4).Value, rs(5).Value, rs(6).Value, rs(7).Value, rs(8).Value, rs(9).Value, rs(10).Value)
                         rs.MoveNext
                     Wend
                     rs.Close
                     var_estatus_movimiento = "I"
                     var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, Now, var_numero_folio, txt_numero, "", "", var_proveedor, var_almacen_origen, var_almacen_destino, "I", fun_NombreUsuario, fun_NombrePc, Now)
                     If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
                        Set reporte = appl.OpenReport(App.Path + "\rep_ENTRADAS_PRODUCCION.rpt")
                        reporte.RecordSelectionFormula = "{VW_NOTAS_ENVIO.INTE_NEN_NUMERO} = " + txt_numero + " AND {VW_NOTAS_ENVIO.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_origen + "' and {VW_NOTAS_ENVIO.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_NOTAS_ENVIO.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "'"
                        frmvistasprevias.cr.ReportSource = reporte
                        frmvistasprevias.cr.ViewReport
                        frmvistasprevias.Caption = "Reporte de Movimientos"
                        frmvistasprevias.Show 1
                        Set reporte = Nothing
                     End If
                     If VAR_TABLA_DESTINO = "TB_ORDENES_COMPRA" Then
                        Set reporte = appl.OpenReport(App.Path + "\rep_ordenes_compra_movimientos.rpt")
                        reporte.RecordSelectionFormula = "{VW_ORDENES_COMPRA_MOVIMIENTOS2.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_ORDENES_COMPRA_MOVIMIENTOS2.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_destino + "' AND {VW_ORDENES_COMPRA_MOVIMIENTOS2.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' {VW_ORDENES_COMPRA_MOVIMIENTOS2.VCHA_UOR_UNIDAD_ID} +'" + var_unidad_organizacional + "'"
                        frmvistasprevias.cr.ReportSource = reporte
                        frmvistasprevias.cr.ViewReport
                        frmvistasprevias.Caption = "Reporte de Movimientos"
                        frmvistasprevias.Show 1
                        Set reporte = Nothing
                     End If
                     txt_codigo.Enabled = False
                     txt_foco.Enabled = False
                  End If
               End If
            Else
               MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
            End If
          Case 5
   End Select
   End If
   If Index = 1 Then
      Unload Me
      frmcodigo_acceso.Show
   End If
ErrHandler:
   
End Sub

Private Sub txt_archivo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      On Error GoTo ersalir:
      rs.Open "select * from tb_principal", cnn, adOpenDynamic, adLockOptimistic
      var_ruta = rs!VCHA_PRI_RUTA_NOTAS_ENVIO
      rs.Close
      var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Tablas de Visual FoxPro;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
      rs.Open "select cvetienda,folio,codigo,cant1,costo from " + txt_leer, var_tabla, adOpenDynamic, adLockOptimistic
      var_almacen_origen_tem = rs(0).Value
      var_posible = 1
      MsgBox "si"
      If var_tipo_permiso = 1 Then
      End If
      If var_posible = 1 Then
         var_almacen_origen = rs(0).Value
         var_numero_salida = rs(1).Value
         rs.Close
      Else
         rs.Close
         MsgBox "No esta autorizado para leer archivos de este almacen", vbOKOnly, "ATENCION"
         frm_devoluciones.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      frm_devoluciones.Visible = False
   End If
   Exit Sub
ersalir:
   MsgBox "A surgido un error al leer el archivo, puede que el archivo este siendo utilizado por otro usuario", vbOKOnly, "ATENCION"
   frm_devoluciones.Visible = False
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
Dim var_busqueda_folio As Integer
Dim var_busqueda_numero As Integer
   If KeyAscii = 13 Then
      rs.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + txt_busqueda_folio + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_estatus_movimiento = rs!CHAR_EMO_ESTATUS
         If Not rs.EOF Then
            var_almacen_destino_tem = rs!VCHA_ALM_ALMACEN_ID
            var_almacen_origen_tem = rs!VCHA_EMO_ALMACEN_ORIGEN
            var_posible = 1
            If var_tipo_permiso = 1 Then
               If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
                  rsaux.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_per_almacen_1 = '" + var_almacen_destino_tem + "'", cnn, adOpenDynamic, adLockBatchOptimistic
                  If rsaux.EOF Then
                     var_posible = 0
                  End If
                  rsaux.Close
                  rsaux.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_per_almacen_2 = '" + var_almacen_origen_tem + "'", cnn, adOpenDynamic, adLockBatchOptimistic
                  If rsaux.EOF Then
                     var_posible = 0
                  End If
                  rsaux.Close
               End If
               If VAR_TABLA_DESTINO = "TB_ORDENES_COMPRA" Then
                  rsaux.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_per_almacen_1 = '" + var_almacen_destino_tem + "'", cnn, adOpenDynamic, adLockBatchOptimistic
                  If rsaux.EOF Then
                     var_posible = 0
                  End If
                  rsaux.Close
               End If
            End If
            If var_posible = 1 Then
               var_almacen_destino = rs!VCHA_ALM_ALMACEN_ID
               var_almacen_origen = rs!VCHA_EMO_ALMACEN_ORIGEN
               var_busqueda_numero = rs!INTE_EMO_NUMERO_ORIGEN
               txt_folio = rs!INTE_EMO_NUMERO
               var_numero_folio = rs!INTE_EMO_NUMERO
               txt_numero = rs!INTE_EMO_NUMERO_ORIGEN
               var_proveedor = rs!VCHA_EMO_PROVEEDOR_ID
               If IsNull(rs!VCHA_EMO_FACTURA) Then
                  var_factura = ""
               Else
                  var_factura = rs!VCHA_EMO_FACTURA
               End If
               txt_factura = var_factura
               rs.Close
               var_primera_vez = False
               lv_entradas.ListItems.Clear
               Dim list_item As ListItem
               Cadena = "select " + VAR_CAMPO_CODIGO_DESTINO + ", '                                                    ' as descripcion, " + VAR_CAMPO_CANTIDAD_DESTINO + ", " + VAR_CAMPO_CANTIDAD_ENTRADA + ", " + VAR_CAMPO_COSTO_DESTINO
               If VAR_TABLA_DESTINO = "TB_ORDENES_COMPRA" Then
                  Cadena = Cadena + ", vcha_pro_proveedor_id"
               End If
               If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
                  Cadena = Cadena + ", VCHA_ALM_ALMACEN_ID"
               End If
               Cadena = Cadena + " from " + VAR_TABLA_DESTINO + " WHERE " + VAR_CAMPO_NUMERO + " = " + txt_numero.Text
               If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
                  Cadena = Cadena + " AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "' AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'"
               End If
               rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  If VAR_TABLA_DESTINO = "TB_ORDENES_COMPRA" Then
                     rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_destino + "'", cnn, adOpenDynamic, adLockOptimistic
                     txt_almacen_destino = rsaux(3).Value
                     cmb_almacen_destino.Text = rsaux(3).Value
                     rsaux.Close
                     rsaux.Open "select * from tb_proveedores where vcha_pro_proveedor_id = '" + rs(5).Value + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_proveedor = rsaux(0).Value
                     txt_proveedor = rsaux(1).Value
                     rsaux.Close
                  End If
                  If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
                     rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_destino + "'", cnn, adOpenDynamic, adLockOptimistic
                     txt_almacen_destino = rsaux(3).Value
                     cmb_almacen_destino.Text = rsaux(3).Value
                     rsaux.Close
                     rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + rs(5).Value + "'", cnn, adOpenDynamic, adLockOptimistic
                     txt_almacen_origen = rsaux(3).Value
                     cmb_almacen_origen.Text = rsaux(3).Value
                     rsaux.Close
                  End If
                  cmb_almacen_destino.Enabled = False
                  While Not rs.EOF
                     rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs(0).Value + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        Set list_item = lv_entradas.ListItems.Add(, , rs(0).Value)
                        list_item.SubItems(1) = IIf(IsNull(rsaux(1).Value), "", rsaux(1).Value)
                        list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
                        list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
                        rsaux2.Open "select floa_ent_cantidad from tb_temporal_entradas where vcha_alm_almacen_id = '" + var_almacen_destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + rs(0).Value + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           If IsNull(rsaux2(0).Value) Then
                              list_item.SubItems(4) = 0
                           Else
                              list_item.SubItems(4) = IIf(IsNull(rsaux2(0).Value), "", rsaux2(0).Value)
                           End If
                        Else
                           list_item.SubItems(4) = 0
                        End If
                        rsaux2.Close
                        list_item.SubItems(5) = list_item.SubItems(2) - list_item.SubItems(3)
                        list_item.SubItems(6) = IIf(IsNull(rs(4).Value), "", rs(4).Value)
                        list_item.SubItems(7) = IIf(IsNull(rsaux(2).Value), "", rsaux(2).Value)
                     End If
                     rsaux.Close
                     rs.MoveNext:
                  Wend
                  rs.Close
                  rs.Open "select sum(" + VAR_CAMPO_CANTIDAD_DESTINO + ") as enviados, sum(" + VAR_CAMPO_CANTIDAD_ENTRADA + ")as recibida from " + VAR_TABLA_DESTINO + " WHERE " + VAR_CAMPO_NUMERO + " = " + Str(var_busqueda_numero), cnn, adOpenDynamic, adLockOptimistic
                  If IsNull(rs(0).Value) Then
                     lbl_enviados = "0"
                     var_cantidad_enviada = 0
                  Else
                     lbl_enviados = rs(0).Value
                     var_cantidad_enviada = rs(0).Value
                  End If
                  If IsNull(rs(1).Value) Then
                     lbl_recibidos = "0"
                     var_cantidad_recibida = 0
                  Else
                     lbl_recibidos = rs(1).Value
                     var_cantidad_recibida = rs(1).Value
                  End If
                  rs.Close
                  txt_numero.Enabled = False
                  If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
                     txt_codigo.Enabled = False
                     lv_entradas.SetFocus
                  Else
                     txt_codigo.Enabled = True
                     txt_codigo.SetFocus
                  End If
               End If
            Else
               rs.Close
               MsgBox "No esta autorizado para modificar este movimiento", vbOKOnly, "ATENCION"
               frm_busqueda.Visible = False
            End If
         Else
            MsgBox "El Movimiento se encuentra vacio", vbOKOnly, "ATENCION"
         End If
      Else
         rs.Close
         MsgBox "El número de movimiento " + Trim(txt_busqueda_folio) + " no existe", vbOKOnly, "ATENCION"
         frm_busqueda.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_busqueda_folio_LostFocus()
      frm_busqueda.Visible = False
End Sub

Private Sub txt_cantidad_eliminar_GotFocus()
   txt_cantidad_eliminar = ""
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Set TB_ORDENES_COMPRA_MODIFICA = New TB_ORDENES_COMPRA_MODIFICA
      Set TB_NOTAS_ENVIO_ACTUALIZA = New TB_NOTAS_ENVIO_ACTUALIZA
      Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
      var_cantidad_eliminar = Val(txt_cantidad_eliminar)
      var_cantidad_eliminar_arch = lv_entradas.SelectedItem.SubItems(3) - Val(txt_cantidad_eliminar)
      var_cantidad_eliminar_mov = lv_entradas.SelectedItem.SubItems(4) - Val(txt_cantidad_eliminar)
      If var_cantidad_eliminar_arch < 0 Or var_cantidad_eliminar_mov < 0 Then
         MsgBox "No esposible eliminar esta cantidad", vbOKOnly, "ATENCION"
      Else
         If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
            var_actualiza = TB_NOTAS_ENVIO_ACTUALIZA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, txt_numero, lv_entradas.SelectedItem, 0 - var_cantidad_eliminar, "")
            lv_entradas.SelectedItem.SubItems(3) = lv_entradas.SelectedItem.SubItems(3) - Val(txt_cantidad_eliminar)
            lv_entradas.SelectedItem.SubItems(4) = lv_entradas.SelectedItem.SubItems(4) - Val(txt_cantidad_eliminar)
            lv_entradas.SelectedItem.SubItems(5) = lv_entradas.SelectedItem.SubItems(2) - lv_entradas.SelectedItem.SubItems(3)
         End If
         If VAR_TABLA_DESTINO = "TB_ORDENES_COMPRA" Then
            var_actualiza = TB_ORDENES_COMPRA_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, txt_numero, Now, var_almacen_destino, var_proveedor, lv_entradas.SelectedItem, 0, 0, 0 - var_cantidad_eliminar)
            lv_entradas.SelectedItem.SubItems(3) = lv_entradas.SelectedItem.SubItems(3) - Val(txt_cantidad_eliminar)
            lv_entradas.SelectedItem.SubItems(4) = lv_entradas.SelectedItem.SubItems(4) - Val(txt_cantidad_eliminar)
            lv_entradas.SelectedItem.SubItems(5) = lv_entradas.SelectedItem.SubItems(2) - lv_entradas.SelectedItem.SubItems(3)
         End If
         Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = " + var_almacen_destino + "and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + lv_entradas.SelectedItem + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         var_inserta = False
         var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, var_numero_folio, lv_entradas.SelectedItem, 0 - Val(txt_cantidad_eliminar))
         rs.Close
         lbl_recibidos = Int(lbl_recibidos) - var_cantidad_eliminar
         frm_eliminar.Visible = False
         txt_codigo.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      frm_eliminar.Visible = False
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   frm_eliminar.Visible = False
   txt_codigo.SetFocus
End Sub

Private Sub txt_cantidad_GotFocus()
   txt_cantidad = "1"
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
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

Private Sub txt_codigo_GotFocus()
   txt_codigo = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   Dim var_recontable As Integer
   txt_codigo = Trim(txt_codigo)
   If KeyAscii = 13 Then
      If Trim(txt_codigo) <> "" Then
         rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If IsNull(rs(43).Value) Then
               var_recontable = 0
            Else
               var_recontable = rs(43).Value
            End If
            rs.Close
            If var_recontable = 1 Then
               var_cantidad_leida = 1#
               lbl_cantidad.Visible = True
               txt_cantidad.Visible = True
               txt_cantidad.SetFocus
            Else
               var_cantidad_leida = 1#
               txt_foco.Enabled = True
               txt_foco.SetFocus
            End If
         Else
            rs.Close
            rs.Open "select * from tb_equivalencias where VCHA_EQU_CODIGO_EQUIVALENTE = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               txt_codigo = rs(0).Value
               rs.Close
               rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  If IsNull(rs(43).Value) Then
                     var_recontable = 0
                  Else
                     var_recontable = rs(43).Value
                  End If
                  rs.Close
                  If var_recontable = 1 Then
                     var_cantidad_leida = 1#
                     lbl_cantidad.Visible = True
                     txt_cantidad.Visible = True
                     txt_cantidad.SetFocus
                  Else
                     var_cantidad_leida = 1#
                     txt_foco.Enabled = True
                     txt_foco.SetFocus
                  End If
               Else
                  MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
               rs.Close
            End If
         End If
      Else
      End If
   End If
End Sub

Private Sub txt_factura_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txt_factura) <> "" Then
         var_factura = txt_factura
         txt_codigo.Enabled = True
         txt_codigo.SetFocus
         txt_factura.Enabled = False
      Else
         MsgBox "Debe de indicarse el número de la factura", vbOKOnly, "ATENCION"
         txt_factura.SetFocus
      End If
   End If
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   Dim Cadena As String
   If KeyAscii = 13 Then
      If Trim(txt_numero) <> "" Then
         lv_entradas.ListItems.Clear
         Dim list_item As ListItem
         If VAR_TABLA_DESTINO = "TB_DEVOLUCIONES" Then
            rs.Open "SELECT * FROM TB_ENCABEZADO_MOVIMIENTOS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emo_referencia = '" + txt_numero + "'", cnn, adOpenDynamic, adLockOptimistic
            If rs.EOF Then
               rs.Close
               On Error GoTo ersalir:
               Set var_tabla = CreateObject("ADODB.connection")
               rs.Open "select * from tb_principal", cnn, adOpenDynamic, adLockOptimistic
               var_ruta = rs!VCHA_PRI_RUTA_NOTAS_ENVIO
               rs.Close
               var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Tablas de Visual FoxPro;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
               rs.Open "select cvetienda,folio,codigo,cant1,costo from " + txt_numero, var_tabla, adOpenDynamic, adLockOptimistic
               var_almacen_origen_tem = rs(0).Value
               var_posible = 1
               If var_tipo_permiso = 1 Then
               End If
               If var_posible = 1 Then
                  var_almacen_origen = rs(0).Value
                  var_numero_salida = rs(1).Value
                  txt_folio_enviado = var_numero_salida
                  cmb_almacen_destino.Enabled = True
                  cmb_almacen_destino.SetFocus
                  rs.Close
               Else
                  rs.Close
                  MsgBox "No esta autorizado para leer archivos de este almacen", vbOKOnly, "ATENCION"
               End If
            Else
               rs.Close
            End If
         End If
         If VAR_TABLA_DESTINO = "TB_ORDENES_COMPRA" Then
            var_modifica = True
         End If
         If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
            rs.Open "select * from tb_encabezado_movimientos where inte_emo_numero_origen = " + txt_numero + " and vcha_emo_almacen_origen = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            Cadena = ""
            If Not rs.EOF Then
               var_modifica = False
               Cadena = "La Nota de envio " + txt_numero + " esta siendo utilizada en el movimiento " + Str(rs!INTE_EMO_NUMERO)
            Else
               var_modifica = True
            End If
            rs.Close
         End If
         
         If var_modifica = True Then
            Cadena = "select " + VAR_CAMPO_CODIGO_DESTINO + ", '                                                    ' as descripcion, " + VAR_CAMPO_CANTIDAD_DESTINO + ", " + VAR_CAMPO_CANTIDAD_ENTRADA + ", " + VAR_CAMPO_COSTO_DESTINO
            If VAR_TABLA_DESTINO = "TB_ORDENES_COMPRA" Then
               Cadena = Cadena + ", vcha_pro_proveedor_id"
            End If
            If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
               Cadena = Cadena + ", VCHA_ALM_ALMACEN_ID"
            End If
            Cadena = Cadena + " from " + VAR_TABLA_DESTINO + " WHERE " + VAR_CAMPO_NUMERO + " = " + txt_numero.Text
            If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
               Cadena = Cadena + " AND VCHA_ALM_ALMACEN_ID = " + var_almacen_origen
            End If
            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               If VAR_TABLA_DESTINO = "TB_ORDENES_COMPRA" Then
                  rsaux.Open "select * from tb_proveedores where vcha_pro_proveedor_id = '" + rs(5).Value + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_proveedor = rsaux(0).Value
                  txt_proveedor = rsaux(1).Value
                  rsaux.Close
               End If
               If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
                  rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + rs(5).Value + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_almacen_origen = rsaux(3).Value
                  rsaux.Close
               End If
               While Not rs.EOF
                  rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs(0).Value + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     Set list_item = lv_entradas.ListItems.Add(, , rs(0).Value)
                     list_item.SubItems(1) = IIf(IsNull(rsaux(1).Value), "", rsaux(1).Value)
                     list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
                     list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
                     list_item.SubItems(4) = 0
                     list_item.SubItems(5) = list_item.SubItems(2) - list_item.SubItems(3)
                     list_item.SubItems(6) = IIf(IsNull(rs(4).Value), "", rs(4).Value)
                     list_item.SubItems(7) = IIf(IsNull(rsaux(2).Value), "", rsaux(2).Value)
                  End If
                  rsaux.Close
                  rs.MoveNext:
               Wend
               rs.Close
               rs.Open "select sum(" + VAR_CAMPO_CANTIDAD_DESTINO + ") as enviados, sum(" + VAR_CAMPO_CANTIDAD_ENTRADA + ")as recibida from " + VAR_TABLA_DESTINO + " WHERE " + VAR_CAMPO_NUMERO + " = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
               If IsNull(rs(0).Value) Then
                  lbl_enviados = "0"
                  var_cantidad_enviada = 0
               Else
                  lbl_enviados = rs(0).Value
                  var_cantidad_enviada = rs(0).Value
               End If
               If IsNull(rs(1).Value) Then
                  lbl_recibidos = "0"
                  var_cantidad_recibida = 0
               Else
                  lbl_recibidos = rs(1).Value
                  var_cantidad_recibida = rs(1).Value
               End If
               rs.Close
               If VAR_TABLA_DESTINO = "TB_ORDENES_COMPRA" Then
                  
                  txt_factura.Enabled = True
                  txt_factura.SetFocus
                  txt_numero.Enabled = False
               Else
                  var_factura = ""
                  txt_codigo.Enabled = True
                  txt_codigo.SetFocus
                  txt_numero.Enabled = False
               End If
            Else
               rs.Close
               lbl_recibidos = ""
               lbl_enviados = ""
               var_cantidad_enviada = 0
               var_cantidad_recibida = 0
               txt_codigo.Enabled = False
               MsgBox "El número no existe", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox Cadena, vbOKOnly, "ATENCION"
            txt_codigo.Enabled = False
         End If
      End If
   End If
Exit Sub
ersalir:
   MsgBox "A surgido un error al leer el archivo, puede que el archivo este siendo utilizado por otro usuario", vbOKOnly, "ATENCION"
End Sub

Private Sub txt_foco_GotFocus()
   Dim var_actualiza As Boolean
   Dim var_inserta As Boolean
   Dim bandera_suma As Boolean
   Dim var_cantidad As Variant
   Dim var_costo As Variant
   Dim var_precio As Variant
   Set TB_NOTAS_ENVIO_ACTUALIZA = New TB_NOTAS_ENVIO_ACTUALIZA
   Set TB_NOTAS_ENVIO_INSERTA = New TB_NOTAS_ENVIO_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_I = New TB_ENCABEZADO_MOVIMIENTOS_I
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   Set TB_ORDENES_COMPRA_INSERTA = New TB_ORDENES_COMPRA_INSERTA
   Set TB_ORDENES_COMPRA_MODIFICA = New TB_ORDENES_COMPRA_MODIFICA
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   If Trim(txt_codigo.Text) <> "" Then
      bandera_suma = False
      If var_primera_vez = True Then
         rs.Open "select max(INTE_EMO_NUMERO) as numero from tb_encabezado_movimientos where VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
         If IsNull(rs(0).Value) Then
            var_numero_folio = 1
            var_inserta = False
            var_inserta = TB_ENCABEZADO_MOVIMIENTOS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, Now, var_numero_folio, txt_numero, "", "", var_proveedor, var_almacen_origen, var_almacen_destino, "", fun_NombreUsuario, fun_NombrePc, var_factura, "", "")
         Else
            var_numero_folio = rs(0).Value + 1
            var_inserta = False
            var_inserta = TB_ENCABEZADO_MOVIMIENTOS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, Now, var_numero_folio, txt_numero, "", "", var_proveedor, var_almacen_origen, var_almacen_destino, "", fun_NombreUsuario, fun_NombrePc, var_factura, "", "")
         End If
         rs.Close
         txt_folio = var_numero_folio
         var_primera_vez = False
      End If
      rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         rsaux.Close
         Cadena = "select " + VAR_CAMPO_CODIGO_DESTINO + ", " + VAR_CAMPO_CANTIDAD_DESTINO + ", " + VAR_CAMPO_CANTIDAD_ENTRADA + ", " + VAR_CAMPO_COSTO_DESTINO + " from " + VAR_TABLA_DESTINO + " WHERE " + VAR_CAMPO_NUMERO + " = " + txt_numero.Text + " and " + VAR_CAMPO_CODIGO_DESTINO + " = " + txt_codigo.Text
         If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
            Cadena = Cadena + " and vcha_alm_almacen_id = " + var_almacen_origen
         End If
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            valor = txt_codigo
            Set itmfound = lv_entradas.FindItem(valor, lvwText, , lvwPartial)
            itmfound.EnsureVisible
            itmfound.Selected = True
            bandera_suma = True
            lv_entradas.SelectedItem.SubItems(3) = lv_entradas.SelectedItem.SubItems(3) + var_cantidad_leida
            lv_entradas.SelectedItem.SubItems(4) = lv_entradas.SelectedItem.SubItems(4) + var_cantidad_leida
            lv_entradas.SelectedItem.SubItems(5) = lv_entradas.SelectedItem.SubItems(2) - lv_entradas.SelectedItem.SubItems(3)
            var_costo = lv_entradas.SelectedItem.SubItems(6)
            var_precio = lv_entradas.SelectedItem.SubItems(7)
            var_cantidad = lv_entradas.SelectedItem.SubItems(4)
            lbl_recibidos = Int(lbl_recibidos) + var_cantidad_leida
            var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
            If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
               var_actualiza = TB_NOTAS_ENVIO_ACTUALIZA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, txt_numero, rs(0).Value, var_cantidad_leida, "I")
            End If
            If VAR_TABLA_DESTINO = "TB_ORDENES_COMPRA" Then
               var_actualiza = TB_ORDENES_COMPRA_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, txt_numero, Now, var_almacen_destino, var_proveedor, rs(0).Value, 0, 0, var_cantidad_leida)
            End If
         Else
            rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               If VAR_TABLA_DESTINO = "TB_NOTAS_ENVIO" Then
                  Set list_item = lv_entradas.ListItems.Add(, , rsaux(0).Value)
                  list_item.SubItems(1) = IIf(IsNull(rsaux(1).Value), "", rsaux(1).Value)
                  list_item.SubItems(2) = 0
                  list_item.SubItems(3) = var_cantidad_leida
                  list_item.SubItems(4) = var_cantidad_leida
                  list_item.SubItems(5) = list_item.SubItems(2) - list_item.SubItems(3)
                  list_item.SubItems(6) = IIf(IsNull(rsaux(3).Value), "", rsaux(3).Value)
                  list_item.SubItems(7) = IIf(IsNull(rsaux(2).Value), "", rsaux(2).Value)
                  var_costo = rsaux(3).Value
                  var_precio = rsaux(2).Value
                  bandera_suma = True
                  lbl_recibidos = Int(lbl_recibidos) + var_cantidad_leida
                  var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                  var_inserta = False
                  var_inserta = TB_NOTAS_ENVIO_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, "P", txt_numero, 0, Now, rsaux(0).Value, rsaux(3).Value, 0, var_cantidad_leida, "", "I")
               End If
               If VAR_TABLA_DESTINO = "TB_ORDENES_COMPRA" Then
                  MsgBox "El artículo no existe en la Orden de Compra " + txt_numero, vbOKOnly, "ATENCION"
                  bandera_suma = False
               End If
            Else
               MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
               bandera_suma = False
            End If
            rsaux.Close
         End If
         rs.Close
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         rsaux.Close
      End If
      If bandera_suma = True Then
         Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = " + var_almacen_destino + "and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = " + txt_codigo
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida)
            rs.Close
         Else
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", var_almacen_origen)
            rs.Close
         End If
         bandera_suma = False
      End If
      txt_codigo.SetFocus
   End If
End Sub
