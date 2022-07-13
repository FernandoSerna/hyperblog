VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmentradas_reempaque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entradas Reempaque"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   Icon            =   "frmentradas_reempaque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11625
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   390
      Left            =   6105
      Top             =   705
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   688
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame frm_busqueda 
      Height          =   1245
      Left            =   525
      TabIndex        =   7
      Top             =   1005
      Width           =   5220
      Begin VB.TextBox txt_almacen_busqueda 
         Height          =   300
         Left            =   795
         TabIndex        =   44
         Top             =   480
         Width           =   870
      End
      Begin VB.TextBox txt_nombre_almacen_busqueda 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   43
         Top             =   480
         Width           =   3435
      End
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   795
         TabIndex        =   8
         Top             =   795
         Width           =   1590
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   105
         TabIndex        =   46
         Top             =   885
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Origen:"
         Height          =   195
         Left            =   105
         TabIndex        =   45
         Top             =   525
         Width           =   510
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   2
         Left            =   30
         TabIndex        =   9
         Top             =   120
         Width           =   5145
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2070
      Left            =   105
      TabIndex        =   33
      Top             =   1125
      Width           =   7350
      Begin VB.TextBox txt_archivo 
         Height          =   300
         Left            =   60
         TabIndex        =   42
         Top             =   555
         Width           =   1575
      End
      Begin VB.TextBox txt_nombre_almacen_origen 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1110
         Width           =   5610
      End
      Begin VB.TextBox txt_nombre_almacen_destino 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1650
         Width           =   5610
      End
      Begin VB.TextBox txt_clave_almacen_origen 
         Height          =   300
         Left            =   60
         TabIndex        =   38
         Top             =   1110
         Width           =   1575
      End
      Begin VB.TextBox txt_clave_almacen_destino 
         Height          =   300
         Left            =   60
         TabIndex        =   36
         Top             =   1635
         Width           =   1575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Archivo:"
         Height          =   195
         Left            =   90
         TabIndex        =   41
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Origen:"
         Height          =   195
         Left            =   90
         TabIndex        =   37
         Top             =   900
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   90
         TabIndex        =   35
         Top             =   1410
         Width           =   585
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Datos del Movimiento "
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   6
         Left            =   30
         TabIndex        =   34
         Top             =   120
         Width           =   7275
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1050
      Index           =   0
      Left            =   7545
      TabIndex        =   17
      Top             =   1110
      Width           =   3975
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
         TabIndex        =   18
         Top             =   450
         Width           =   3840
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   1
         Left            =   30
         TabIndex        =   19
         Top             =   120
         Width           =   3885
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Index           =   3
      Left            =   7545
      TabIndex        =   14
      Top             =   2100
      Width           =   1935
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad Enviada"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   4
         Left            =   30
         TabIndex        =   16
         Top             =   120
         Width           =   1860
      End
      Begin VB.Label lbl_enviados 
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
         Height          =   525
         Left            =   165
         TabIndex        =   15
         Top             =   465
         Width           =   1590
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Index           =   4
      Left            =   9570
      TabIndex        =   11
      Top             =   2100
      Width           =   1935
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad Recibida"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   5
         Left            =   30
         TabIndex        =   13
         Top             =   120
         Width           =   1860
      End
      Begin VB.Label lbl_recibidos 
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
         Height          =   525
         Left            =   105
         TabIndex        =   12
         Top             =   465
         Width           =   1695
      End
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   12345
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1785
      Width           =   2100
   End
   Begin VB.TextBox txt_clave_movimiento 
      Height          =   285
      Left            =   2190
      TabIndex        =   6
      Top             =   750
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox txt_tipo_documento 
      Height          =   285
      Left            =   3105
      TabIndex        =   5
      Top             =   750
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmentradas_reempaque.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Nuevo Movimiento Alt + N"
      Top             =   705
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmentradas_reempaque.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Buscar Movimiento Alt + B"
      Top             =   705
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      Picture         =   "frmentradas_reempaque.frx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Movimiento Alt + I"
      Top             =   705
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1050
      Picture         =   "frmentradas_reempaque.frx":0BD0
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cancelar Movimiento Alt + C"
      Top             =   705
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11100
      Picture         =   "frmentradas_reempaque.frx":0CD2
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   720
      Width           =   330
   End
   Begin MSComDlg.CommonDialog cmdentradas 
      Left            =   3045
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Busqueda de archivo"
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   0
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
            Picture         =   "frmentradas_reempaque.frx":130C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_reempaque.frx":1BE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_reempaque.frx":24C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_reempaque.frx":2A5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_reempaque.frx":3338
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_reempaque.frx":3C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_reempaque.frx":44EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_reempaque.frx":45FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_reempaque.frx":4710
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_reempaque.frx":4822
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_reempaque.frx":4934
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_reempaque.frx":4A46
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   0
      TabIndex        =   30
      Top             =   585
      Width           =   11535
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   0
      TabIndex        =   31
      Top             =   960
      Width           =   11550
   End
   Begin VB.Frame Frame2 
      Height          =   4095
      Left            =   90
      TabIndex        =   20
      Top             =   3120
      Width           =   11430
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
         TabIndex        =   25
         Top             =   495
         Width           =   2640
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   1785
         TabIndex        =   22
         Top             =   1755
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   75
            TabIndex        =   23
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
            TabIndex        =   24
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
         Left            =   5115
         TabIndex        =   21
         Top             =   555
         Width           =   1890
      End
      Begin MSComctlLib.ListView lv_entradas 
         Height          =   2895
         Left            =   75
         TabIndex        =   26
         Top             =   1125
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   5106
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Folio"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripción"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Env."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Rec."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Mov."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Faltan"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Costo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "precio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Codigo Salida"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "año"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "consecutivo"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl_cancelado 
         Alignment       =   2  'Center
         Caption         =   "MOVIMIENTO CANCELADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   4755
         TabIndex        =   47
         Top             =   495
         Width           =   6045
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   675
         Width           =   1395
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   28
         Top             =   120
         Width           =   11355
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   4410
         TabIndex        =   27
         Top             =   675
         Width           =   675
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
      TabIndex        =   32
      Top             =   75
      Width           =   11445
   End
End
Attribute VB_Name = "frmentradas_reempaque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_entrada As Integer
Dim var_folio_entrada As String
Dim var_numero_entrada As Double
Dim var_origen As String
Dim var_transporto As String
Dim var_tipo_proveedor As String
Dim var_primera_vez As Boolean
Dim var_numero_folio As Double
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
Dim var_articulo_enviado As String
Dim var_costo_enviado As Double
Dim var_almacen_Destino As String
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
Dim var_folio_enviado As Double
Dim var_referencia As String
Dim var_suma_cantidad_enviada As Double
Dim var_suma_cantidad_recibida As Double
Dim var_numero_causa As Integer
Dim ntablas As Integer
Dim var_fecha_movimiento As Date
Dim var_solo_lectura As Boolean

Dim var_entrada_calidad As Boolean
Dim var_almacen_costeo As String
Dim var_ventana As Integer
Dim var_tipo_Cambio As Double
Dim var_moneda_local As Integer
Dim var_clave_moneda As String
Dim var_unidad_organizacional_origen As String
Dim var_renglon As Double





Private Sub cmd_buscar_Click()
   txt_almacen_busqueda = ""
   txt_nombre_almacen_busqueda = ""
   txt_busqueda_folio = ""
   frm_busqueda.Visible = True
   txt_almacen_busqueda.SetFocus
End Sub



Sub ilumina_grid()
   var_n = lv_entradas.ListItems.Count
   For var_i = 1 To var_n
       If var_i = var_renglon Then
          lv_entradas.ListItems.Item(var_i).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(3).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(4).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(5).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(6).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(7).Bold = True
          lv_entradas.ListItems.Item(var_i).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H8000&
       Else
          If (lv_entradas.ListItems.Item(var_i).ListSubItems(5) * 1) < 0 Then
             lv_entradas.ListItems.Item(var_i).Bold = True
             lv_entradas.ListItems.Item(var_i).ListSubItems(1).Bold = True
             lv_entradas.ListItems.Item(var_i).ListSubItems(2).Bold = True
             lv_entradas.ListItems.Item(var_i).ListSubItems(3).Bold = True
             lv_entradas.ListItems.Item(var_i).ListSubItems(4).Bold = True
             lv_entradas.ListItems.Item(var_i).ListSubItems(5).Bold = True
             lv_entradas.ListItems.Item(var_i).ListSubItems(6).Bold = True
             lv_entradas.ListItems.Item(var_i).ListSubItems(7).Bold = True
             lv_entradas.ListItems.Item(var_i).ForeColor = &HFF&
             lv_entradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF&
             lv_entradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF&
             lv_entradas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF&
             lv_entradas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF&
             lv_entradas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF&
             lv_entradas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF&
             lv_entradas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HFF&
          Else
             lv_entradas.ListItems.Item(var_i).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(1).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(2).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(3).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(4).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(5).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(6).Bold = False
             lv_entradas.ListItems.Item(var_i).ListSubItems(7).Bold = False
             lv_entradas.ListItems.Item(var_i).ForeColor = &H80000012
             lv_entradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000012
             lv_entradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000012
             lv_entradas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000012
             lv_entradas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000012
             lv_entradas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H80000012
             lv_entradas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H80000012
             lv_entradas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H80000012
          End If
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_entradas.ListItems.Item(var_renglon).Selected = True
      lv_entradas.selectedItem.EnsureVisible
   End If
   lv_entradas.Refresh
End Sub




Private Sub cmd_imprimir_Click()
   Dim var_posible_cerrar_movimiento As Integer
   var_posible_cerrar_movimiento = 1
   
   Dim dl As Long                                 ' Valor devuelto por la función API
   Dim sAttributes As String                  ' Aributos
   Dim sDriver As String                       ' Nombre del controlador
   Dim sDescription As String                ' Descripción del DSN
   Dim sDsnName As String                  ' Nombre del DSN

   Const ODBC_ADD_SYS_DSN As Long = 4         ' Se creará un DSN de sistema
   Const vbAPINull As Long = 0&                         ' Puntero NULL

   ' se elimina
   Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema
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
   sAttributes = sAttributes & "Database=" + var_bd_movimientos & Chr(0)
   strAttributes = strAttributes & "UID=sa" & Chr$(0)
   strAttributes = strAttributes & "PWD=elia" & Chr$(0)
   dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
   Dim var_n As Integer, var_i As Integer
   Dim var_cantidad_sobrante As Double
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_ENTRADAS_I = New TB_ENTRADAS_I
   Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_ENTRADAS_VISTAS_I = New TB_ENTRADAS_VISTAS_I
   Set TB_ENTRADAS_INSERTA = New TB_ENTRADAS_INSERTA
   Dim var_codigo_salida As String
   Dim var_codigo_entrada As String
   Dim var_primera_vez_ajuste As Integer
   Dim var_numero_folio_ajuste_entrada As Double
   Dim var_numero_folio_ajuste_salida As Double
   Dim var_numero_folio_sobrante As Double
   Dim var_numero_folio_sobrante_salida As Double
   Dim var_numero_folio_sobrante_almacen As Double
   
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Dim var_clave_movimiento_ajuste_salida As String
   Dim var_clave_movimiento_ajuste_entrada As String
   Dim var_clave_movimiento_sobrante As String
   Dim var_clave_movimiento_sobrante_salida As String
   Dim var_clave_almacen_sobrante As String
   Dim var_clave_unidad_organizacional_sobrante As String
   Dim var_consecutivo As Double
   If var_numero_folio > 0 Then
      If Trim(var_estatus_movimiento) = "" Then
         var_n = lv_entradas.ListItems.Count
         var_cantidad_sobrante = 0
         For var_i = 1 To var_n
         lv_entradas.ListItems.Item(var_i).Selected = True
         If Trim(lv_entradas.selectedItem) = "SOBRANTE" Then
            var_cantidad_sobrante = var_cantidad_sobrante + (lv_entradas.selectedItem.SubItems(5) * 1)
         End If
         Next var_i
         If var_cantidad_sobrante > 0 Then
            rsaux.Open "select vcha_mov_movimiento_id from tb_movimientos where inte_mov_sobrante = 1 and char_mov_afectacion = '+'"
            If Not rsaux.EOF Then
               var_clave_movimiento_sobrante = rsaux!VCHA_MOV_MOVIMIENTO_ID
            End If
            rsaux.Close
            rsaux.Open "select vcha_mov_movimiento_id from tb_movimientos where inte_mov_sobrante = 1 and char_mov_afectacion = '-'"
            If Not rsaux.EOF Then
               var_clave_movimiento_sobrante_salida = rsaux!VCHA_MOV_MOVIMIENTO_ID
            End If
            rsaux.Close
            var_clave_almacen_sobrante = ""
            var_clave_almacen_sobrante_saliDA = ""
            var_clave_unidad_organizacional_sobrante = ""
            rsaux.Open "select vcha_alm_almacen_id, VCHA_UOR_UNIDAD_ID from tb_almacenes where INTE_ALM_SOBRANTES = 1 and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_clave_almacen_sobrante = rsaux!VCHA_ALM_ALMACEN_ID
               var_clave_unidad_organizacional_sobrante = rsaux!VCHA_UOR_UNIDAD_ID
            End If
            rsaux.Close
         End If
         var_clave_movimiento_ajuste_entrada = ""
         var_clave_movimiento_ajuste_salida = ""
         rsaux.Open "select vcha_mov_movimiento_id from tb_movimientos where inte_mov_ajuste_reempaque = 1 and char_mov_afectacion = '+'"
         If Not rsaux.EOF Then
            var_clave_movimiento_ajuste_entrada = rsaux!VCHA_MOV_MOVIMIENTO_ID
         End If
         rsaux.Close
         rsaux.Open "select vcha_mov_movimiento_id from tb_movimientos where inte_mov_ajuste_reempaque = 1 and char_mov_afectacion = '-'"
         If Not rsaux.EOF Then
            var_clave_movimiento_ajuste_salida = rsaux!VCHA_MOV_MOVIMIENTO_ID
         End If
         rsaux.Close
         If var_clave_movimiento_ajuste_entrada = "" Then
            MsgBox "No existe un movimiento de ajuste de entrada para reempaque", vbOKOnly, "ATENCION"
         Else
            If var_clave_movimiento_ajuste_salida = "" Then
               MsgBox "No existe un movimiento de ajuste de salida para reempaque", vbOKOnly, "ATENCION"
            Else
               If var_clave_movimiento_sobrante = "" And var_cantidad_sobrante > 0 Then
                  MsgBox "No existe un movimiento para los sobrantes", vbOKOnly, "ATENCION"
               Else
                  If var_clave_almacen_sobrante = "" And var_cantidad_sobrante > 0 Then
                     MsgBox "No se a indicado un almacen de sobrantes", vbOKOnly, "ATENCION"
                  Else
                     If var_clave_movimiento_sobrante_salida = "" And var_cantidad_sobrante > 0 Then
                        MsgBox "No se a indicado un movimiento de sobrantes de salida", vbOKOnly, "ATENCION"
                     Else
                        var_estatus_movimiento = "I"
                        Me.txt_codigo.Enabled = False
                        var_primera_vez_ajuste = 0
                        var_almacen_Destino = txt_clave_almacen_destino
                        var_almacen_origen = txt_clave_almacen_origen
                        cnn.BeginTrans
                        'SE INSERTAN LOS ARTICULOS SOBRANTES
                        If var_cantidad_sobrante > 0 Then
                           var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_clave_unidad_organizacional_sobrante, var_clave_almacen_sobrante, var_clave_movimiento_sobrante, Now, 0, 0, "", "", txt_clave_almacen_origen, var_clave_almacen_sobrante, "", var_clave_usuario_global, fun_NombrePc, "", "", txt_archivo, "", "B", "", "", 0, 0, 0, var_clave_moneda, var_tipo_Cambio)
                           var_numero_folio_sobrante = var_numero_folio_regreso
                           var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_clave_unidad_organizacional_sobrante, var_clave_almacen_sobrante, var_clave_movimiento_sobrante_salida, Now, 0, 0, "", "", var_clave_almacen_sobrante, txt_clave_almacen_origen, "", var_clave_usuario_global, fun_NombrePc, "", "", txt_archivo, "", "B", "", "", 0, 0, 0, var_clave_moneda, var_tipo_Cambio)
                           var_numero_folio_sobrante_salida = var_numero_folio_regreso
                           var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_clave_unidad_organizacional_sobrante, txt_clave_almacen_destino, var_clave_movimiento_sobrante, Now, var_numero_folio, 0, "", "", var_clave_almacen_sobrante, txt_clave_almacen_destino, "", var_clave_usuario_global, fun_NombrePc, var_factura, "", txt_archivo, "", "B", "", "", 0, 0, 0, var_clave_moneda, var_tipo_Cambio)
                           var_numero_folio_sobrante_almacen = var_numero_folio_regreso
                           Cadena = "select * from TB_REEMPAQUE_SOBRANTES where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional_origen + "'"
                           rs.Open "INSERT INTO TB_REEMPAQUE_MOVIMIENTOS_SOBRANTES (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, VCHA_REE_ALMACEN_SOBRANTE, VCHA_REE_MOVIMIENTO_SOBRANTE, INTE_REE_NUMERO_SOBRANTE, VCHA_REE_MOVIMIENTO_SOBRANTE_SALIDA, INTE_REE_NUMERO_SOBRANTE_SALIDA) Values  ('" + var_empresa + "', '" + var_unidad_organizacional_origen + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', '" + CStr(var_numero_folio) + "', '" + var_clave_almacen_sobrante + "', '" + var_clave_movimiento_sobrante + "', " + CStr(var_numero_folio_sobrante) + ",'" + var_clave_movimiento_sobrante_salida + "'," + CStr(var_numero_folio_sobrante_salida) + ")", cnn, adOpenDynamic, adLockOptimistic
                           rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                           While Not rs.EOF
                                 rsaux.Open "insert into tb_entradas ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_ENT_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_ENT_CANTIDAD], [FLOA_ENT_COSTO], [FLOA_ENT_PRECIO], [FLOA_ENT_DESCUENTO], [VCHA_ENT_ALMACEN_ORIGEN], [INTE_ENT_AÑO]) values ('" + var_empresa + "', '" + var_clave_unidad_organizacional_sobrante + "', '" + var_clave_almacen_sobrante + "', '" + var_clave_movimiento_sobrante + "', " + CStr(var_numero_folio_sobrante) + ", '" + rs!vcha_Art_Articulo_id + "', " + CStr(rs!floa_Sal_Cantidad) + ", " + CStr(rs!floa_Sal_costo) + ", " + CStr(rs!floa_Sal_precio) + ", 0,''," + CStr(rs!INTE_sAL_AÑO) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 rsaux.Open "insert into tb_salidas ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD], [FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2], [VCHA_REE_FOLIO], [VCHA_SAL_REFERENCIA], [INTE_SAL_AÑO]) values ('" + var_empresa + "', '" + var_clave_unidad_organizacional_sobrante + "', '" + var_clave_almacen_sobrante + "', '" + var_clave_movimiento_sobrante_salida + "', " + CStr(var_numero_folio_sobrante_salida) + ", '" + rs!vcha_Art_Articulo_id + "', " + CStr(rs!floa_Sal_Cantidad) + ", " + CStr(rs!floa_Sal_costo) + ", " + CStr(rs!floa_Sal_precio) + ",0, 0, 0, '', ''," + CStr(rs!INTE_sAL_AÑO) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 rsaux.Open "insert into tb_entradas ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_ENT_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_ENT_CANTIDAD], [FLOA_ENT_COSTO], [FLOA_ENT_PRECIO], [FLOA_ENT_DESCUENTO], [VCHA_ENT_ALMACEN_ORIGEN], [VCHA_ENT_REFERENCIA], VCHA_REE_FOLIO, [INTE_ENT_AÑO]) Values ('" + var_empresa + "', '" + var_clave_unidad_organizacional_sobrante + "', '" + txt_clave_almacen_destino + "','" + var_clave_movimiento_sobrante + "', " + CStr(var_numero_folio_sobrante_almacen) + ", '" + rs!vcha_Art_Articulo_id + "', " + CStr(rs!floa_Sal_Cantidad) + ", " + CStr(rs!floa_Sal_costo) + ", " + CStr(rs!floa_Sal_precio) + ", 0, '" + var_almacen_origen + "', '', '', " + CStr(rs!INTE_sAL_AÑO) + ") ", cnn, adOpenDynamic, adLockOptimistic
                                 rs.MoveNext
                           Wend
                           rs.Close
                        End If
                        Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional_origen + "'"
                        rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                        While Not rs.EOF
                              rsaux.Open "SELECT VCHA_ART_ARTICULO_ID, VCHA_REE_ARTICULO_SALIDA FROM TB_REEMPAQUE_ENTRADA WHERE VCHA_REE_ALMACEN_ORIGEN = '" + txt_clave_almacen_origen + "' AND VCHA_REE_NUMERO = '" + txt_archivo + "' AND VCHA_ART_ARTICULO_ID = '" + rs!vcha_Art_Articulo_id + "' and vcha_ree_articulo_salida = '" + rs!vcha_sal_referencia + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux.EOF Then
                                 var_codigo_salida = Trim(rsaux!VCHA_ree_ARTICULO_SALIDA)
                                 var_codigo_entrada = Trim(rsaux!vcha_Art_Articulo_id)
                              Else
                                 var_codigo_salida = ""
                                 var_codigo_entrada = ""
                              End If
                              rsaux.Close
                              ' SE INSERTAN LOS ARTICULOS QUE HAY QUE HACERLES AJUSTES
                              If var_codigo_salida <> var_codigo_entrada Then
                                 If var_primera_vez_ajuste = 0 Then
                                    var_primera_vez_ajuste = 1
                                    var_inserta = False
                                    var_numero_folio_ajuste_entrada = 0
                                    var_numero_folio_ajuste_salida = 0
                                    var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional_origen, txt_clave_almacen_origen, var_clave_movimiento_ajuste_entrada, Now, var_numero_folio, 0, "", "", "", txt_clave_almacen_origen, "", var_clave_usuario_global, fun_NombrePc, "", "", txt_archivo, "", "B", "", "", 0, 0, 0, var_clave_moneda, var_tipo_Cambio)
                                    var_numero_folio_ajuste_entrada = var_numero_folio_regreso
                                    var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional_origen, txt_clave_almacen_origen, var_clave_movimiento_ajuste_salida, Now, var_numero_folio, 0, "", "", txt_clave_almacen_origen, txt_clave_almacen_destino, "", var_clave_usuario_global, fun_NombrePc, var_factura, "", txt_archivo, "", "B", "", "", 0, 0, 0, var_clave_moneda, var_tipo_Cambio)
                                    var_numero_folio_ajuste_salida = var_numero_folio_regreso
                                    rsaux.Open "insert into TB_REEMPAQUE_MOVIMIENTOS_AJUSTE (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, VCHA_REE_MOVIMIENTO_ENTRADA, INTE_REE_NUMERO_ENTRADA, VCHA_REE_MOVIMIENTO_SALIDA, INTE_REE_NUMERO_SALIDA) values ('" + var_empresa + "', '" + var_unidad_organizacional_origen + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + var_clave_movimiento_ajuste_entrada + "', " + CStr(var_numero_folio_ajuste_entrada) + ", '" + var_clave_movimiento_ajuste_salida + "', " + CStr(var_numero_folio_ajuste_salida) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux2.Open "insert into tb_salidas ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD], [FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2], [VCHA_REE_FOLIO], [VCHA_SAL_REFERENCIA],[INTE_SAL_AÑO],[inte_sal_consecutivo_reempaque]) values ('" + var_empresa + "', '" + var_unidad_organizacional_origen + "', '" + var_almacen_origen + "', '" + var_clave_movimiento_ajuste_salida + "', " + CStr(var_numero_folio_ajuste_salida) + ", '" + var_codigo_salida + "', " + CStr(rs!floa_Sal_Cantidad) + ", " + CStr(rs!floa_Sal_costo) + ", " + CStr(rs!floa_Sal_precio) + ",0, 0, 0, '" + rs!VCHA_REE_FOLIO + "', '" + rs!vcha_Art_Articulo_id + "'," + CStr(rs!INTE_sAL_AÑO) + "," + CStr(rs!INTE_SAL_CONSECUTIVO_REEMPAQUE) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 rsaux2.Open "insert into tb_entradas ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_ENT_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_ENT_CANTIDAD], [FLOA_ENT_COSTO], [FLOA_ENT_PRECIO], [FLOA_ENT_DESCUENTO], [VCHA_ENT_ALMACEN_ORIGEN], [VCHA_ENT_REFERENCIA], [VCHA_REE_FOLIO],[INTE_ENT_AÑO],[inte_ent_consecutivo_reempaque]) Values ('" + var_empresa + "', '" + var_unidad_organizacional_origen + "', '" + var_almacen_origen + "', '" + var_clave_movimiento_ajuste_entrada + "', " + CStr(var_numero_folio_ajuste_entrada) + ", '" + rs!vcha_Art_Articulo_id + "', " + CStr(rs!floa_Sal_Cantidad) + ", " + CStr(rs!floa_Sal_costo) + ", " + CStr(rs!floa_Sal_precio) + ", " + CStr(rs!FLOA_SAL_DESCUENTO) + ", '" + var_almacen_origen + "', '" + var_codigo_salida + "', '" + rs!VCHA_REE_FOLIO + "', " + CStr(rs!INTE_sAL_AÑO) + "," + CStr(rs!INTE_SAL_CONSECUTIVO_REEMPAQUE) + ") ", cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux2.Open "insert into tb_salidas ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD], [FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2], [VCHA_REE_FOLIO], [VCHA_SAL_REFERENCIA],[INTE_SAL_AÑO], [inte_sal_consecutivo_reempaque]) values ('" + var_empresa + "', '" + var_unidad_organizacional_origen + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + rs!vcha_Art_Articulo_id + "', " + CStr(rs!floa_Sal_Cantidad) + ", " + CStr(rs!floa_Sal_costo) + ", " + CStr(rs!floa_Sal_precio) + ",0, 0, 0, '" + rs!VCHA_REE_FOLIO + "', '" + rs!vcha_sal_referencia + "', " + CStr(rs!INTE_sAL_AÑO) + "," + CStr(rs!INTE_SAL_CONSECUTIVO_REEMPAQUE) + ")", cnn, adOpenDynamic, adLockOptimistic
                              rsaux2.Open "insert into tb_entradas ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_ENT_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_ENT_CANTIDAD], [FLOA_ENT_COSTO], [FLOA_ENT_PRECIO], [FLOA_ENT_DESCUENTO], [VCHA_ENT_ALMACEN_ORIGEN], [VCHA_ENT_REFERENCIA], VCHA_REE_FOLIO, [INTE_ENT_AÑO], [inte_ent_consecutivo_reempaque]) Values ('" + var_empresa + "', '" + var_unidad_organizacional_origen + "', '" + var_almacen_Destino + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_Articulo_id + "', " + CStr(rs!floa_Sal_Cantidad) + ", " + CStr(rs!floa_Sal_costo) + ", " + CStr(rs!floa_Sal_precio) + ", " + CStr(rs!FLOA_SAL_DESCUENTO) + ", '" + var_almacen_origen + "', '" + rs!vcha_sal_referencia + "', '" + rs!VCHA_REE_FOLIO + "', " + CStr(rs!INTE_sAL_AÑO) + "," + CStr(rs!INTE_SAL_CONSECUTIVO_REEMPAQUE) + ") ", cnn, adOpenDynamic, adLockOptimistic
                              rs.MoveNext
                        Wend
                        If var_numero_folio_ajuste_entrada > 0 Then
                           rsaux.Open "update tb_encabezado_movimientos set char_emo_estatus = 'I', DTIM_EMO_FECHA_FINALIZO = GETDATE() where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento_ajuste_entrada + "' and inte_emo_numero = " + Str(var_numero_folio_ajuste_entrada) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional_origen + "'"
                        End If
                        If var_numero_folio_ajuste_salida > 0 Then
                           rsaux.Open "update tb_encabezado_movimientos set char_emo_estatus = 'I', DTIM_EMO_FECHA_FINALIZO = GETDATE() where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento_ajuste_salida + "' and inte_emo_numero = " + Str(var_numero_folio_ajuste_salida) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional_origen + "'"
                        End If
                        rsaux.Open "update tb_encabezado_movimientos set char_emo_estatus = 'I', DTIM_EMO_FECHA_FINALIZO = GETDATE() where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_emo_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional_origen + "'"
                        rs.Close
                        cnn.CommitTrans
                        
                        cnn.BeginTrans
                        rs.Open "Select max(INTE_REP_CONSECUTIVO) as consecutivo from tb_temp_reempaque_entradas", cnn, adOpenDynamic
                        If Not rs.EOF Then
                           var_consecutivo = IIf(IsNull(rs!consecutivo), 0, rs!consecutivo)
                        Else
                           var_consecutivo = 0
                        End If
                        var_consecutivo = var_consecutivo + 1
                        rs.Close
                        rs.Open "insert into tb_temp_reempaque_entradas (inte_rep_consecutivo, vcha_emp_empresa_id,vcha_rep_almacen_origen, vcha_rep_almacen_destino, vcha_mov_movimiento_id, VCHA_EMO_REFERENCIA, INTE_EMO_NUMERO) values (" + CStr(var_consecutivo) + ", '" + var_empresa + "', '" + txt_clave_almacen_origen + "', '" + txt_clave_almacen_destino + "', '" + var_clave_movimiento + "', '" + txt_archivo + "', " + CStr(var_numero_folio) + ")", cnn, adOpenDynamic, adLockOptimistic
                        Me.txt_codigo.Enabled = False
                        cnn.CommitTrans
                        rs.Open "select *  from tb_reempaque_entrada where VCHA_REE_NUMERO = " + txt_archivo + " and vcha_ree_almacen_origen = '" + txt_clave_almacen_origen + "' and vcha_ree_folio <> 'SOBRANTE'", cnn, adOpenDynamic, adLockOptimistic
                        While Not rs.EOF
                              rsaux.Open "insert into tb_temp_reempaque_entradas (inte_rep_consecutivo, vcha_emp_empresa_id,vcha_rep_almacen_origen, vcha_rep_almacen_destino, vcha_mov_movimiento_id, VCHA_EMO_REFERENCIA, INTE_EMO_NUMERO, vcha_art_articulo_id, FLOA_REP_CANTIDAD_ENVIADA, FLOA_REP_CANTIDAD_LEIDA, FLOA_REP_CANTIDAD_MOVIMIENTO, VCHA_REE_FOLIO, vcha_ree_articulo_salida, INTE_REE_CONSECUTIVO_REEMPAQUE) values (" + CStr(var_consecutivo) + ", '" + var_empresa + "', '" + txt_clave_almacen_origen + "', '" + txt_clave_almacen_destino + "', '" + var_clave_movimiento + "', '" + txt_archivo + "', " + CStr(var_numero_folio) + ", '" + rs!vcha_Art_Articulo_id + "', " + CStr(rs!FLOA_REE_CANTIDAD_ENTRADA) + ", " + CStr(rs!FLOA_REE_CANTIDAD_LEIDA) + ",0, '" + rs!VCHA_REE_FOLIO + "', '" + rs!VCHA_ree_ARTICULO_SALIDA + "'," + CStr(rs!INTE_REE_CONSECUTIVO) + ")", cnn, adOpenDynamic, adLockOptimistic
                              rs.MoveNext
                        Wend
                        rs.Close
                        Cadena = "select * from tb_temporal_SALIDAS with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional_origen + "'"
                        rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                        While Not rs.EOF
                              rsaux.Open "select * from tb_temp_reempaque_entradas where inte_rep_consecutivo = " + CStr(var_consecutivo) + " and vcha_rep_almacen_origen = '" + txt_clave_almacen_origen + "' and VCHA_EMO_REFERENCIA = '" + txt_archivo + "' and vcha_art_articulo_id = '" + rs!vcha_Art_Articulo_id + "' and vcha_ree_folio = '" + rs!VCHA_REE_FOLIO + "' and vcha_ree_folio <> 'SOBRANTE' and vcha_ree_Articulo_salida = '" + rs!vcha_sal_referencia + "' AND INTE_REE_CONSECUTIVO_REEMPAQUE = " + CStr(rs!INTE_SAL_CONSECUTIVO_REEMPAQUE), cnn, adOpenDynamic, adLockOptimistic
                              If rsaux.EOF Then
                                 rsaux2.Open "insert into tb_temp_reempaque_entradas (inte_rep_consecutivo, vcha_emp_empresa_id,vcha_rep_almacen_origen, vcha_rep_almacen_destino, vcha_mov_movimiento_id, VCHA_EMO_REFERENCIA, INTE_EMO_NUMERO, vcha_art_articulo_id, FLOA_REP_CANTIDAD_ENVIADA, FLOA_REP_CANTIDAD_LEIDA, FLOA_REP_CANTIDAD_MOVIMIENTO, vcha_ree_folio, vcha_ree_articulo_salida, INTE_REE_CONSECUTIVO_REEMPAQUE) values (" + CStr(var_consecutivo) + ", '" + var_empresa + "', '" + txt_clave_almacen_origen + "', '" + txt_clave_almacen_destino + "', '" + var_clave_movimiento + "', '" + txt_archivo + "', " + CStr(var_numero_folio) + ", '" + rs!vcha_Art_Articulo_id + "', 0, " + CStr(rs!floa_Sal_Cantidad) + "," + CStr(rs!floa_Sal_Cantidad) + ",'" + rs!VCHA_REE_FOLIO + "', '" + rs!vcha_sal_referencia + "'," + CStr(rs!INTE_SAL_CONSECUTIVO_REEMPAQUE) + ")", cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 rsaux2.Open "update tb_temp_reempaque_entradas set floa_rep_cantidad_movimiento =  floa_rep_cantidad_movimiento + " + CStr(rs!floa_Sal_Cantidad) + " where inte_rep_consecutivo = " + CStr(var_consecutivo) + " and vcha_rep_almacen_origen = '" + txt_clave_almacen_origen + "' and VCHA_EMO_REFERENCIA = '" + txt_archivo + "' and vcha_art_articulo_id = '" + rs!vcha_Art_Articulo_id + "' and vcha_ree_folio =  '" + rs!VCHA_REE_FOLIO + "' and vcha_ree_articulo_salida = '" + rs!vcha_sal_referencia + "' AND INTE_REE_CONSECUTIVO_REEMPAQUE = " + CStr(rs!INTE_SAL_CONSECUTIVO_REEMPAQUE), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux.Close
                              rs.MoveNext
                        Wend
                        rs.Close
                        rs.Open "select * from TB_REEMPAQUE_MOVIMIENTOS_AJUSTE where vcha_alm_almacen_id = '" + var_almacen_origen + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                        Set reporte = appl.OpenReport(App.Path + "\rep_reempaque_entradas.rpt")
                        reporte.RecordSelectionFormula = "{VW_REEMPAQUE_ENTRADAS.INTE_REP_CONSECUTIVO} = " + CStr(var_consecutivo)
                        frmvistasprevias.cr.ReportSource = reporte
                        For ntablas = 1 To reporte.Database.Tables.Count
                            reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                        Next ntablas
                        frmvistasprevias.cr.ViewReport
                        frmvistasprevias.Caption = "Reporte de Movimientos"
                        frmvistasprevias.Show 1
                        Set reporte = Nothing
                        rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                        rs.Close
                     End If
                  End If
               End If
            End If
         End If
      Else
         cnn.BeginTrans
         var_almacen_origen = txt_clave_almacen_origen
         rs.Open "Select max(INTE_REP_CONSECUTIVO) as consecutivo from tb_temp_reempaque_entradas", cnn, adOpenDynamic
         If Not rs.EOF Then
            var_consecutivo = IIf(IsNull(rs!consecutivo), 0, rs!consecutivo)
         Else
            var_consecutivo = 0
         End If
         var_consecutivo = var_consecutivo + 1
         rs.Close
         rs.Open "insert into tb_temp_reempaque_entradas (inte_rep_consecutivo, vcha_emp_empresa_id,vcha_rep_almacen_origen, vcha_rep_almacen_destino, vcha_mov_movimiento_id, VCHA_EMO_REFERENCIA, INTE_EMO_NUMERO) values (" + CStr(var_consecutivo) + ", '" + var_empresa + "', '" + txt_clave_almacen_origen + "', '" + txt_clave_almacen_destino + "', '" + var_clave_movimiento + "', '" + txt_archivo + "', " + CStr(var_numero_folio) + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
         rs.Open "select *  from tb_reempaque_entrada where VCHA_REE_NUMERO = " + txt_archivo + " and vcha_ree_almacen_origen = '" + txt_clave_almacen_origen + "' and vcha_ree_folio <> 'SOBRANTE'", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux.Open "insert into tb_temp_reempaque_entradas (inte_rep_consecutivo, vcha_emp_empresa_id,vcha_rep_almacen_origen, vcha_rep_almacen_destino, vcha_mov_movimiento_id, VCHA_EMO_REFERENCIA, INTE_EMO_NUMERO, vcha_art_articulo_id, FLOA_REP_CANTIDAD_ENVIADA, FLOA_REP_CANTIDAD_LEIDA, FLOA_REP_CANTIDAD_MOVIMIENTO, VCHA_REE_FOLIO, vcha_ree_articulo_salida, INTE_REE_CONSECUTIVO_REEMPAQUE) values (" + CStr(var_consecutivo) + ", '" + var_empresa + "', '" + txt_clave_almacen_origen + "', '" + txt_clave_almacen_destino + "', '" + var_clave_movimiento + "', '" + txt_archivo + "', " + CStr(var_numero_folio) + ", '" + rs!vcha_Art_Articulo_id + "', " + CStr(rs!FLOA_REE_CANTIDAD_ENTRADA) + ", " + CStr(rs!FLOA_REE_CANTIDAD_LEIDA) + ",0, '" + rs!VCHA_REE_FOLIO + "', '" + rs!VCHA_ree_ARTICULO_SALIDA + "', " + CStr(rs!INTE_REE_CONSECUTIVO) + ")", cnn, adOpenDynamic, adLockOptimistic
               rs.MoveNext
         Wend
         rs.Close
       
         Cadena = "select * from tb_temporal_SALIDAS with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional_origen + "'"
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux.Open "select * from tb_temp_reempaque_entradas where inte_rep_consecutivo = " + CStr(var_consecutivo) + " and vcha_rep_almacen_origen = '" + txt_clave_almacen_origen + "' and VCHA_EMO_REFERENCIA = '" + txt_archivo + "' and vcha_art_articulo_id = '" + rs!vcha_Art_Articulo_id + "' and vcha_ree_folio = '" + rs!VCHA_REE_FOLIO + "' and vcha_ree_articulo_salida = '" + rs!vcha_sal_referencia + "' AND INTE_REE_CONSECUTIVO_REEMPAQUE = " + CStr(rs!INTE_SAL_CONSECUTIVO_REEMPAQUE), cnn, adOpenDynamic, adLockOptimistic
               If rsaux.EOF Then
                  rsaux2.Open "insert into tb_temp_reempaque_entradas (inte_rep_consecutivo, vcha_emp_empresa_id,vcha_rep_almacen_origen, vcha_rep_almacen_destino, vcha_mov_movimiento_id, VCHA_EMO_REFERENCIA, INTE_EMO_NUMERO, vcha_art_articulo_id, FLOA_REP_CANTIDAD_ENVIADA, FLOA_REP_CANTIDAD_LEIDA, FLOA_REP_CANTIDAD_MOVIMIENTO, vcha_ree_folio, vcha_ree_articulo_salida, INTE_SAL_CONSECUTIVO_REEMPAQUE) values (" + CStr(var_consecutivo) + ", '" + var_empresa + "', '" + txt_clave_almacen_origen + "', '" + txt_clave_almacen_destino + "', '" + var_clave_movimiento + "', '" + txt_archivo + "', " + CStr(var_numero_folio) + ", '" + rs!vcha_Art_Articulo_id + "', 0, " + CStr(rs!floa_Sal_Cantidad) + "," + CStr(rs!floa_Sal_Cantidad) + ", '" + rs!VCHA_REE_FOLIO + "', '" + rs!vcha_sal_referencia + "'," + CStr(rs!INTE_SAL_CONSECUTIVO_REEMPAQUE) + ")", cnn, adOpenDynamic, adLockOptimistic
               Else
                  rsaux2.Open "update tb_temp_reempaque_entradas set floa_rep_cantidad_movimiento =  floa_rep_cantidad_movimiento + " + CStr(rs!floa_Sal_Cantidad) + " where inte_rep_consecutivo = " + CStr(var_consecutivo) + " and vcha_rep_almacen_origen = '" + txt_clave_almacen_origen + "' and VCHA_EMO_REFERENCIA = '" + txt_archivo + "' and vcha_art_articulo_id = '" + rs!vcha_Art_Articulo_id + "' and vcha_ree_folio = '" + rs!VCHA_REE_FOLIO + "' and vcha_ree_articulo_salida = '" + rs!vcha_sal_referencia + "' AND INTE_REE_CONSECUTIVO_REEMPAQUE = " + CStr(rs!INTE_SAL_CONSECUTIVO_REEMPAQUE), cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux.Close
               rs.MoveNext
         Wend
         rs.Close
                                                              
         Set reporte = appl.OpenReport(App.Path + "\rep_reempaque_entradas.rpt")
         reporte.RecordSelectionFormula = "{VW_REEMPAQUE_ENTRADAS.INTE_REP_CONSECUTIVO} = " + CStr(var_consecutivo)
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Movimientos"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
      End If
   Else
      MsgBox "No se a seleccionado un movimiento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   If var_numero_folio > 0 Then
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 1 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional_origen + "' AND VCHA_ALM_ALMACEN_ID = '" + txt_clave_almacen_origen + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
   End If
   lbl_cancelado = ""
   If var_solo_lectura = False Then
      Set TB_BLOQUEOS = New TB_BLOQUEOS
      var_global_bloqueado = 0
      ok = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, txt_archivo, Now, "", "")
   End If
   lv_entradas.ListItems.Clear
   var_primera_vez = True
   txt_destino = ""
   txt_clave_almacen_origen = ""
   txt_clave_almacen_destino = ""
   txt_nombre_almacen_origen = ""
   txt_nombre_almacen_destino = ""
   txt_folio_salida = ""
   txt_codigo = ""
   txt_clave_almacen_origen.Enabled = False
   txt_clave_almacen_destino.Enabled = False
   txt_codigo.Enabled = False
   txt_archivo = ""
   txt_archivo.Enabled = True
   var_cantidad_enviada = 0
   var_cantidad_recibida = 0
   var_numero_folio = 0
   txt_numero = ""
   lbl_recibidos = ""
   lbl_enviados = ""
   txt_folio = ""
   txt_codigo = ""
   var_estatus_movimiento = ""
   txt_archivo.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = 116 Then
      frmexisten_rapidas.Show 1
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
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If Me.frm_busqueda.Visible = True Then
         Me.frm_busqueda.Visible = False
      End If
      If Me.frm_eliminar.Visible = True Then
         Me.frm_eliminar.Visible = False
      End If
      If Me.frm_busqueda.Visible = False And False And Me.frm_eliminar.Visible = False Then
         Unload Me
      End If
   End If
End Sub

Private Sub Form_Load()
   var_numero_folio = 0
   lbl_cancelado = ""
   var_cadena_seguridad = ""
   Top = 0
   Left = 0
   Top = 0
   Left = 0
   var_ventana = 0
   var_cantidad_leida = 1#
   var_estatus_movimiento = ""
   var_almacen_Destino = ""
   var_almacen_origen = ""
   var_proveedor = ""
   var_factura = ""
   frm_eliminar.Visible = False
   var_modifica = False
   txt_cantidad.Visible = False
   lbl_cantidad.Visible = False
   txt_codigo.Enabled = False
   var_primera_vez = True
   
   frm_busqueda.Visible = False
   Set var_tabla = CreateObject("ADODB.connection")
   var_suma_cantidad_enviada = 0
   var_suma_cantidad_recibida = 0
   txt_archivo.Enabled = False
   txt_clave_almacen_destino.Enabled = False
   txt_clave_almacen_origen.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_numero_folio > 0 Then
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional_origen + "' AND VCHA_ALM_ALMACEN_ID = '" + txt_clave_almacen_origen + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
   End If
   If var_solo_lectura = False Then
   End If
   Call activa_forma(var_activa_forma_entradas_reempaque)
End Sub

Private Sub lv_entradas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         MsgBox "Imposible modificar el movimiento", vbOKOnly, "ATENCION"
      Else
         var_ventana = 2
         frm_eliminar.Visible = True
         txt_cantidad_eliminar.SetFocus
      End If
   End If
End Sub

Private Sub txt_almacen_busqueda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "SELECT * FROM TB_ALMACENES WHERE VCHA_ALM_ALMACEN_ID = '" + txt_almacen_busqueda + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_almacen_busqueda = rs!VCHA_ALM_NOMBRE
         txt_busqueda_folio = ""
         txt_busqueda_folio.SetFocus
      Else
         txt_almacen_busqueda = ""
         txt_nombre_almacen_busqueda = ""
         MsgBox "Clave de almacén incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
   If KeyAscii = 27 Then
      frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_archivo_KeyPress(KeyAscii As Integer)
   
   On Error GoTo salir:
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      rs.Open "select VCHA_PRI_RUTA_REEMPAQUE from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_ruta = IIf(IsNull(rs!VCHA_PRI_RUTA_REEMPAQUE), "", rs!VCHA_PRI_RUTA_REEMPAQUE)
      Else
         var_ruta = ""
      End If
      rs.Close
      If Trim(var_ruta) <> "" Then
         VAR_MAQUINA = fun_NombrePc
         'MsgBox "1"
         If Not UCase(VAR_MAQUINA) = "JFSERNA" Then
            var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
         Else
            If var_tabla.State = 1 Then
               var_tabla.Close
            End If
            var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=DSN=dBASE Files;DBQ=" & var_ruta & ";DefaultDir=" & var_ruta & ";DriverId=533;MaxBufferSize=2048;PageTimeout=5;"
            'var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
            'var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver={Microsoft FoxPro VFP Driver (*.dbf)};UID=;SourceDB= " & var_ruta + ";SourceType=DBF;"

         End If
         'MsgBox var_tabla.ConnectionString
         
         rs.Open "SELECT NUMNOTA,CODIGO,COSTO,CANT1, CODIGO_ENT ,FOLIO_ORIG AS FOLIO,FOLIO_DIF, FECHA_DOC, PLANTA_ID,TIPO, ANOCOSTO  FROM " + txt_archivo, var_tabla, adOpenDynamic, adLockOptimistic
         'MsgBox "3"
         If Not rs.EOF Then
            var_numero_entrada = rs!numnota
            var_almacen_origen = Trim(rs!planta_id)
            If var_almacen_origen = "4" Then
               var_almacen_origen = "6"
            End If
            If rsaux3.State = 1 Then
               rsaux3.Close
            End If
            
            rsaux3.Open "select * from TB_REEMPAQUE_ENTRADA where VCHA_REE_NUMERO = " + CStr(var_numero_entrada) + " and vcha_ree_almacen_origen = '" + var_almacen_origen + "'  ORDER BY VCHA_REE_FOLIO, VCHA_ART_ARTICULO_ID, INTE_REE_AÑO", cnn, adOpenDynamic, adLockOptimistic
            'MsgBox "select * from TB_REEMPAQUE_ENTRADA where VCHA_REE_NUMERO = " + CStr(var_numero_entrada) + " and vcha_ree_almacen_origen = '" + var_almacen_origen + "'  ORDER BY VCHA_REE_FOLIO, VCHA_ART_ARTICULO_ID, INTE_REE_AÑO"
            If rsaux3.EOF Then
               var_folio_entrada = rs!FOLIO
               var_tipo_entrada = rs!tipo
               txt_folio_salida = rs!FOLIO
               rsaux2.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  txt_clave_almacen_origen = rsaux2!VCHA_ALM_ALMACEN_ID
                  txt_nombre_almacen_origen = rsaux2!VCHA_ALM_NOMBRE
                  rsaux.Open "select vcha_uor_unidad_id from tb_almacenes where vcha_alm_almacen_id = '" + txt_clave_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_unidad_organizacional_origen = rsaux!VCHA_UOR_UNIDAD_ID
                  rsaux.Close
                  rsaux2.Close
                  var_consecutivo = 0
                  While Not rs.EOF
                        rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs!codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           var_codigo = rsaux2!vcha_Art_Articulo_id
                        Else
                           rsaux4.Open "select * from tb_equivalencias where VCHA_EQU_CODIGO_EQUIVALENTE = '" + rs!codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux4.EOF Then
                              var_codigo = rsaux4!vcha_Art_Articulo_id
                           Else
                              var_codigo = rs!codigo
                           End If
                           rsaux4.Close
                        End If
                        rsaux2.Close
                        
                        
                        rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs!codigo_ent + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           var_codigo_entrada = rsaux2!vcha_Art_Articulo_id
                        Else
                           rsaux4.Open "select * from tb_equivalencias where VCHA_EQU_CODIGO_EQUIVALENTE = '" + rs!codigo_ent + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux4.EOF Then
                              var_codigo_entrada = rsaux4!vcha_Art_Articulo_id
                           Else
                              var_codigo_entrada = rs!codigo_ent
                           End If
                           rsaux4.Close
                        End If
                        rsaux2.Close
                        
                        
                        rsaux2.Open "SELECT * FROM TB_REEMPAQUE_ENTRADA WHERE VCHA_REE_FOLIO = '" + rs!FOLIO + "' AND VCHA_REE_NUMERO = " + CStr(rs!numnota) + " AND VCHA_ART_ARTICULO_ID = '" + rs!codigo + "' and vcha_ree_articulo_salida = '" + rs!codigo_ent + "' and inte_REE_AÑO = " + Trim(rs!ANOCOSTO), cnn, adOpenDynamic, adLockOptimistic
                        If rsaux2.EOF Then
                           rsaux.Open "select max(inte_ree_consecutivo) from tb_reempaque_entrada where VCHA_REE_FOLIO = '" + rs!FOLIO + "' and INTE_REE_TIPO_ENTRADA = " + CStr(var_tipo_entrada) + " and VCHA_REE_ALMACEN_ORIGEN = '" + var_almacen_origen + "' and VCHA_ALM_ALMACEN_ID = '" + txt_clave_almacen_destino + "' and VCHA_REE_NUMERO = " + CStr(rs!numnota), cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux.EOF Then
                              var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                           Else
                              var_consecutivo = 0
                           End If
                           rsaux.Close
                           var_consecutivo = var_consecutivo + 1
                           rsaux.Open "insert into TB_REEMPAQUE_ENTRADA ([VCHA_REE_FOLIO], [INTE_REE_TIPO_ENTRADA], [VCHA_REE_ALMACEN_ORIGEN], [VCHA_ALM_ALMACEN_ID], [VCHA_REE_NUMERO],[VCHA_REE_ARTICULO_SALIDA] ,[VCHA_ART_ARTICULO_ID], [FLOA_REE_COSTO_ENTRADA], [FLOA_REE_CANTIDAD_ENTRADA], [FLOA_REE_CANTIDAD_LEIDA], [INTE_REE_AÑO], [INTE_REE_CONSECUTIVO])  values ('" + rs!FOLIO + "'," + CStr(var_tipo_entrada) + " ,'" + var_almacen_origen + "', '" + txt_clave_almacen_destino + "', " + CStr(rs!numnota) + ", '" + Trim(var_codigo_entrada) + "','" + var_codigo + "', " + CStr(rs!Costo) + ", " + CStr(rs!cant1) + ", 0, " + rs!ANOCOSTO + "," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux2.Close
                        var_consecutivo = var_consecutivo + 1
                        rs.MoveNext
                  Wend
                  
                  rsaux2.Open "select * from TB_REEMPAQUE_ENTRADA where VCHA_REE_NUMERO = " + CStr(var_numero_entrada) + " and vcha_ree_almacen_origen = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                  lv_entradas.ListItems.Clear
                  var_suma_cantidad_enviada = 0
                  var_suma_cantidad_recibida = 0
                  While Not rsaux2.EOF
                        Set list_item = lv_entradas.ListItems.Add(, , Trim(rsaux2!VCHA_REE_FOLIO))
                        list_item.SubItems(1) = IIf(IsNull(rsaux2!vcha_Art_Articulo_id), "", rsaux2!vcha_Art_Articulo_id)
                        rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + Trim(rsaux2!vcha_Art_Articulo_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                        list_item.SubItems(2) = IIf(IsNull(rsaux!vcha_Art_nombre_español), "", rsaux!vcha_Art_nombre_español)
                        list_item.SubItems(3) = Format(IIf(IsNull(rsaux2!FLOA_REE_CANTIDAD_ENTRADA), 0, rsaux2!FLOA_REE_CANTIDAD_ENTRADA), "###,###,##0.00")
                        list_item.SubItems(4) = Format(IIf(IsNull(rsaux2!FLOA_REE_CANTIDAD_LEIDA), 0, rsaux2!FLOA_REE_CANTIDAD_LEIDA), "###,###,##0.00")
                        list_item.SubItems(5) = Format(0, "###,###,##0.00")
                        list_item.SubItems(6) = Format(list_item.SubItems(3) - list_item.SubItems(4), "###,###,##0.00")
                        list_item.SubItems(7) = IIf(IsNull(rsaux2!FLOA_REE_COSTO_ENTRADA), "", rsaux2!FLOA_REE_COSTO_ENTRADA)
                        list_item.SubItems(8) = IIf(IsNull(rsaux!mone_Art_precio_base), "", rsaux!mone_Art_precio_base)
                        list_item.SubItems(9) = IIf(IsNull(rsaux2!VCHA_ree_ARTICULO_SALIDA), "", rsaux2!VCHA_ree_ARTICULO_SALIDA)
                        list_item.SubItems(10) = IIf(IsNull(rsaux2!inte_ree_año), "", rsaux2!inte_ree_año)
                        list_item.SubItems(11) = IIf(IsNull(rsaux2!INTE_REE_CONSECUTIVO), "", rsaux2!INTE_REE_CONSECUTIVO)
                        rsaux.Close
                        var_suma_cantidad_enviada = Format(var_suma_cantidad_enviada + rsaux2!FLOA_REE_CANTIDAD_ENTRADA, "###,###,##0.00")
                        var_suma_cantidad_recibida = Format(var_suma_cantidad_recibida + rsaux2!FLOA_REE_CANTIDAD_LEIDA, "###,###,##0.00")
                        rsaux2.MoveNext
                  Wend
                  lbl_enviados = Format(Str(var_suma_cantidad_enviada), "###,###,##0.00")
                  lbl_recibidos = Format(Str(var_suma_cantidad_recibida), "###,###,##0.00")
                  rsaux2.Close
                  txt_archivo.Enabled = False
                  If var_tipo_entrada > 0 Then
                     If var_tipo_entrada = 1 Then
                        rsaux.Open "SELECT * FROM TB_ALMACENES WHERE INTE_ALM_TIPO_ENTRADA_REEMPAQUE = 1 AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                        txt_clave_almacen_destino = rsaux!VCHA_ALM_ALMACEN_ID
                        txt_nombre_almacen_destino = rsaux!VCHA_ALM_NOMBRE
                        rsaux.Close
                        rsaux.Open "update tb_reempaque_entrada set vcha_alm_almacen_id = '" + txt_clave_almacen_destino + "' where VCHA_REE_NUMERO = " + txt_archivo + " and vcha_ree_almacen_origen = '" + txt_clave_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                        txt_clave_almacen_destino.Enabled = False
                     End If
                     If var_tipo_entrada = 2 Then
                        rsaux.Open "SELECT * FROM TB_ALMACENES WHERE INTE_ALM_TIPO_ENTRADA_REEMPAQUE = 2 AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                        txt_clave_almacen_destino = rsaux!VCHA_ALM_ALMACEN_ID
                        txt_nombre_almacen_destino = rsaux!VCHA_ALM_NOMBRE
                        rsaux.Close
                        rsaux.Open "update tb_reempaque_entrada set vcha_alm_almacen_id = '" + txt_clave_almacen_destino + "' where VCHA_REE_NUMERO = " + txt_archivo + " and vcha_ree_almacen_origen = '" + txt_clave_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                        txt_clave_almacen_destino.Enabled = False
                     End If
                     txt_codigo.Enabled = True
                     txt_codigo.SetFocus
                  Else
                     txt_clave_almacen_destino.Enabled = True
                     txt_clave_almacen_destino.SetFocus
                  End If
               Else
                  rsaux2.Close
                  MsgBox "Clave del almacen origen incorrecta", vbOKOnly, "ATENCION"
               End If
            Else
               txt_clave_almacen_origen = rsaux3!vcha_ree_almacen_origen
               rsaux.Open "select vcha_uor_unidad_id from tb_almacenes where vcha_alm_almacen_id = '" + txt_clave_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               var_unidad_organizacional_origen = rsaux!VCHA_UOR_UNIDAD_ID
               rsaux.Close
               rsaux.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + txt_clave_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               txt_nombre_almacen_origen = rsaux!VCHA_ALM_NOMBRE
               rsaux.Close
               txt_clave_almacen_destino = rsaux3!VCHA_ALM_ALMACEN_ID
               rsaux.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + txt_clave_almacen_destino + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  txt_nombre_almacen_destino = rsaux!VCHA_ALM_NOMBRE
               End If
               rsaux.Close
               txt_clave_almacen_destino.Enabled = False
               txt_clave_almacen_origen.Enabled = False
               txt_archivo.Enabled = False
               txt_codigo.Enabled = True
               var_folio_entrada = rsaux3!VCHA_REE_FOLIO
               var_tipo_entrada = IIf(IsNull(rsaux3!INTE_REE_TIPO_ENTRADA), 0, rsaux3!INTE_REE_TIPO_ENTRADA)
               lv_entradas.ListItems.Clear
               var_suma_cantidad_enviada = 0
               var_suma_cantidad_recibida = 0
               While Not rsaux3.EOF
                     Set list_item = lv_entradas.ListItems.Add(, , Trim(rsaux3!VCHA_REE_FOLIO))
                     'MsgBox rsaux3!vcha_art_Articulo_id
                     If rsaux2.State = 1 Then
                        rsaux2.Close
                     End If
                     rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + Trim(rsaux3!vcha_Art_Articulo_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        list_item.SubItems(2) = IIf(IsNull(rsaux2!vcha_Art_nombre_español), "", rsaux2!vcha_Art_nombre_español)
                        list_item.SubItems(8) = IIf(IsNull(rsaux2!mone_Art_precio_base), "", rsaux2!mone_Art_precio_base)
                        list_item.SubItems(1) = IIf(IsNull(rsaux3!vcha_Art_Articulo_id), "", rsaux3!vcha_Art_Articulo_id)
                     Else
                        'MsgBox "select * from tb_Articulos where substring(vcha_Art_articulo_id,7,5) = " + Trim(rsaux3!vcha_art_Articulo_id)
                        rsaux4.Open "select * from tb_Articulos where substring(vcha_Art_articulo_id,7,5) = '" + Trim(rsaux3!vcha_Art_Articulo_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux4.EOF Then
                           list_item.SubItems(2) = IIf(IsNull(rsaux4!vcha_Art_nombre_español), "", rsaux4!vcha_Art_nombre_español)
                           list_item.SubItems(8) = IIf(IsNull(rsaux4!mone_Art_precio_base), "", rsaux4!mone_Art_precio_base)
                           list_item.SubItems(1) = IIf(IsNull(rsaux4!vcha_Art_Articulo_id), "", rsaux4!vcha_Art_Articulo_id)
                        End If
                        rsaux4.Close
                     End If
                     list_item.SubItems(3) = Format(IIf(IsNull(rsaux3!FLOA_REE_CANTIDAD_ENTRADA), 0, rsaux3!FLOA_REE_CANTIDAD_ENTRADA), "###,###,##0.00")
                     list_item.SubItems(4) = Format(IIf(IsNull(rsaux3!FLOA_REE_CANTIDAD_LEIDA), 0, rsaux3!FLOA_REE_CANTIDAD_LEIDA), "###,###,##0.00")
                     list_item.SubItems(5) = Format(0, "###,###,##0.00")
                     list_item.SubItems(6) = Format(list_item.SubItems(3) - list_item.SubItems(4), "###,###,##0.00")
                     list_item.SubItems(7) = IIf(IsNull(rsaux3!FLOA_REE_COSTO_ENTRADA), "", rsaux3!FLOA_REE_COSTO_ENTRADA)
                     list_item.SubItems(9) = IIf(IsNull(rsaux3!VCHA_ree_ARTICULO_SALIDA), "", rsaux3!VCHA_ree_ARTICULO_SALIDA)
                     list_item.SubItems(10) = IIf(IsNull(rsaux3!inte_ree_año), "", rsaux3!inte_ree_año)
                     list_item.SubItems(11) = IIf(IsNull(rsaux3!INTE_REE_CONSECUTIVO), "", rsaux3!INTE_REE_CONSECUTIVO)
                     rsaux2.Close
                     var_suma_cantidad_enviada = Format(var_suma_cantidad_enviada + rsaux3!FLOA_REE_CANTIDAD_ENTRADA, "###,###,##0.00")
                     var_suma_cantidad_recibida = Format(var_suma_cantidad_recibida + IIf(IsNull(rsaux3!FLOA_REE_CANTIDAD_LEIDA), 0, FLOA_REE_CANTIDAD_LEIDA), "###,###,##0.00")
                     rsaux3.MoveNext
               Wend
               lbl_enviados = Format(Str(var_suma_cantidad_enviada), "###,###,##0.00")
               lbl_recibidos = Format(Str(var_suma_cantidad_recibida), "###,###,##0.00")
               txt_archivo.Enabled = False
               If var_tipo_entrada > 0 Then
                  If var_tipo_entrada = 1 Then
                     rsaux.Open "SELECT * FROM TB_ALMACENES WHERE INTE_ALM_TIPO_ENTRADA_REEMPAQUE = 1 AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                     txt_clave_almacen_destino = rsaux!VCHA_ALM_ALMACEN_ID
                     txt_nombre_almacen_destino = rsaux!VCHA_ALM_NOMBRE
                     rsaux.Close
                     rsaux.Open "update tb_reempaque_entrada set vcha_alm_almacen_id = '" + txt_clave_almacen_destino + "' where VCHA_REE_NUMERO = " + txt_archivo + " and vcha_ree_almacen_origen = '" + txt_clave_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                     txt_clave_almacen_destino.Enabled = False
                  End If
                  If var_tipo_entrada = 2 Then
                     rsaux.Open "SELECT * FROM TB_ALMACENES WHERE INTE_ALM_TIPO_ENTRADA_REEMPAQUE = 2 AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                     txt_clave_almacen_destino = rsaux!VCHA_ALM_ALMACEN_ID
                     txt_nombre_almacen_destino = rsaux!VCHA_ALM_NOMBRE
                     rsaux.Close
                     rsaux.Open "update tb_reempaque_entrada set vcha_alm_almacen_id = '" + txt_clave_almacen_destino + "' where VCHA_REE_NUMERO = " + txt_archivo + " and vcha_ree_almacen_origen = '" + txt_clave_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                     txt_clave_almacen_destino.Enabled = False
                  End If
                  txt_codigo.Enabled = True
                  txt_codigo.SetFocus
               End If
            End If
            rsaux3.Close
         Else
            MsgBox "El archivo no tiene información", vbOKOnly, "ATENCION"
         End If
         rs.Close
         var_tabla.Close
      Else
         MsgBox "No se a indicado una ruta para el archivo enviado por la planta", vbOKOnly, "ATENCION"
      End If
   End If
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If var_tabla.State = 1 Then
      var_tabla.Close
   End If
  
Exit Sub
salir:
   MsgBox "A surgido un error al leer el archivo. Puede ser que el archivo " + Trim(txt_archivo) + " no existe", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If var_tabla.State = 1 Then
      var_tabla.Close
   End If
   txt_archivo.Enabled = False
   txt_clave_almacen_origen.Enabled = True
   txt_clave_almacen_origen.SetFocus
   Exit Sub
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txt_almacen_busqueda) <> "" Then
         If IsNumeric(txt_busqueda_folio) Then
            If var_numero_folio = CDbl(txt_busqueda_folio) And txt_clave_almacen_origen = txt_almacen_busqueda Then
               rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional_origen + "' AND VCHA_ALM_ALMACEN_ID = '" + txt_clave_almacen_origen + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
            End If
            rs.Open "SELECT * FROM TB_ENCABEZADO_MOVIMIENTOS WHERE VCHA_ALM_ALMACEN_ID = '" + txt_almacen_busqueda + "' AND INTE_EMO_NUMERO = " + txt_busqueda_folio + " AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               If var_numero_folio > 0 Then
                  rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional_origen + "' AND VCHA_ALM_ALMACEN_ID = '" + txt_clave_almacen_origen + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
               End If
               var_movimiento_bloqueado = IIf(IsNull(rs!INTE_EMO_BLOQUEADO), 0, rs!INTE_EMO_BLOQUEADO)
               If var_movimiento_bloqueado = 0 Then
                  Dim var_codigo_salida As String
                  lv_entradas.ListItems.Clear
                  var_primera_vez = False
                  var_estatus_movimiento = IIf(IsNull(rs!char_Emo_estatus), "", rs!char_Emo_estatus)
                  var_cantidad_recibida = 0
                  var_numero_folio = rs!INTE_EMO_NUMERO
                  Me.txt_folio = var_numero_folio
                  txt_archivo = rs!vcha_Emo_referencia
                  lbl_recibidos = "0"
                  txt_clave_almacen_origen = rs!vcha_emo_almacen_origen
                  rsaux.Open "SELECT * FROM TB_ALMACENES WHERE VCHA_ALM_ALMACEN_ID = '" + txt_clave_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_nombre_almacen_origen = rsaux!VCHA_ALM_NOMBRE
                  rsaux.Close
                  txt_clave_almacen_destino = rs!VCHA_EMO_ALMACEN_DESTINO
                  rsaux.Open "SELECT * FROM TB_ALMACENES WHERE VCHA_ALM_ALMACEN_ID = '" + txt_clave_almacen_destino + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_nombre_almacen_destino = rsaux!VCHA_ALM_NOMBRE
                  rsaux.Close
                  var_unidad_organizacional_origen = rs!VCHA_UOR_UNIDAD_ID
                  'MsgBox "SELECT * FROM TB_REEMPAQUE_ENTRADA WHERE VCHA_REE_ALMACEN_ORIGEN = '" + txt_clave_almacen_origen + "' AND VCHA_REE_NUMERO = '" + txt_archivo + "' ORDER BY VCHA_REE_FOLIO, VCHA_ART_ARTICULO_ID, INTE_REE_AÑO"
                  rsaux.Open "SELECT * FROM TB_REEMPAQUE_ENTRADA WHERE VCHA_REE_ALMACEN_ORIGEN = '" + txt_clave_almacen_origen + "' AND VCHA_REE_NUMERO = '" + txt_archivo + "' ORDER BY VCHA_REE_FOLIO, VCHA_ART_ARTICULO_ID, INTE_REE_AÑO", cnn, adOpenDynamic, adLockOptimistic
                  var_cantidad_enviada = 0
                  While Not rsaux.EOF
                        Set list_item = lv_entradas.ListItems.Add(, , Trim(rsaux!VCHA_REE_FOLIO))
                        list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_Art_Articulo_id), "", rsaux!vcha_Art_Articulo_id)
                        rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        list_item.SubItems(2) = IIf(IsNull(rsaux2!vcha_Art_nombre_español), "", rsaux2!vcha_Art_nombre_español)
                        list_item.SubItems(3) = Format(IIf(IsNull(rsaux!FLOA_REE_CANTIDAD_ENTRADA), 0, rsaux!FLOA_REE_CANTIDAD_ENTRADA), "###,###,##0.00")
                        list_item.SubItems(4) = Format(IIf(IsNull(rsaux!FLOA_REE_CANTIDAD_LEIDA), 0, rsaux!FLOA_REE_CANTIDAD_LEIDA), "###,###,##0.00")
                        list_item.SubItems(5) = Format(0, "###,###,##0.00")
                        list_item.SubItems(6) = Format(list_item.SubItems(3) - list_item.SubItems(4), "###,###,##0.00")
                        list_item.SubItems(7) = IIf(IsNull(rsaux!FLOA_REE_COSTO_ENTRADA), "", rsaux!FLOA_REE_COSTO_ENTRADA)
                        list_item.SubItems(8) = IIf(IsNull(rsaux2!mone_Art_precio_base), "", rsaux2!mone_Art_precio_base)
                        list_item.SubItems(9) = IIf(IsNull(rsaux!VCHA_ree_ARTICULO_SALIDA), "", rsaux!VCHA_ree_ARTICULO_SALIDA)
                        list_item.SubItems(10) = IIf(IsNull(rsaux!inte_ree_año), "", rsaux!inte_ree_año)
                        list_item.SubItems(11) = IIf(IsNull(rsaux!INTE_REE_CONSECUTIVO), 0, rsaux!INTE_REE_CONSECUTIVO)
                        rsaux2.Close
                        var_cantidad_enviada = var_cantidad_enviada + IIf(IsNull(rsaux!FLOA_REE_CANTIDAD_ENTRADA), 0, rsaux!FLOA_REE_CANTIDAD_ENTRADA)
                        rsaux.MoveNext
                  Wend
                  lbl_enviados = Format(Str(var_cantidad_enviada), "###,###,##0.00")
                  rsaux.Close
                  Cadena = "select * from tb_temporal_SALIDAS with (nolock) where vcha_alm_almacen_id = '" + txt_clave_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional_origen + "'"
                  rsaux.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux.EOF
                        If rsaux2.State = 1 Then
                           rsaux2.Close
                        End If
                        var_n = lv_entradas.ListItems.Count
                        var_i = 1
                        While var_i <= var_n
                            lv_entradas.ListItems.Item(var_i).Selected = True
                            var_folio_salida = rsaux!VCHA_REE_FOLIO
                            var_codigo_salida = rsaux!vcha_sal_referencia
                            'MsgBox CStr(rsaux!inte_sal_consecutivo_reempaque)
                            If Trim(rsaux!vcha_Art_Articulo_id) = Trim(lv_entradas.selectedItem.SubItems(1)) And Trim(lv_entradas.selectedItem) = var_folio_salida And Trim(lv_entradas.selectedItem.SubItems(9)) = var_codigo_salida And (lv_entradas.selectedItem.SubItems(10) * 1) = rsaux!INTE_sAL_AÑO And (lv_entradas.selectedItem.SubItems(11) * 1) = IIf(IsNull(rsaux!INTE_SAL_CONSECUTIVO_REEMPAQUE), 0, rsaux!INTE_SAL_CONSECUTIVO_REEMPAQUE) Then
                               lv_entradas.selectedItem.SubItems(5) = Format(lv_entradas.selectedItem.SubItems(5) + rsaux!floa_Sal_Cantidad, "###,###,##0.00")
                               var_cantidad = lv_entradas.selectedItem.SubItems(4)
                               lbl_recibidos = Format(Int(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                               var_cantidad_recibida = var_cantidad_recibida + rsaux!floa_Sal_Cantidad
                            End If
                            var_i = var_i + 1
                        Wend
                        rsaux.MoveNext
                  Wend
                  lbl_recibidos = Format(Str(var_cantidad_recibida), "###,###,##0.00")
                  rsaux.Close
                  
                  rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 1 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional_origen + "' AND VCHA_ALM_ALMACEN_ID = '" + txt_clave_almacen_origen + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  
                  Cadena = "select * from TB_REEMPAQUE_SOBRANTES where vcha_alm_almacen_id = '" + txt_clave_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional_origen + "'"
                  rsaux.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux.EOF
                        If rsaux2.State = 1 Then
                           rsaux2.Close
                        End If
                        var_n = lv_entradas.ListItems.Count
                        var_i = 1
                        While var_i <= var_n
                            lv_entradas.ListItems.Item(var_i).Selected = True
                            If Trim(rsaux!vcha_Art_Articulo_id) = Trim(lv_entradas.selectedItem.SubItems(1)) And Trim(lv_entradas.selectedItem) = "SOBRANTE" Then
                               lv_entradas.selectedItem.SubItems(5) = Format(lv_entradas.selectedItem.SubItems(5) + rsaux!floa_Sal_Cantidad, "###,###,##0.00")
                               var_cantidad = lv_entradas.selectedItem.SubItems(4)
                               lbl_recibidos = Format(Int(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                               var_cantidad_recibida = var_cantidad_recibida + rsaux!floa_Sal_Cantidad
                            End If
                            var_i = var_i + 1
                        Wend
                        rsaux.MoveNext
                  Wend
                  lbl_recibidos = Format(Str(var_cantidad_recibida), "###,###,##0.00")
                  rsaux.Close
                  If Trim(var_estatus_movimiento) = "" Then
                     txt_codigo.Enabled = True
                  Else
                     txt_codigo.Enabled = False
                  End If
                  frm_busqueda.Visible = False
               Else
                  MsgBox "El movimiento esta siendo utilizado por otro usuario", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
               frm_busqueda.Visible = False
            End If
            rs.Close
         Else
            MsgBox "Número de movimiento incorrecto", vbOKOnly, "ATENCION"
            frm_busqueda.Visible = False
         End If
      Else
         MsgBox "No se a seleccionado un almacen", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_cantidad_eliminar_GotFocus()
   txt_cantidad_eliminar = ""
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
      Set TB_ARCH_COMPARACION_M = New TB_ARCH_COMPARACION_M
      Dim var_posible_eliminar As Double
      var_cantidad_eliminar = Val(txt_cantidad_eliminar)
      var_cantidad_eliminar_mov = lv_entradas.selectedItem.SubItems(5)
      var_almacen_origen = txt_clave_almacen_origen
      var_posible_eliminar = var_cantidad_eliminar_mov - var_cantidad_eliminar
      var_año = lv_entradas.selectedItem.SubItems(10)
      If var_posible_eliminar >= 0 Then
         If var_cantidad_eliminar_mov <= var_cantidad_eliminar Then
            MsgBox "No esposible eliminar esta cantidad", vbOKOnly, "ATENCION"
         Else
            var_inserta = False
            var_folio_salida = lv_entradas.selectedItem
            If var_folio_salida = "SOBRANTE" Then
               rsaux2.Open "update TB_REEMPAQUE_SOBRANTES set floa_sal_cantidad = floa_sal_cantidad - " + CStr(var_cantidad_eliminar) + " where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'  and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional_origen + "' and  VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and INTE_SAL_NUMERO =" + CStr(var_numero_folio) + " and VCHA_ART_ARTICULO_ID = '" + lv_entradas.selectedItem.SubItems(1) + "' AND INTE_SAL_AÑO = " + CStr(var_año), cnn, adOpenDynamic, adLockOptimistic
            Else
               rsaux2.Open "update tb_Temporal_salidas set floa_sal_cantidad = floa_sal_cantidad - " + CStr(var_cantidad_eliminar) + " where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'  and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional_origen + "' and  VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and INTE_SAL_NUMERO =" + CStr(var_numero_folio) + " and VCHA_ART_ARTICULO_ID = '" + lv_entradas.selectedItem.SubItems(1) + "' and VCHA_REE_FOLIO = '" + var_folio_salida + "' and vcha_ree_folio = '" + var_folio_salida + "' AND INTE_SAL_AÑO = " + CStr(var_año) + " and inte_sal_consecutivo_reempaque = " + Me.lv_entradas.selectedItem.SubItems(11), cnn, adOpenDynamic, adLockOptimistic
            End If
            rsaux3.Open "update TB_REEMPAQUE_ENTRADA  set FLOA_REE_CANTIDAD_LEIDA = FLOA_REE_CANTIDAD_LEIDA - " + CStr(var_cantidad_eliminar) + " where VCHA_REE_FOLIO = '" + var_folio_salida + "' and VCHA_REE_ALMACEN_ORIGEN = '" + txt_clave_almacen_origen + "' and VCHA_REE_NUMERO = '" + txt_archivo + "' and VCHA_ART_ARTICULO_ID = '" + lv_entradas.selectedItem.SubItems(1) + "' AND INTE_REE_AÑO = " + CStr(var_año) + " and inte_ree_consecutivo = " + Me.lv_entradas.selectedItem.SubItems(11), cnn, adOpenDynamic, adLockOptimistic
            lbl_recibidos = Int(lbl_recibidos) - var_cantidad_eliminar
            frm_eliminar.Visible = False
            txt_codigo.SetFocus
            lv_entradas.selectedItem.SubItems(4) = Format((lv_entradas.selectedItem.SubItems(4) * 1) - var_cantidad_eliminar, "###,###,##0.00")
            lv_entradas.selectedItem.SubItems(5) = Format((lv_entradas.selectedItem.SubItems(5) * 1) - var_cantidad_eliminar, "###,###,##0.00")
            lv_entradas.selectedItem.SubItems(6) = Format((lv_entradas.selectedItem.SubItems(6) * 1) + var_cantidad_eliminar, "###,###,##0.00")
            var_ventana = 0
            var_renglon = lv_entradas.selectedItem.Index
            Call ilumina_grid
         End If
      Else
            MsgBox "No esposible eliminar esta cantidad", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      frm_eliminar.Visible = False
      txt_codigo.SetFocus
      var_ventana = 0
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   frm_eliminar.Visible = False
   txt_codigo.SetFocus
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
      var_cantidad_leida = CInt(txt_cantidad)
      txt_foco.SetFocus
   End If
End Sub

Private Sub txt_cantidad_LostFocus()
   Me.txt_cantidad.Visible = False
End Sub

Private Sub txt_clave_almacen_destino_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_clave_almacen_destino) <> "" Then
         If var_tipo_permiso = 1 Then
            rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_alm_almacen_id = '" + txt_clave_almacen_destino + "'order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
            If Not rs.EOF Then
               rsaux.Open "update tb_reempaque_entrada set vcha_alm_almacen_id = '" + txt_clave_almacen_destino + "' where VCHA_REE_NUMERO = " + txt_archivo + " and vcha_ree_almacen_origen = '" + txt_clave_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               txt_nombre_almacen_destino = rs!VCHA_ALM_NOMBRE
               txt_clave_almacen_destino.Enabled = False
               txt_codigo.Enabled = True
               txt_codigo.SetFocus
            Else
               txt_clave_almacen_destino.Enabled = True
               txt_clave_almacen_destino = ""
               txt_nombre_almacen_destino = ""
               MsgBox "Clave de almacen incorrecta", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            rs.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + txt_clave_almacen_destino + "'", cnn, adOpenDynamic, adLockBatchOptimistic
            If Not rs.EOF Then
               rsaux.Open "update tb_reempaque_entrada set vcha_alm_almacen_id = '" + txt_clave_almacen_destino + "' where VCHA_REE_NUMERO = " + txt_archivo + " and vcha_ree_almacen_origen = '" + txt_clave_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               txt_nombre_almacen_destino = rs!VCHA_ALM_NOMBRE
               txt_clave_almacen_destino.Enabled = False
               txt_codigo.Enabled = True
               txt_codigo.SetFocus
            Else
               txt_clave_almacen_destino.Enabled = True
               txt_clave_almacen_destino = ""
               txt_nombre_almacen_destino = ""
               MsgBox "Clave de almacen incorrecta", vbOKOnly, "ATENCION"
            End If
            rs.Close
         End If
      End If
   End If
End Sub

Private Sub txt_clave_almacen_origen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_clave_almacen_origen) <> "" Then
         rs.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + txt_clave_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux3.Open "select vcha_uor_unidad_id from tb_almacenes where vcha_alm_almacen_id = '" + txt_clave_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
            var_unidad_organizacional_origen = rsaux3!VCHA_UOR_UNIDAD_ID
            rsaux3.Close
            rsaux2.Open "select * from TB_REEMPAQUE_ENTRADA where VCHA_REE_NUMERO = " + txt_archivo + " and vcha_ree_almacen_origen = '" + txt_clave_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               txt_nombre_almacen_origen = rs!VCHA_ALM_NOMBRE
               lv_entradas.ListItems.Clear
               var_suma_cantidad_enviada = 0
               var_suma_cantidad_recibida = 0
               txt_folio_salida = rsaux2!VCHA_REE_FOLIO
               While Not rsaux2.EOF
                     Set list_item = lv_entradas.ListItems.Add(, , Trim(rsaux2!vcha_Art_Articulo_id))
                     rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux2!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_Art_nombre_español), "", rsaux!vcha_Art_nombre_español)
                     list_item.SubItems(2) = Format(IIf(IsNull(rsaux2!FLOA_REE_CANTIDAD_ENTRADA), 0, rsaux2!FLOA_REE_CANTIDAD_ENTRADA), "###,###,##0.00")
                     list_item.SubItems(3) = Format(IIf(IsNull(rsaux2!FLOA_REE_CANTIDAD_LEIDA), 0, rsaux2!FLOA_REE_CANTIDAD_LEIDA), "###,###,##0.00")
                     list_item.SubItems(4) = Format(0, "###,###,##0.00")
                     list_item.SubItems(5) = Format(list_item.SubItems(2) - list_item.SubItems(3), "###,###,##0.00")
                     list_item.SubItems(6) = IIf(IsNull(rsaux2!FLOA_REE_COSTO_ENTRADA), "", rsaux2!FLOA_REE_COSTO_ENTRADA)
                     list_item.SubItems(7) = IIf(IsNull(rsaux!mone_Art_precio_base), "", rsaux!mone_Art_precio_base)
                     rsaux.Close
                     var_suma_cantidad_enviada = Format(var_suma_cantidad_enviada + rsaux2!FLOA_REE_CANTIDAD_ENTRADA, "###,###,##0.00")
                     var_suma_cantidad_recibida = Format(var_suma_cantidad_recibida + rsaux2!FLOA_REE_CANTIDAD_LEIDA, "###,###,##0.00")
                     rsaux2.MoveNext
               Wend
               lbl_enviados = Format(Str(var_suma_cantidad_enviada), "###,###,##0.00")
               lbl_recibidos = Format(Str(var_suma_cantidad_recibida), "###,###,##0.00")
               txt_clave_almacen_origen.Enabled = False
               txt_codigo.Enabled = True
               txt_codigo.SetFocus
            Else
               MsgBox "La entrada no existe", vbOKOnly, "ATENCION"
            End If
            rsaux2.Close
         Else
            MsgBox "Clave de almacen incorrecta", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
   End If
End Sub

Private Sub txt_codigo_GotFocus()
   txt_codigo = ""
End Sub

Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_codigo_seleccionado = ""
      frmbusqueda_articulo.Show 1
      Me.txt_codigo = var_codigo_seleccionado
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Dim var_recontable As Integer
   Dim var_cantidad_caja As Integer
   Dim var_caja As String
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   txt_codigo = Trim(txt_codigo)
   If KeyAscii = 13 Then
      var_verificador = True
      If Len(Trim(txt_codigo)) = 12 Then
         Call calcula_verificador(Trim(txt_codigo))
      End If
      If var_verificador = True Then
         var_caja = Left(txt_codigo, 6)
         'If var_caja = "000005" Or var_caja = "000010" Or var_caja = "000015" Or var_caja = "000020" Then
         If var_caja = "000005" Or var_caja = "000010" Or var_caja = "000015" Or var_caja = "000020" Or var_caja = "000001" Or var_caja = "000002" Or var_caja = "000003" Or var_caja = "000004" Or var_caja = "000006" Or var_caja = "000007" Or var_caja = "000008" Or var_caja = "000009" Or var_caja = "000011" Or var_caja = "0000012" Or var_caja = "0000013" Or var_caja = "0000014" Or var_caja = "000015" Or var_caja = "000016" Or var_caja = "000017" Or var_caja = "000018" Or var_caja = "000019" Or var_caja = "000021" Or var_caja = "000022" Or var_caja = "000023" Or var_caja = "000024" Or var_caja = "000025" Or var_caja = "000026" Or var_caja = "000027" Or var_caja = "000028" Or var_caja = "000029" Or var_caja = "000030" Then
            var_cantidad_caja = CInt(var_caja)
            txt_codigo = Mid(txt_codigo, 7, 5)
         End If
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
                     If var_cantidad_caja = 0 Then
                        If IsNull(rs(43).Value) Then
                           var_recontable = 0
                        Else
                           var_recontable = rs(43).Value
                        End If
                     Else
                        var_recontable = 0
                     End If
                     rs.Close
                     If var_recontable = 1 Then
                        var_cantidad_leida = 1#
                        lbl_cantidad.Visible = True
                        txt_cantidad.Visible = True
                        txt_cantidad.SetFocus
                     Else
                        If var_cantidad_caja = 0 Then
                           var_cantidad_leida = 1#
                        Else
                           var_cantidad_leida = var_cantidad_caja
                        End If
                        txt_foco.Enabled = True
                        txt_foco.SetFocus
                     End If
                  Else
                     txt_codigo = ""
                     frmmensaje.lbl_mensaje = "El artículo no existe"
                     frmmensaje.Show
                     'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                  End If
               Else
                  txt_codigo = ""
                  frmmensaje.lbl_mensaje = "El artículo no existe"
                  frmmensaje.Show
                  'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                  rs.Close
               End If
            End If
         Else
            txt_codigo = ""
            frmmensaje.lbl_mensaje = "Código Incorrecto"
            frmmensaje.Show
            'MsgBox "Código Incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         txt_codigo = ""
         frmmensaje.lbl_mensaje = "Error en Código"
         frmmensaje.Show
         'MsgBox "Error en Código", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   Dim var_encontro As Boolean
   Dim var_actualiza As Boolean
   Dim var_inserta As Boolean
   Dim bandera_suma As Boolean
   Dim var_cantidad As Variant
   Dim var_costo As Variant
   Dim var_precio As Variant
   Dim var_posible As Boolean
   Dim var_i As Integer
   Dim var_n As Integer
   Dim var_faltan As Double
   Dim var_codigo_salida As String
   Set TB_ARCH_COMPARACION_M = New TB_ARCH_COMPARACION_M
   Set TB_ARCH_COMPARACION_I = New TB_ARCH_COMPARACION_I
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   Dim var_folio_salida As String
   Dim var_cantidad_sobrante As Double
   If Trim(txt_codigo.Text) <> "" Then
      var_almacen_Destino = txt_clave_almacen_destino
      var_almacen_origen = txt_clave_almacen_origen
      bandera_suma = False
      If var_primera_vez = True Then
         var_inserta = False
         var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional_origen, txt_clave_almacen_origen, var_clave_movimiento, Now, var_numero_folio, 0, "", "", txt_clave_almacen_origen, txt_clave_almacen_destino, "", var_clave_usuario_global, fun_NombrePc, var_factura, "", txt_archivo, "", "B", "", "", 0, 0, 0, var_clave_moneda, var_tipo_Cambio)
         var_numero_folio = var_numero_folio_regreso
         txt_folio = var_numero_folio
         var_primera_vez = False
         var_fecha_movimiento = Date
      End If
      var_i = 1
      var_faltan = 0
      rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         var_cantidad_sobrante = var_cantidad_leida
         valor = txt_codigo
         var_n = lv_entradas.ListItems.Count
         var_encontro = 0
         var_i = 1
         var_cantidad_llegar = var_cantidad_leida
         var_cantidad_leida = 0
         var_suma_cantidad = 0
         While var_suma_cantidad < var_cantidad_llegar
               var_encontro = False
               var_cantidad_leida = var_cantidad_llegar - var_suma_cantidad
               var_i = 1
               While (var_i <= var_n)
                     lv_entradas.ListItems.Item(var_i).Selected = True
                     valor = Trim(lv_entradas.selectedItem.SubItems(1))
                     var_folio_salida = lv_entradas.selectedItem
                     If txt_codigo = valor Then
                        If var_folio_salida = "SOBRANTE" Then
                           var_encontro = True
                           var_i = var_n + 1
                           var_consecutivo = lv_entradas.selectedItem.SubItems(11)
                        Else
                           var_cantidad_posible = lv_entradas.selectedItem.SubItems(6) * 1
                           If var_cantidad_posible <= 0 Then
                              var_encontro = False
                           Else
                              var_consecutivo = lv_entradas.selectedItem.SubItems(11)
                              var_encontro = True
                              var_i = var_n + 1
                           End If
                        End If
                     End If
                     var_i = var_i + 1
               Wend
               If var_encontro = True Then
                  If var_folio_salida = "SOBRANTE" Then
                     var_cantidad_leida = var_cantidad_llegar - var_suma_cantidad
                  Else
                     If var_cantidad_posible >= var_cantidad_leida Then
                        var_cantidad_leida = var_cantidad_llegar
                        var_suma_cantidad = var_cantidad_llegar
                     Else
                        var_cantidad_leida = var_cantidad_posible
                        var_suma_cantidad = var_suma_cantidad + var_cantidad_leida
                     End If
                  End If
               End If
               If var_encontro = True Then
                  var_folio_salida = lv_entradas.selectedItem
                  var_codigo_salida = lv_entradas.selectedItem.SubItems(9)
                  var_faltan = lv_entradas.selectedItem.SubItems(6) * 1
                  If var_folio_salida = "SOBRANTE" Then
                     var_año = lv_entradas.selectedItem.SubItems(10) * 1
                     rsaux3.Open "update TB_REEMPAQUE_ENTRADA  set FLOA_REE_CANTIDAD_LEIDA = FLOA_REE_CANTIDAD_LEIDA + " + CStr(var_cantidad_leida) + " where VCHA_REE_FOLIO = '" + var_folio_salida + "' and VCHA_REE_ALMACEN_ORIGEN = '" + txt_clave_almacen_origen + "' and VCHA_REE_NUMERO = '" + txt_archivo + "' and VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' and vcha_ree_almacen_origen = '" + txt_clave_almacen_origen + "' and vcha_ree_articulo_salida = '" + var_codigo_salida + "' AND INTE_REE_AÑO = " + CStr(var_año), cnn, adOpenDynamic, adLockOptimistic
                     lv_entradas.selectedItem.SubItems(4) = Format(lv_entradas.selectedItem.SubItems(4) + var_cantidad_leida, "###,###,##0.00")
                     lv_entradas.selectedItem.SubItems(5) = Format(lv_entradas.selectedItem.SubItems(5) + var_cantidad_leida, "###,###,##0.00")
                     lv_entradas.selectedItem.SubItems(6) = Format(lv_entradas.selectedItem.SubItems(3) - lv_entradas.selectedItem.SubItems(4), "###,###,##0.00")
                     var_renglon = lv_entradas.selectedItem.Index
                     Call ilumina_grid
                     var_costo = lv_entradas.selectedItem.SubItems(7)
                     var_precio = lv_entradas.selectedItem.SubItems(8)
                     var_cantidad = lv_entradas.selectedItem.SubItems(5)
                     lbl_recibidos = Format(Int(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                     var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                     Cadena = "select * from TB_TEMPORAL_SALIDAS with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "' and vcha_ree_folio = '" + var_folio_salida + "' and vcha_sal_referencia = '" + var_codigo_salida + "' and inte_sal_año = " + CStr(var_año)
                     rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        If var_folio_salida = "SOBRANTE" Then
                           rsaux3.Open "SELECT * FROM TB_REEMPAQUE_SOBRANTES where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'  and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional_origen + "' and  VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and INTE_SAL_NUMERO =" + CStr(var_numero_folio) + " and VCHA_ART_ARTICULO_ID = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux3.EOF Then
                              rsaux2.Open "update TB_REEMPAQUE_SOBRANTES set floa_sal_cantidad = floa_sal_cantidad +" + CStr(var_cantidad_leida) + " where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'  and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional_origen + "' and  VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and INTE_SAL_NUMERO =" + CStr(var_numero_folio) + " and VCHA_ART_ARTICULO_ID = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                           Else
                              rsaux2.Open "select max(inte_ree_consecutivo) from tb_reempaque_sobrantes where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional_origen + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero  = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 var_consecutivo = IIf(IsNull(rsaux2(0).Value), 0, rsaux2(0).Value)
                              Else
                                 var_consecutivo = 0
                              End If
                              rsaux2.Close
                              var_consecutivo = var_consecutivo + 1
                              rsaux2.Open "insert into tb_reempaque_sobrantes ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD], [FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [inte_ree_año],[inte_ree_consecutivo]) values ('" + var_empresa + "', '" + var_unidad_organizacional_origen + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ", " + CStr(var_año) + "," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux3.Close
                        Else
                           rsaux2.Open "update tb_Temporal_salidas set floa_sal_cantidad = floa_sal_cantidad +" + CStr(var_cantidad_leida) + " where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'  and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional_origen + "' and  VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and INTE_SAL_NUMERO =" + CStr(var_numero_folio) + " and VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' and VCHA_REE_FOLIO = '" + var_folio_salida + "' and vcha_sal_referencia = '" + var_codigo_salida + "' and inte_sal_año = " + CStr(var_año), cnn, adOpenDynamic, adLockOptimistic
                        End If
                        var_suma_cantidad = var_suma_cantidad + var_cantidad_sobrante
                     Else
                        If var_folio_salida = "SOBRANTE" Then
                           rsaux3.Open "SELECT * FROM TB_REEMPAQUE_SOBRANTES where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'  and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional_origen + "' and  VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and INTE_SAL_NUMERO =" + CStr(var_numero_folio) + " and VCHA_ART_ARTICULO_ID = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux3.EOF Then
                              rsaux2.Open "update TB_REEMPAQUE_SOBRANTES set floa_sal_cantidad = floa_sal_cantidad +" + CStr(var_cantidad_leida) + " where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'  and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional_origen + "' and  VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and INTE_SAL_NUMERO =" + CStr(var_numero_folio) + " and VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' and inte_SAL_año = " + CStr(var_año), cnn, adOpenDynamic, adLockOptimistic
                           Else
                              rsaux2.Open "insert into tb_reempaque_sobrantes ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD], [FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [INTE_REE_AÑO]) values ('" + var_empresa + "', '" + var_unidad_organizacional_origen + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + "," + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux3.Close
                        Else
                           rsaux2.Open "insert into tb_Temporal_salidas ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD], [FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2], [VCHA_REE_FOLIO], [vcha_SAL_referencia], [INTE_SAL_AÑO]) values ('" + var_empresa + "', '" + var_unidad_organizacional_origen + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ",0, 0, 0, '" + var_folio_salida + "', '" + var_codigo_salida + "'," + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        var_suma_cantidad = var_suma_cantidad + var_cantidad_sobrante
                     End If
                     rs.Close
                     var_cantidad_sobrante = 0
                     var_cantidad_leida = 0
                  Else
                     var_año = lv_entradas.selectedItem.SubItems(10)
                     var_consecutivo = lv_entradas.selectedItem.SubItems(11)
                     rsaux3.Open "update TB_REEMPAQUE_ENTRADA  set FLOA_REE_CANTIDAD_LEIDA = FLOA_REE_CANTIDAD_LEIDA + " + CStr(var_cantidad_leida) + " where VCHA_REE_FOLIO = '" + var_folio_salida + "' and VCHA_REE_ALMACEN_ORIGEN = '" + txt_clave_almacen_origen + "' and VCHA_REE_NUMERO = '" + txt_archivo + "' and VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' and vcha_ree_almacen_origen = '" + txt_clave_almacen_origen + "' and vcha_ree_Articulo_salida = '" + var_codigo_salida + "' AND INTE_REE_AÑO = " + CStr(var_año) + " AND INTE_REE_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                     lv_entradas.selectedItem.SubItems(4) = Format(lv_entradas.selectedItem.SubItems(4) + var_cantidad_leida, "###,###,##0.00")
                     lv_entradas.selectedItem.SubItems(5) = Format(lv_entradas.selectedItem.SubItems(5) + var_cantidad_leida, "###,###,##0.00")
                     lv_entradas.selectedItem.SubItems(6) = Format(lv_entradas.selectedItem.SubItems(3) - lv_entradas.selectedItem.SubItems(4), "###,###,##0.00")
                     var_renglon = lv_entradas.selectedItem.Index
                     Call ilumina_grid
                     var_costo = lv_entradas.selectedItem.SubItems(7)
                     var_precio = lv_entradas.selectedItem.SubItems(8)
                     var_cantidad = lv_entradas.selectedItem.SubItems(5)
                     var_consecutivo = lv_entradas.selectedItem.SubItems(11)
                     var_año = lv_entradas.selectedItem.SubItems(10)
                     lbl_recibidos = Format(Int(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                     var_cantidad_recibida = var_cantidad_recibida + var_faltan
                     Cadena = "select * from TB_TEMPORAL_SALIDAS with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "' and vcha_ree_folio = '" + var_folio_salida + "' AND INTE_SAL_AÑO = " + CStr(var_año) + " AND INTE_SAL_CONSECUTIVO_reempaque = " + CStr(var_consecutivo)
                     rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        If var_folio_salida = "SOBRANTE" Then
                           rsaux3.Open "SELECT * FROM TB_REEMPAQUE_SOBRANTES where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'  and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional_origen + "' and  VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and INTE_SAL_NUMERO =" + CStr(var_numero_folio) + " and VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND INTE_REE_AÑO = " + CStr(var_año) + " and inte_ree_consecutivo =" + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux3.EOF Then
                              rsaux2.Open "update TB_REEMPAQUE_SOBRANTES set floa_sal_cantidad = floa_sal_cantidad +" + CStr(var_cantidad_leida) + " where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'  and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional_origen + "' and  VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and INTE_SAL_NUMERO =" + CStr(var_numero_folio) + " and VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND INTE_REE_AÑO = " + CStr(var_año) + " and inte_ree_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                           Else
                              rsaux2.Open "insert into tb_reempaque_sobrantes ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD], [FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [INTE_REE_AÑO],[inte_ree_consecutivo]) values ('" + var_empresa + "', '" + var_unidad_organizacional_origen + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + "," + CStr(var_año) + "," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux3.Close
                        Else
                           rsaux2.Open "update tb_Temporal_salidas set floa_sal_cantidad = floa_sal_cantidad +" + CStr(var_cantidad_leida) + " where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'  and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional_origen + "' and  VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and INTE_SAL_NUMERO =" + CStr(var_numero_folio) + " and VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' and VCHA_REE_FOLIO = '" + var_folio_salida + "' and vcha_sal_referencia = '" + var_codigo_salida + "' AND INTE_SAL_AÑO = " + CStr(var_año) + " AND inte_sal_consecutivo_reempaque = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                        End If
                     Else
                       
                        'MsgBox "insert into tb_Temporal_salidas ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD], [FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2], [VCHA_REE_FOLIO], [VCHA_SAL_REFERENCIA], INTE_SAL_AÑO, INTE_SAL_CONSECUTIVO_REEMPAQUE) values ('" + var_empresa + "', '" + var_unidad_organizacional_origen + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ",0, 0, 0, '" + var_folio_salida + "', '" + var_codigo_salida + "', " + var_año + "," + CStr(var_consecutivo) + ")"
                        'MsgBox var_precio
                        If var_precio = "" Then
                           var_precio = 0
                        End If
                        rsaux2.Open "insert into tb_Temporal_salidas ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD], [FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2], [VCHA_REE_FOLIO], [VCHA_SAL_REFERENCIA], INTE_SAL_AÑO, INTE_SAL_CONSECUTIVO_REEMPAQUE) values ('" + var_empresa + "', '" + var_unidad_organizacional_origen + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(IIf(IsNull(var_precio), 0, var_precio)) + ",0, 0, 0, '" + var_folio_salida + "', '" + var_codigo_salida + "', " + var_año + "," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rs.Close
                     var_cantidad_leida = var_cantidad_sobrante
                  End If
               End If
               If var_cantidad_leida > 0 And var_encontro = False Then
                  rsaux2.Open "select max(inte_ree_consecutivo) from tb_reempaque_Entrada where  VCHA_REE_ALMACEN_ORIGEN = '" + var_almacen_origen + "' and VCHA_ALM_ALMACEN_ID = '" + txt_clave_almacen_destino + "' and VCHA_REE_NUMERO = " + txt_archivo, cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     var_consecutivo = IIf(IsNull(rsaux2(0).Value), 0, rsaux2(0).Value)
                  Else
                     var_consecutivo = 0
                  End If
                  rsaux2.Close
                  var_consecutivo = var_consecutivo + 1
                  
                  lbl_recibidos = Format(Int(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                  var_precio = rsaux(2).Value
                  var_costo = rsaux!mone_Art_costo_estandar
                  var_año = 2005
                  Set list_item = lv_entradas.ListItems.Add(, , "SOBRANTE")
                  list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_Art_Articulo_id), "", rsaux!vcha_Art_Articulo_id)
                  list_item.SubItems(2) = IIf(IsNull(rsaux!vcha_Art_nombre_español), "", rsaux!vcha_Art_nombre_español)
                  list_item.SubItems(3) = Format(0, "###,###,##0.00")
                  list_item.SubItems(4) = Format(var_cantidad_leida, "###,###,##0.00")
                  list_item.SubItems(5) = Format(var_cantidad_leida, "###,###,##0.00")
                  list_item.SubItems(6) = Format(list_item.SubItems(3) - list_item.SubItems(4), "###,###,##0.00")
                  list_item.SubItems(10) = 2005
                  list_item.SubItems(11) = var_consecutivo
                  If var_entrada_calidad = True Then
                     rsaux2.Open "select * from tb_existencias where vcha_alm_almacen_id = '8' and vcha_art_articulo_id = '" + Trim(txt_codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_costo = 0
                     If Not rsaux2.EOF Then
                        var_costo = rsaux2!floa_exi_costo_2005
                     End If
                     rsaux2.Close
                     If var_costo = 0 Then
                        rsaux2.Open "select * from tb_articulos where vcha_art_articulo = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF() Then
                           var_costo = IIf(IsNull(rsuax2!mone_Art_costo_estandar), 0, rsaux2!mone_Art_costo_estandar)
                        End If
                        rsaux2.Close
                     End If
                  End If
                  list_item.SubItems(7) = IIf(IsNull(var_costo), 0, var_costo)
                  list_item.SubItems(8) = var_precio
                  list_item.SubItems(9) = txt_codigo
                  var_renglon = lv_entradas.ListItems.Count
                  Call ilumina_grid
                  'rsaux2.Open "insert into tb_Temporal_salidas ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD], [FLOA_SAL_COSTO], [FLOA_SAL_PRECIO], [FLOA_SAL_DESCUENTO], [FLOA_SAL_PROMOCION_1], [FLOA_SAL_PROMOCION_2], [VCHA_REE_FOLIO], [INTE_SAL_AÑO]) values ('" + var_empresa + "', '" + var_unidad_organizacional_origen + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_sobrante) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ",0, 0, 0, 'SOBRANTE',)", cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "insert into tb_reempaque_sobrantes ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_SAL_NUMERO], [VCHA_ART_ARTICULO_ID], [FLOA_SAL_CANTIDAD], [FLOA_SAL_COSTO], [FLOA_SAL_PRECIO],[INTE_SAL_AÑO]) values ('" + var_empresa + "', '" + var_unidad_organizacional_origen + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ", 2005)", cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "insert into TB_REEMPAQUE_ENTRADA ([VCHA_REE_FOLIO], [INTE_REE_TIPO_ENTRADA], [VCHA_REE_ALMACEN_ORIGEN], [VCHA_ALM_ALMACEN_ID], [VCHA_REE_NUMERO],[VCHA_REE_ARTICULO_SALIDA] ,[VCHA_ART_ARTICULO_ID], [FLOA_REE_COSTO_ENTRADA], [FLOA_REE_CANTIDAD_ENTRADA], [FLOA_REE_CANTIDAD_LEIDA], [INTE_REE_AÑO],[inte_ree_consecutivo])  values ('SOBRANTE'," + CStr(var_tipo_entrada) + " ,'" + var_almacen_origen + "', '" + txt_clave_almacen_destino + "', " + txt_archivo + ", '" + txt_codigo + "','" + txt_codigo + "', " + CStr(var_costo) + ",0, " + CStr(var_cantidad_leida) + ", 2005," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  var_suma_cantidad = var_cantidad_llegar
                  var_cantidad_leida = var_cantidad_llegar
               End If
         Wend
      End If
      rsaux.Close
      txt_codigo.SetFocus
   End If
End Sub

