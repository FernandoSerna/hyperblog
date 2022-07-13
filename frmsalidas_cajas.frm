VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmsalidas_cajas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_sellos 
      Height          =   2340
      Left            =   1065
      TabIndex        =   0
      Top             =   510
      Width           =   3045
      Begin VB.CommandButton cmd_cancelar_sello 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         Picture         =   "frmsalidas_cajas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   330
         Width           =   330
      End
      Begin VB.CommandButton cmd_aceptar_sello 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   45
         Picture         =   "frmsalidas_cajas.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   330
         Width           =   330
      End
      Begin VB.TextBox txt_sello 
         Height          =   315
         Left            =   585
         TabIndex        =   3
         Top             =   795
         Width           =   2385
      End
      Begin VB.CommandButton cmd_cerrar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   690
         Picture         =   "frmsalidas_cajas.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cerrar Alt + C"
         Top             =   330
         Width           =   330
      End
      Begin VB.Frame Frame4 
         Height          =   75
         Left            =   30
         TabIndex        =   1
         Top             =   645
         Width           =   2970
      End
      Begin MSComctlLib.ListView lv_sellos 
         Height          =   1200
         Left            =   30
         TabIndex        =   6
         Top             =   1110
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   2117
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Número de Sello"
            Object.Width           =   5115
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Sellos"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   7
         Left            =   30
         TabIndex        =   8
         Top             =   120
         Width           =   2970
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sello:"
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   840
         Width           =   390
      End
   End
   Begin VB.CommandButton cmd_cerrar_embarque 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmsalidas_cajas.frx":0396
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cerrar Embarque"
      Top             =   630
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      Picture         =   "frmsalidas_cajas.frx":0498
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cancelar Movimiento"
      Top             =   630
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   4185
      Left            =   90
      TabIndex        =   32
      Top             =   3150
      Width           =   7155
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
         Left            =   1560
         TabIndex        =   33
         Top             =   435
         Width           =   3390
      End
      Begin MSComctlLib.ListView lv_salidas 
         Height          =   3090
         Left            =   15
         TabIndex        =   34
         Top             =   1035
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   5450
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "          Código"
            Object.Width           =   9172
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cantidad"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "O.S."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Número Caja"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Factura ceros"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Tipo_pedido"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Estatus"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de Cajas"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   0
         Left            =   30
         TabIndex        =   36
         Top             =   120
         Width           =   7080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código de la Caja:"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   585
         Width           =   1290
      End
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   8190
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3630
      Width           =   165
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   45
      TabIndex        =   16
      Top             =   495
      Width           =   7200
   End
   Begin VB.TextBox txt_clave_movimiento 
      Height          =   285
      Left            =   2190
      TabIndex        =   14
      Top             =   660
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6870
      Picture         =   "frmsalidas_cajas.frx":059A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Salir"
      Top             =   630
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmsalidas_cajas.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cerrar Movimiento"
      Top             =   630
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmsalidas_cajas.frx":0CD6
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   630
      Width           =   330
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   1260
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
      Left            =   645
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   60
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_cajas.frx":0DD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_cajas.frx":16B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_cajas.frx":1F8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_cajas.frx":2528
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_cajas.frx":2E04
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_cajas.frx":36DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_cajas.frx":3FB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_cajas.frx":40CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_cajas.frx":41DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_cajas.frx":42EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_cajas.frx":4400
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_cajas.frx":4512
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_cajas.frx":4624
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_cajas.frx":47C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_cajas.frx":5618
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_cajas.frx":57EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_cajas.frx":5900
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   30
      TabIndex        =   15
      Top             =   870
      Width           =   7200
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Index           =   3
      Left            =   2865
      TabIndex        =   21
      Top             =   945
      Width           =   2220
      Begin VB.Label lbl_enviados 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   195
         TabIndex        =   23
         Top             =   420
         Width           =   1845
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad a Surtir"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   4
         Left            =   30
         TabIndex        =   22
         Top             =   120
         Width           =   2145
      End
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Index           =   4
      Left            =   5130
      TabIndex        =   24
      Top             =   945
      Width           =   2115
      Begin VB.Label lbl_recibidos 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   150
         TabIndex        =   26
         Top             =   420
         Width           =   1770
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad Surtida"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   5
         Left            =   30
         TabIndex        =   25
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Index           =   0
      Left            =   75
      TabIndex        =   17
      Top             =   945
      Width           =   2760
      Begin VB.TextBox txt_embarque 
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
         Left            =   915
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   390
         Width           =   1620
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Embarque"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   1
         Left            =   30
         TabIndex        =   20
         Top             =   120
         Width           =   2685
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   105
         TabIndex        =   19
         Top             =   540
         Width           =   765
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1350
      Index           =   1
      Left            =   90
      TabIndex        =   28
      Top             =   1800
      Width           =   7155
      Begin VB.TextBox txt_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   885
         TabIndex        =   38
         Top             =   825
         Width           =   6150
      End
      Begin VB.TextBox txt_origen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   885
         TabIndex        =   29
         Top             =   480
         Width           =   6150
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   39
         Top             =   855
         Width           =   555
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   31
         Top             =   510
         Width           =   660
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   30
         TabIndex        =   30
         Top             =   120
         Width           =   7080
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
      Left            =   135
      TabIndex        =   37
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmsalidas_cajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_estatus_embarque As String
Dim var_tipo_pedido As String
Dim var_orden_surtido As Double
Dim var_caja As Double
Dim var_numero_pedido_cliente As Double
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
Dim var_clave_agente As String
Dim var_clave_establecimiento As String
Dim var_clave_titular As String
Dim var_clave_cliente As String
Dim var_clave_ruta As String
Dim var_plazo As Integer
Dim var_descuento_1 As Double
Dim var_descuento_3 As Double
Dim var_descuento_2 As Double
Dim var_iva As Variant
Dim var_agrupador As String
Dim var_correo_electronico As String
Dim var_autorizo_embarque As Boolean
Dim var_es_caja As Boolean
Dim var_cajas As Boolean
Dim var_almacen_OS As String
Dim var_nota As recordSet
Dim var_movimiento_dependencia As String
Dim var_clave_moneda As String
Dim var_factura_ceros As Integer
Dim var_renglon As Double
Dim var_importe_total As Double
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Function fun_copia_archivo(Origen, Destino)
    Copy_File = CopyFile(Origen, Destino, 1)
End Function

Private Sub ilumina_grid()
    var_n = lv_salidas.ListItems.Count
    For var_i = 1 To var_n
        lv_salidas.ListItems.item(var_i).Selected = True
        If Trim(lv_salidas.selectedItem.SubItems(6)) = "S" Then
           lv_salidas.ListItems.item(var_i).Bold = False
           lv_salidas.ListItems.item(var_i).ListSubItems(1).Bold = False
           lv_salidas.ListItems.item(var_i).ListSubItems(2).Bold = False
           lv_salidas.ListItems.item(var_i).ListSubItems(3).Bold = False
           lv_salidas.ListItems.item(var_i).ListSubItems(4).Bold = False
           lv_salidas.ListItems.item(var_i).ListSubItems(5).Bold = False
           lv_salidas.ListItems.item(var_i).ListSubItems(6).Bold = False
           lv_salidas.ListItems.item(var_i).ForeColor = &HFF&
           lv_salidas.ListItems.item(var_i).ListSubItems(1).ForeColor = &HFF&
           lv_salidas.ListItems.item(var_i).ListSubItems(2).ForeColor = &HFF&
           lv_salidas.ListItems.item(var_i).ListSubItems(3).ForeColor = &HFF&
           lv_salidas.ListItems.item(var_i).ListSubItems(4).ForeColor = &HFF&
           lv_salidas.ListItems.item(var_i).ListSubItems(5).ForeColor = &HFF&
           lv_salidas.ListItems.item(var_i).ListSubItems(6).ForeColor = &HFF&
        Else
           lv_salidas.ListItems.item(var_i).Bold = False
           lv_salidas.ListItems.item(var_i).ListSubItems(1).Bold = False
           lv_salidas.ListItems.item(var_i).ListSubItems(2).Bold = False
           lv_salidas.ListItems.item(var_i).ListSubItems(3).Bold = False
           lv_salidas.ListItems.item(var_i).ListSubItems(4).Bold = False
           lv_salidas.ListItems.item(var_i).ListSubItems(5).Bold = False
           lv_salidas.ListItems.item(var_i).ListSubItems(6).Bold = False
           lv_salidas.ListItems.item(var_i).ForeColor = &H80000008
           lv_salidas.ListItems.item(var_i).ListSubItems(1).ForeColor = &H80000008
           lv_salidas.ListItems.item(var_i).ListSubItems(2).ForeColor = &H80000008
           lv_salidas.ListItems.item(var_i).ListSubItems(3).ForeColor = &H80000008
           lv_salidas.ListItems.item(var_i).ListSubItems(4).ForeColor = &H80000008
           lv_salidas.ListItems.item(var_i).ListSubItems(5).ForeColor = &H80000008
           lv_salidas.ListItems.item(var_i).ListSubItems(6).ForeColor = &H80000008
        End If
    Next var_i
    If var_renglon > 0 Then
       If var_renglon <= var_n Then
          var_i = var_renglon
          lv_salidas.ListItems.item(var_i).Bold = True
          lv_salidas.ListItems.item(var_i).ListSubItems(1).Bold = True
          lv_salidas.ListItems.item(var_i).ListSubItems(2).Bold = True
          lv_salidas.ListItems.item(var_i).ListSubItems(3).Bold = True
          lv_salidas.ListItems.item(var_i).ListSubItems(4).Bold = True
          lv_salidas.ListItems.item(var_i).ListSubItems(5).Bold = True
          lv_salidas.ListItems.item(var_i).ListSubItems(6).Bold = True
          lv_salidas.ListItems.item(var_i).ForeColor = &H8000&
          lv_salidas.ListItems.item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_salidas.ListItems.item(var_i).ListSubItems(2).ForeColor = &H8000&
          lv_salidas.ListItems.item(var_i).ListSubItems(3).ForeColor = &H8000&
          lv_salidas.ListItems.item(var_i).ListSubItems(4).ForeColor = &H8000&
          lv_salidas.ListItems.item(var_i).ListSubItems(5).ForeColor = &H8000&
          lv_salidas.ListItems.item(var_i).ListSubItems(6).ForeColor = &H8000&
       End If
    End If
    lv_salidas.Refresh
End Sub
Private Sub cmd_aceptar_sello_Click()
   If Trim(txt_sello) <> "" Then
      rs.Open "select * from tb_sellos where vcha_emp_empresa_id ='" + var_empresa + "' and vcha_sel_sello = '" + txt_sello + "'", cnn, adOpenDynamic, adLockOptimistic
      If rs.EOF Then
         rs.Close
         rs.Open "Insert Into tb_sellos (vcha_emp_empresa_id, inte_emb_embarque, vcha_sel_sello)  values ('" + var_empresa + "'," + Str(var_numero_embarque) + ",'" + txt_sello + "')", cnn, adOpenDynamic, adLockOptimistic
         Set list_item = lv_sellos.ListItems.Add(, , txt_sello)
      Else
         rs.Close
         MsgBox "El Sello ya existe", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de sello incorecto", vbOKOnly, "ATENCION"
   End If
   txt_sello = ""
   txt_sello.SetFocus
End Sub

Private Sub cmd_cancelar_sello_Click()
   frm_sellos.Visible = False
End Sub

Private Sub cmd_cerrar_Click()
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
   Set TB_ENTRADAS_INSERTA = New TB_ENTRADAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_LIBERA_APARTADOS = New TB_LIBERA_APARTADOS
   Set TB_SALIDA_VISTAS_I = New TB_SALIDA_VISTAS_I
   Set TB_ARCH_COMPARACION_I = New TB_ARCH_COMPARACION_I
   Set TB_ENC_EMBARQUE_M = New TB_ENC_EMBARQUE_M
   Dim var_referencia_vi As String
   Dim var_contador_renglones As Integer
   Dim var_cadena_cajas As String
   Dim var_posible As Boolean
   Dim var_copia As Boolean
   Dim var_eliminar As Boolean
   Dim var_nombre_archivo As String
   Dim var_numero_folio_anterior As Double
   Dim var_clave_moneda As String
   Dim var_moneda_local As Integer
   Dim var_tipo_Cambio As Double
   Dim var_posible_tipo_cambio As Boolean
   Dim var_clave_movimiento_anterior As String
   
   Dim var_catalogo_1 As String
   Dim var_catalogo_2 As String
   Dim var_fecha_surtido_catalogo As Date
   Dim var_importe_posible_surtido As Double
   Dim var_importe_surtir As Double
   Dim var_lista_precios_catalogo As String
   Dim var_precio_catalogo_1 As Double
   Dim var_precio_catalogo_2 As Double
   Dim var_importe_disponible As Double
   Dim var_importe_catalogos As Double
   Dim var_mes_catalogo As Integer
   Dim var_año_catalogo As Integer
   Dim var_numero_os As Double
   Dim var_numero_pedido As Double
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
   
   
   rs.Open "select * from tb_encabezado_embarques where inte_emb_embarque = " + Str(var_numero_embarque) + " AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_embarque_cerrado = IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", Trim(rs!CHAR_EMB_ESTATUS))
   End If
   rs.Close
   si = MsgBox("¿Esta seguro que desea cerrar el embarque?", vbYesNo, "ATENCION")
   If si = 6 Then
      si = MsgBox("Confirmar el cerrado del embarque", vbOKCancel, "ATENCION")
      If si = 1 Then
         var_cantidad_Salida = 0
         var_cantidad_cajas = 0
         var_cadena = "SELECT SUM(dbo.TB_TEMPORAL_SALIDAS.FLOA_SAL_CANTIDAD) AS CANTIDAD_SALIDA, dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID, dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE fROM dbo.TB_TEMPORAL_SALIDAS INNER JOIN dbo.TB_DETALLE_EMBARQUES ON dbo.TB_TEMPORAL_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID AND dbo.TB_TEMPORAL_SALIDAS.VCHA_UOR_UNIDAD_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID AND dbo.TB_TEMPORAL_SALIDAS.VCHA_ALM_ALMACEN_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID AND dbo.TB_TEMPORAL_SALIDAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_TEMPORAL_SALIDAS.INTE_SAL_NUMERO = dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO "
         var_cadena = var_cadena + " and dbo.TB_TEMPORAL_SALIDAS.vcha_uor_unidad_id = dbo.TB_DETALLE_EMBARQUES.vcha_uor_unidad_id GROUP BY dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID, dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE HAVING (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque + ") AND (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')"
         rsaux9.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux9.EOF Then
            var_cantidad_Salida = IIf(IsNull(rsaux9!cantidad_salida), 0, rsaux9!cantidad_salida)
         Else
            var_cantidad_Salida = 0
         End If
         rsaux9.Close
         rsaux9.Open "SELECT SUM(FLOA_PAQ_CANTIDAD) AS CANTIDAD_CAJAS, VCHA_EMP_EMPRESA_ID, INTE_EMB_EMBARQUE From dbo.TB_DETALLE_CAJAS WHERE     (CHAR_PAQ_ESTATUS = 'S') GROUP BY VCHA_EMP_EMPRESA_ID, INTE_EMB_EMBARQUE HAVING      (VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (INTE_EMB_EMBARQUE = " + Me.txt_embarque + ")", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux9.EOF Then
            var_cantidad_cajas = IIf(IsNull(rsaux9!cantidad_cajas), 0, rsaux9!cantidad_cajas)
         Else
            var_cantidad_cajas = 0
         End If
         rsaux9.Close
         var_clave_movimiento_anterior = var_clave_movimiento
         'var_cantidad_Salida = 0
         'var_cantidad_cajas = 0
         If Round(var_cantidad_Salida, 2) = Round(var_cantidad_cajas, 2) Then
            If Trim(var_embarque_cerrado) = "E" Then
               If rsaux3.State = 1 Then
                  rsaux3.Close
               End If
               
               var_posible_cerrar_KANBAN = True
               If var_posible_kanban = 1 Then
                  Set TB_PROC_KANBANS_EN_MOVIMIENTO = New TB_PROC_KANBANS_EN_MOVIMIENTO
                  rsaux3.Open "SELECT distinct VCHA_ALM_ALMACEN_ID, 'CAJA-'+VCHA_EMP_EMPRESA_ID+'-'+CAST(INTE_PAQ_CAJA AS VARCHAR(50)) AS MOVIMIENTO FROM TB_dETALLE_cAJAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND FLOA_PAQ_CANTIDAD > 0 and CHAR_PAQ_ESTATUS <> 'C'"
                  'rsaux3.Open "distinct 'CAJA-'+VCHA_EMP_EMPRESA_ID+CAST(INTE_PAQ_NUMERO AS VARCHAR(50)) FROM TB_dETALLE_cAJAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND FLOA_PAQ_CANTIDAD > 0", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux3.EOF
                        var_inserta = TB_PROC_KANBANS_EN_MOVIMIENTO.Anadir(rsaux3!VCHA_ALM_ALMACEN_ID, rsaux3!MOVIMIENTO, CDbl(Me.txt_embarque), "", "")
                        If var_kanban_exito = "N" Then
                           var_posible_cerrar_KANBAN = False
                        End If
                        rsaux3.MoveNext
                  Wend
                  rsaux3.Close
               Else
                  var_posible_cerrar_KANBAN = True
               End If
               
               
               If rsaux3.State = 1 Then
                  rsaux3.Close
               End If
               
               rsaux3.Open "select distinct * from vw_embarques_cerrar where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_emb_embarque = " + txt_embarque + " and char_emb_estatus = 'E'", cnn, adOpenDynamic, adLockOptimistic
               var_tipo_Cambio = 0
               var_posible_tipo_cambio = True
               While Not rsaux3.EOF
                  var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
                  If var_moneda_local = 0 Then
                     var_tipo_Cambio = IIf(IsNull(rsaux3!mone_tca_importe), 0, rsaux3!mone_tca_importe)
                     If var_tipo_Cambio = 0 Then
                        var_posible_tipo_cambio = False
                     End If
                  End If
                  rsaux3.MoveNext
               Wend
               'MsgBox var_unidad_organizacional
               If var_posible_tipo_cambio = True Then
                  var_numero_folio_anterior = var_numero_folio
                  rsaux3.MoveFirst
                  While Not rsaux3.EOF
                        var_clave_movimiento = rsaux3!VCHA_MOV_MOVIMIENTO_ID
                        var_numero_folio = rsaux3!INTE_SAL_NUMERO
                        var_clave_moneda = rsaux3!vcha_mon_moneda_id
                        var_almacen_origen = rsaux3!VCHA_ALM_ALMACEN_ID
                        var_almacen_OS = var_almacen_origen
                        var_estatus_movimiento = rsaux3!char_Emo_estatus
                        var_numero_os = rsaux3!inte_emo_numero_origen
                        var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
                        var_clave_cliente = IIf(IsNull(rsaux3!vcha_cli_clave_id), "", rsaux3!vcha_cli_clave_id)
                        If var_moneda_local = 0 Then
                           var_tipo_Cambio = IIf(IsNull(rsaux3!mone_tca_importe), 0, rsaux3!mone_tca_importe)
                        Else
                           var_tipo_Cambio = 1
                        End If
                        If var_numero_folio > 0 Then
                           If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
                           Else
                              If var_tipo_Cambio > 0 Then
                                 If var_fecha_surtido_catalogo <= Date Then
                                    var_si_surtir_catalogo = 1
                                 Else
                                    var_si_surtir_catalogo = 0
                                 End If
                                 If var_clave_movimiento = "FT" Then
                                    rsaux4.Open "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES_TIENDAS '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", " + CStr(var_tipo_Cambio) + ",'" + var_catalogo_1 + "','" + var_catalogo_2 + "','" + var_clave_titular + "','" + var_clave_cliente + "'," + CStr(var_año_catalogo) + "," + CStr(var_mes_catalogo) + "," + CStr(var_si_surtir_catalogo), cnn, adOpenDynamic, adLockOptimistic
                                 Else
                                    If parametros(0) = "SQLHOUSTON" Then
                                       rsaux4.Open "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", " + CStr(var_tipo_Cambio) + ",'" + var_catalogo_1 + "','" + var_catalogo_2 + "','" + var_clave_titular + "','" + var_clave_cliente + "'," + CStr(var_año_catalogo) + "," + CStr(var_mes_catalogo) + "," + CStr(var_si_surtir_catalogo), cnn, adOpenDynamic, adLockOptimistic
                                    Else
                                       rsaux4.Open "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", " + CStr(var_tipo_Cambio) + ",'" + var_catalogo_1 + "','" + var_catalogo_2 + "','" + var_clave_titular + "','" + var_clave_cliente + "'," + CStr(var_año_catalogo) + "," + CStr(var_mes_catalogo) + "," + CStr(var_si_surtir_catalogo), cnn, adOpenDynamic, adLockOptimistic
                                    End If
                                 End If
                                 
                                 If var_tipo_lectura = 1 Then
                                    rsaux4.Open "select inte_ped_numero from tb_enc_orden_surtido where inte_ors_orden_surtido = " + CStr(var_numero_os), cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux4.EOF Then
                                       var_numero_pedido = rsaux4!inte_ped_numero
                                    Else
                                       var_numero_pedido = 0
                                    End If
                                    rsaux4.Close
                                    rsaux4.Open "SELECT * FROM TB_DETALLE_CAJAS with (nolock) WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_EMB_EMBARQUE = " + txt_embarque + " AND CHAR_PAQ_ESTATUS <> 'C'", cnn, adOpenDynamic, adLockOptimistic
                                    'rsaux4.Open "SELECT * FROM TB_TEMPORAL_SALIDAS  with (nolock) WHERE VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID  = '" + var_almacen_origen + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_SAL_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                                    While Not rsaux4.EOF
                                          'rsaux5.Open "UPDATE TB_DETALLE_pedidos set floa_ped_cantidad_surtida = floa_ped_cantidad_surtida + " + CStr(rsaux4!floa_sal_cantidad) + " where inte_ped_numero = " + CStr(var_numero_pedido) + " and vcha_art_articulo_id = '" + rsaux4!vcha_art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                          rsaux6.Open "SELECT * FROM TB_ENC_ORDEN_SURTIDO WHERE INTE_ORS_ORDEN_SURTIDO = " + CStr(rsaux4!INTE_ORS_ORDEN_SURTIDO), cnn, adOpenDynamic, adLockOptimistic
                                          var_numero_pedido = IIf(IsNull(rsaux6!inte_ped_numero), 0, rsaux6!inte_ped_numero)
                                          rsaux6.Close
                                          rsaux5.Open "UPDATE TB_DETALLE_pedidos set floa_ped_cantidad_surtida = floa_ped_cantidad_surtida + " + CStr(rsaux4!floa_paq_cantidad) + " where inte_ped_numero = " + CStr(var_numero_pedido) + " and vcha_art_articulo_id = '" + rsaux4!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                          rsaux4.MoveNext
                                    Wend
                                    rsaux4.Close
                                    '' lo inivire para que no lo ejecute
                                    x = 1
                                    If x = 0 Then
                                    If var_posible_paqueteria = 1 Then
                                       Dim oleppembaruqe  As Paqueteria.Embarque
                                       Set oleppembaruqe = CreateObject("paqueteria.embarque")
                                       Dim oleAppcaja As Paqueteria.Caja
                                       Set oleAppcaja = CreateObject("paqueteria.caja")
                                       Dim oleppfactura As Paqueteria.FacturaCXC
                                       Set oleppfactura = CreateObject("paqueteria.facturacxc")
                                                                  
                                       var_cadena = "SELECT     SUM((dbo.TB_SALIDAS.FLOA_SAL_PRECIO * (1 - dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1 / 100)) * (1 - dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2 / 100) * 1 + dbo.VW_CLIENTES.FLOA_TPE_IVA / 100) AS IMPORTE, dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_EMBARQUES.VCHA_EMP_EMPRESA_ID, dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.VW_CLIENTES.VCHA_AGE_AGENTE_ID, dbo.VW_CLIENTES.VCHA_CLI_REFERENCIA, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID FROM dbo.TB_ENCABEZADO_EMBARQUES INNER JOIN dbo.TB_DETALLE_EMBARQUES ON dbo.TB_ENCABEZADO_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE = dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE INNER JOIN dbo.TB_SALIDAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND"
                                       var_cadena = var_cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_SALIDAS.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN dbo.VW_CLIENTES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_CLI_CLAVE_ID = dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID "
                                       var_cadena = var_cadena + " GROUP BY dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_EMBARQUES.VCHA_EMP_EMPRESA_ID, dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.VW_CLIENTES.VCHA_AGE_AGENTE_ID, dbo.VW_CLIENTES.VCHA_CLI_REFERENCIA, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID HAVING (dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque + ") AND (dbo.TB_ENCABEZADO_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')"
                                       rsaux9.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux9.EOF Then
                                          var_importe_total = IIf(IsNull(rsaux9!Importe), 0, rsaux9!Importe)
                                          oleppembaruqe.Identificador = 0
                                          oleppembaruqe.fecha = Date
                                          oleppembaruqe.Agente = rsaux9!VCHA_AGE_AGENTE_ID
                                          oleppembaruqe.Cliente = rsaux9!vcha_cli_clave_id
                                          oleppembaruqe.Referencia = rsaux9!VCHA_CLI_REFERENCIA
                                          rsaux10.Open "select distinct inte_paq_caja, floa_pca_precio, floa_paq_seguro, floa_paq_seguro_Costo from tb_detalle_Cajas where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque + " and floa_paq_cantidad > 0", cnn, adOpenDynamic, adLockOptimistic
                                          var_Importe_flete = 0
                                          While Not rsaux10.EOF
                                                var_Importe_flete = var_Importe_flete + IIf(IsNull(rsaux10!floa_pca_precio), 0, rsaux10!floa_pca_precio)
                                                rsaux10.MoveNext
                                          Wend
                                          rsaux10.MoveFirst
                                          var_seguro_1000_cliente = IIf(IsNull(rsaux10!floa_paq_seguro), 0, rsaux10!floa_paq_seguro)
                                          var_seguro_1000 = IIf(IsNull(rsaux10!FLOA_PAQ_SEGURO_COSTO), 0, rsaux10!FLOA_PAQ_SEGURO_COSTO)
                                          rsaux10.Close
                                          
                                          rsaux10.Open "select DBO.CALCULO_SEGURO (" + CStr(IIf(IsNull(rsaux9!Importe), 0, rsaux9!Importe)) + "," + CStr(var_seguro_1000_cliente) + ") from tb_lineas", cnn, adOpenDynamic, adLockOptimistic
                                          If rsaux10.EOF Then
                                             var_importe_seguro_cliente = 0
                                          Else
                                             var_importe_seguro_cliente = rsaux10(0).Value
                                          End If
                                          rsaux10.Close
                                          
                                          rsaux10.Open "select DBO.CALCULO_SEGURO (" + CStr(IIf(IsNull(rsaux9!Importe), 0, rsaux9!Importe)) + "," + CStr(var_seguro_1000) + ") from tb_lineas", cnn, adOpenDynamic, adLockOptimistic
                                          If rsaux10.EOF Then
                                             var_importe_seguro = 0
                                          Else
                                             var_importe_seguro = rsaux10(0).Value
                                          End If
                                          rsaux10.Close
                                        
                                          rsaux10.Open "select top 1 isnull(vcha_paq_clave_id,'') as vcha_paq_clave_id from tb_Detalle_cajas where vcha_Emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque + " and vcha_paq_clave_id is not null", cnn, adOpenDynamic, adLockOptimistic
                                          VAR_CLAVE_PAQUETERIA = ""
                                          If Not rsaux10.EOF Then
                                             VAR_CLAVE_PAQUETERIA = rsaux10!vcha_paq_clave_id
                                          End If
                                          rsaux10.Close
                                          oleppembaruqe.CostoSeguro = var_importe_seguro 'lo que cuesta el seguro
                                          oleppembaruqe.Seguro = var_importe_seguro_cliente 'lo que paga el cliente
                                          oleppembaruqe.FleteCliente = var_Importe_flete
                                          oleppembaruqe.ServicioPaqueteria = VAR_CLAVE_PAQUETERIA
                                          var_cadena = "SELECT SUM(FLOA_PAQ_CANTIDAD) AS cantidad, VCHA_EMP_EMPRESA_ID, INTE_EMB_EMBARQUE, VCHA_PAQ_CLAVE_ID, VCHA_CAJ_CAJA_ID, VCHA_PAQ_GUIA , FLOA_PCA_PRECIO, FLOA_PCA_COSTO, CHAR_PAQ_ESTATUS From dbo.TB_DETALLE_CAJAS GROUP BY VCHA_EMP_EMPRESA_ID, INTE_EMB_EMBARQUE, VCHA_PAQ_CLAVE_ID, VCHA_CAJ_CAJA_ID, VCHA_PAQ_GUIA, FLOA_PCA_PRECIO, FLOA_PCA_COSTO, CHAR_PAQ_ESTATUS HAVING      (VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (INTE_EMB_EMBARQUE = " + Me.txt_embarque + ") AND (CHAR_PAQ_ESTATUS <> 'C') and VCHA_PAQ_GUIA is not null"
                                          rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                          While Not rsaux10.EOF
                                                oleAppcaja.Cantidad = IIf(IsNull(rsaux10!Cantidad), 0, rsaux10!Cantidad)
                                                oleAppcaja.Precio = IIf(IsNull(rsaux10!floa_pca_precio), 0, rsaux10!floa_pca_precio)
                                                oleAppcaja.Costo = IIf(IsNull(rsaux10!floa_pca_costo), 0, rsaux10!floa_pca_costo)
                                                oleAppcaja.Guia = IIf(IsNull(rsaux10!vcha_paq_guia), "", rsaux10!vcha_paq_guia)
                                                oleAppcaja.TipoCaja = IIf(IsNull(rsaux10!vcha_caj_caja_id), "", rsaux10!vcha_caj_caja_id)
                                                oleppembaruqe.Cajas.Add oleAppcaja
                                                rsaux10.MoveNext
                                          Wend
                                          rsaux10.Close
                                          oleppfactura.Identificador = Me.txt_embarque
                                          oleppfactura.Serie = var_empresa
                                          oleppfactura.Importe = var_importe_total
                                          oleppembaruqe.facturas.Add oleppfactura
                                          
                                          oleppembaruqe.Registrar
                                       
                                       
                                       End If
                                       rsaux9.Close
                                    End If
                                    End If '' se termina la inivicion de la paqueteria
                                 End If
                              End If
                           End If
                        End If
                        rsaux3.MoveNext
                  Wend
                  rsaux3.Close
                  ok = False
                  ok = TB_ENC_EMBARQUE_M.Anadir(var_empresa, var_unidad_organizacional, var_numero_embarque, "I")
                  var_si = MsgBox("¿Desea cerrar los pedidos del embarque?", vbYesNo, "ATENCION")
                  var_si = 6
                  If var_si = 6 Then
   
                     rsaux4.Open "SELECT     VCHA_EMP_EMPRESA_ID, INTE_EMB_EMBARQUE, INTE_ORS_ORDEN_SURTIDO, VCHA_ART_ARTICULO_ID, SUM(FLOA_PAQ_CANTIDAD) AS CANTIDAD From dbo.TB_DETALLE_CAJAS WHERE     (VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (INTE_EMB_EMBARQUE = " + Me.txt_embarque + ") AND (CHAR_PAQ_ESTATUS = 'S') GROUP BY VCHA_EMP_EMPRESA_ID, INTE_EMB_EMBARQUE, VCHA_ART_ARTICULO_ID, INTE_ORS_ORDEN_SURTIDO", cnn, adOpenDynamic, adLockOptimistic
                     While Not rsaux4.EOF
                           rsaux.Open "UPDATE TB_dET_ORDEN_SURTIDO SET FLOA_ORS_CANTIDAD_SALIDA = FLOA_ORS_CANTIDAD_SALIDA + " + CStr(rsaux4!Cantidad) + " WHERE INTE_ORS_ORDEN_SURTIDO = " + CStr(rsaux4!INTE_ORS_ORDEN_SURTIDO) + " AND VCHA_ART_ARTICULO_ID = '" + CStr(rsaux4!VCHA_ART_ARTICULO_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
                           rsaux4.MoveNext
                     Wend
                     rsaux4.Close

                     'rsaux4.Open "SELECT * FROM VW_EMBARQUES_PEDIDOS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                     rsaux4.Open "SELECT DISTINCT dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO, dbo.TB_DETALLE_CAJAS.VCHA_EMP_EMPRESA_ID, dbo.TB_DETALLE_CAJAS.INTE_EMB_EMBARQUE FROM dbo.TB_DETALLE_CAJAS INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO ON dbo.TB_DETALLE_CAJAS.INTE_ORS_ORDEN_SURTIDO = dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO WHERE (dbo.TB_DETALLE_CAJAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_DETALLE_CAJAS.INTE_EMB_EMBARQUE = " + Me.txt_embarque + ")"
                     While Not rsaux4.EOF
                           rsaux.Open "update tb_encabezado_pedidos set CHAR_PED_ESTATUS = 'E' where inte_ped_numero = " + CStr(rsaux4!inte_ped_numero), cnn, adOpenDynamic, adLockOptimistic
                           rsaux4.MoveNext
                     Wend
                     rsaux4.Close
                  
                  
                  End If
                  var_estatus_movimiento = "I"
                  var_numero_folio = var_numero_folio_anterior
                  var_embarque_cerrado = "I"
                  MsgBox "Se a cerrado el embarque", vbOKOnly, "ATENCION"
                  If var_clave_movimiento = "FA" Or var_clave_movimiento = "EX" Then
                     x = Shell("net send temporal Imprimir facturas de embarque " + Trim(txt_embarque), vbHide)
                  End If
                  rsaux5.Open "update tb_encabezado_embarques set dtim_emb_fecha_final = getdate() where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
               Else
                  rsaux3.Close
                  MsgBox "No es posible cerrar el embarque ya que no se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El embaruqe ya habia sido cerrado con anterioridad", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se puede cerrar el embarque ya que la cantidad de salida es diferente a la cantidad en cajas", vbOKOnly, "ATENCION"
            Dim var_actualiza As Boolean
            Dim bandera_suma As Boolean
            Dim var_cantidad As Variant
            Dim var_costo As Variant
            Dim var_precio As Variant
            Dim var_posible_caja As Boolean
            Dim var_cantidad_posible As Variant
            Dim var_embarque_paquete As Integer
            Dim var_embarque_caja As Integer
            Dim var_estatus_caja As String
            Dim var_orden_surtido_caja As Double
            Dim var_posible_empaque As Boolean
            Dim var_promocion_1 As Double
            Dim var_promocion_2 As Double
            Dim var_encontrado As Integer
            Dim var_canal_venta As String
            Dim var_i As Integer
            Dim var_n As Integer
            Dim var_j As Integer
            Dim var_jj As Integer
            Dim var_orden_surtido_03 As Double
            Dim var_agente_coppel As String
            Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
            Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
            Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
            Set TB_DET_ORDEN_SURTIDO_M = New TB_DET_ORDEN_SURTIDO_M
            Set TB_DET_EMBARQUE_I = New TB_DET_EMBARQUE_I
            Set TB_DETALLE_CAJAS_M = New TB_DETALLE_CAJAS_M
            
            
            
            
            rsaux9.Open "select * from tb_detalle_embarques where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux9.EOF
                  rsaux10.Open "delete from tb_temporal_salidas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + rsaux9!VCHA_UOR_UNIDAD_ID + "' and vcha_mov_movimiento_id = '" + rsaux9!VCHA_MOV_MOVIMIENTO_ID + "' and inte_sal_numero = " + CStr(rsaux9!INTE_SAL_NUMERO), cnn, adOpenDynamic, adLockOptimistic
                  rsaux10.Open "delete from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + rsaux9!VCHA_UOR_UNIDAD_ID + "' and vcha_mov_movimiento_id = '" + rsaux9!VCHA_MOV_MOVIMIENTO_ID + "' and inte_emo_numero = " + CStr(rsaux9!INTE_SAL_NUMERO), cnn, adOpenDynamic, adLockOptimistic
                  rsaux9.MoveNext
            Wend
            rsaux9.Close
            rsaux10.Open "delete from tb_detalle_embarques where vcha_Emp_empresa_id = '" + var_empresa + "' and inte_Emb_embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
            rsaux10.Open "update tb_detalle_cajas set char_paq_estatus = 'I' where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque + " and char_paq_estatus = 'S'", cnn, adOpenDynamic, adLockOptimistic
            lbl_recibidos = Format(0, "###,###,##0.00")
            var_primera_vez = True
            For var_jj = 1 To Me.lv_salidas.ListItems.Count
                Me.lv_salidas.ListItems.item(var_jj).Selected = True
                Me.txt_codigo = Me.lv_salidas.selectedItem
                var_orden_surtido = lv_salidas.selectedItem.SubItems(2)
                var_caja = lv_salidas.selectedItem.SubItems(3)
                var_factura_ceros = lv_salidas.selectedItem.SubItems(4)
                var_tipo_pedido = lv_salidas.selectedItem.SubItems(5)

                z = 0
                cnn.CommandTimeout = 360
                If Trim(txt_codigo.Text) <> "" Then
                   rs.Open "select vcha_age_Agente_id from tb_encabezado_embarques where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque
                   If Not rs.EOF Then
                      var_agente_coppel = rs!VCHA_AGE_AGENTE_ID
                   End If
                   rs.Close
                   If var_clave_titular <> "T000001474" Then
                      var_agente_coppel = ""
                   End If
                   If var_primera_vez = True Then
                      var_inserta = False
                      var_agente_coppel = ""
                      If var_empresa = "03" Or var_agente_coppel = "00143" Then
                         rs.Open "select vcha_emp_empresa_id, min(inte_ors_orden_surtido) as inte_ors_orden_surtido from tb_detalle_cajas where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque + " group by vcha_emp_empresa_id", cnn, adOpenDynamic, adLockOptimistic
                         While Not rs.EOF
                               rsaux4.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + CStr(rs!INTE_ORS_ORDEN_SURTIDO), cnn, adOpenDynamic, adLockOptimistic
                               var_orden_surtido = 0
                               var_clave_cliente = ""
                               var_almacen_Destino = ""
                               var_clave_establecimiento = ""
                               var_clave_titular = ""
                               var_descuento_1 = 0
                               var_descuento_2 = 0
                               var_clave_moneda = 0
                
                               var_orden_surtido = rs!INTE_ORS_ORDEN_SURTIDO
                               var_clave_cliente = rsaux4!vcha_cli_clave_id
                               var_almacen_Destino = ""
                               var_clave_establecimiento = rsaux4!vcha_ESB_ESTABLECIMIENTO_id
                               var_clave_titular = rsaux4!vcha_tit_titular_id
                               var_descuento_1 = IIf(IsNull(rsaux4!FLOA_ORS_DESCUENTO_1), 0, rsaux4!FLOA_ORS_DESCUENTO_1)
                               var_descuento_2 = IIf(IsNull(rsaux4!FLOA_ORS_DESCUENTO_2), 0, rsaux4!FLOA_ORS_DESCUENTO_2)
                               var_clave_moneda = IIf(IsNull(rsaux4!vcha_mon_moneda_id), "", rsaux4!vcha_mon_moneda_id)
                               rsaux4.Close
                               rsaux.Open "select vcha_can_canal_venta_id from tb_agentes where vcha_age_agente_id = '" + var_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                               var_canal_venta = IIf(IsNull(rsaux!vcha_can_canal_venta_id), "", rsaux!vcha_can_canal_venta_id)
                               rsaux.Close
                               If var_empresa = "03" Or var_agente_coppel = "00143" Then
                                  'rsaux4.Open "select * from tb_encabezado_movimientos where inte_emo_numero_origen = " + CStr(var_orden_surtido) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                                  'If rsaux4.EOF Then
                                  var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, var_orden_surtido, var_clave_cliente, "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", CStr(rs!INTE_ORS_ORDEN_SURTIDO), var_clave_establecimiento, "B", var_clave_titular, var_clave_agente, var_descuento_1, var_descuento_2, var_descuento_3, var_clave_moneda, 0)
                                  var_numero_folio = var_numero_folio_regreso
                                  If var_factura_ceros = 1 Then
                                     rsaux.Open "update tb_encabezado_movimientos set inte_emo_factura_ceros = 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                                  End If
                                  var_pedido_credito = 1
                      
                                  If var_clave_movimiento = "FT" Then
                                     rsaux.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + CStr(var_orden_surtido), cnn, adOpenDynamic, adLockOptimistic
                                     var_pedido_credito = 1
                                     If Not rsaux.EOF Then
                                        var_pedido_credito = IIf(IsNull(rsaux!inte_ors_pedido_credito), 1, rsaux!inte_ors_pedido_credito)
                                     End If
                                     rsaux.Close
                                  End If
                       
                                  rsaux.Open "update tb_encabezado_movimientos set VCHA_CAN_CANAL_VENTA_ID = '" + var_canal_venta + "', inte_emo_pedido_credito = " + CStr(var_pedido_credito) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                                  var_inserta = False
                                  var_inserta = TB_DET_EMBARQUE_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_numero_embarque, var_clave_movimiento, var_numero_folio, "")
                                  txt_folio = var_numero_folio
                                  var_primera_vez = False
                                  'End If
                                  'rsaux4.Close
                               Else
                                  rsaux4.Open "select * from tb_encabezado_movimientos where inte_emo_numero_origen = " + CStr(var_orden_surtido) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                                  If rsaux4.EOF Then
                                     var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, var_orden_surtido, var_clave_cliente, "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", CStr(rs!INTE_ORS_ORDEN_SURTIDO), var_clave_establecimiento, "B", var_clave_titular, var_clave_agente, var_descuento_1, var_descuento_2, var_descuento_3, var_clave_moneda, 0)
                                     var_numero_folio = var_numero_folio_regreso
                                     If var_factura_ceros = 1 Then
                                        rsaux.Open "update tb_encabezado_movimientos set inte_emo_factura_ceros = 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                                     End If
                                     var_pedido_credito = 1
                                     If var_clave_movimiento = "FT" Then
                                        rsaux.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + CStr(var_orden_surtido), cnn, adOpenDynamic, adLockOptimistic
                                        var_pedido_credito = 1
                                        If Not rsaux.EOF Then
                                           var_pedido_credito = IIf(IsNull(rsaux!inte_ors_pedido_credito), 1, rsaux!inte_ors_pedido_credito)
                                        End If
                                        rsaux.Close
                                     End If
                        
                                     rsaux.Open "update tb_encabezado_movimientos set VCHA_CAN_CANAL_VENTA_ID = '" + var_canal_venta + "', inte_emo_pedido_credito = " + CStr(var_pedido_credito) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                                     var_inserta = False
                                     var_inserta = TB_DET_EMBARQUE_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_numero_embarque, var_clave_movimiento, var_numero_folio, "")
                                     txt_folio = var_numero_folio
                                     var_primera_vez = False
                                  End If
                                  rsaux4.Close
                               End If
                               rs.MoveNext
                         Wend
                         rs.Close
                      Else
                         rs.Open "select distinct vcha_emp_empresa_id, inte_ors_orden_surtido from tb_detalle_cajas where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
                         While Not rs.EOF
                               rsaux4.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + CStr(rs!INTE_ORS_ORDEN_SURTIDO), cnn, adOpenDynamic, adLockOptimistic
                               var_orden_surtido = 0
                               var_clave_cliente = ""
                               var_almacen_Destino = ""
                               var_clave_establecimiento = ""
                               var_clave_titular = ""
                               var_descuento_1 = 0
                               var_descuento_2 = 0
                               var_clave_moneda = 0
                               var_orden_surtido = rs!INTE_ORS_ORDEN_SURTIDO
                               var_clave_cliente = rsaux4!vcha_cli_clave_id
                               var_almacen_Destino = ""
                               var_clave_establecimiento = rsaux4!vcha_ESB_ESTABLECIMIENTO_id
                               var_clave_titular = rsaux4!vcha_tit_titular_id
                               var_descuento_1 = IIf(IsNull(rsaux4!FLOA_ORS_DESCUENTO_1), 0, rsaux4!FLOA_ORS_DESCUENTO_1)
                               var_descuento_2 = IIf(IsNull(rsaux4!FLOA_ORS_DESCUENTO_2), 0, rsaux4!FLOA_ORS_DESCUENTO_2)
                               var_clave_moneda = IIf(IsNull(rsaux4!vcha_mon_moneda_id), "", rsaux4!vcha_mon_moneda_id)
                               rsaux4.Close
                               rsaux.Open "select vcha_can_canal_venta_id from tb_agentes where vcha_age_agente_id = '" + var_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                               var_canal_venta = IIf(IsNull(rsaux!vcha_can_canal_venta_id), "", rsaux!vcha_can_canal_venta_id)
                               rsaux.Close
                               If var_clave_titular <> "T000001474" Then
                                  var_agente_coppel = ""
                               End If
                               If var_empresa = "03" Or var_agente_coppel = "00143" Then
                                  'rsaux4.Open "select * from tb_encabezado_movimientos where inte_emo_numero_origen = " + CStr(var_orden_surtido) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                                  'If rsaux4.EOF Then
                                  var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, var_orden_surtido, var_clave_cliente, "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", CStr(rs!INTE_ORS_ORDEN_SURTIDO), var_clave_establecimiento, "B", var_clave_titular, var_clave_agente, var_descuento_1, var_descuento_2, var_descuento_3, var_clave_moneda, 0)
                                  var_numero_folio = var_numero_folio_regreso
                                  If var_factura_ceros = 1 Then
                                     rsaux.Open "update tb_encabezado_movimientos set inte_emo_factura_ceros = 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                                  End If
                                  var_pedido_credito = 1
                                  If var_clave_movimiento = "FT" Then
                                     rsaux.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + CStr(var_orden_surtido), cnn, adOpenDynamic, adLockOptimistic
                                     var_pedido_credito = 1
                                     If Not rsaux.EOF Then
                                        var_pedido_credito = IIf(IsNull(rsaux!inte_ors_pedido_credito), 1, rsaux!inte_ors_pedido_credito)
                                     End If
                                     rsaux.Close
                                  End If
                                  rsaux.Open "update tb_encabezado_movimientos set VCHA_CAN_CANAL_VENTA_ID = '" + var_canal_venta + "', inte_emo_pedido_credito = " + CStr(var_pedido_credito) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                                  var_inserta = False
                                  var_inserta = TB_DET_EMBARQUE_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_numero_embarque, var_clave_movimiento, var_numero_folio, "")
                                  txt_folio = var_numero_folio
                                  var_primera_vez = False
                                  'End If
                                  'rsaux4.Close
                               Else
                                  rsaux4.Open "select * from tb_encabezado_movimientos where inte_emo_numero_origen = " + CStr(var_orden_surtido) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                                  If rsaux4.EOF Then
                                     var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, var_orden_surtido, var_clave_cliente, "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", CStr(rs!INTE_ORS_ORDEN_SURTIDO), var_clave_establecimiento, "B", var_clave_titular, var_clave_agente, var_descuento_1, var_descuento_2, var_descuento_3, var_clave_moneda, 0)
                                     var_numero_folio = var_numero_folio_regreso
                                     If var_factura_ceros = 1 Then
                                        rsaux.Open "update tb_encabezado_movimientos set inte_emo_factura_ceros = 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                                     End If
                                     var_pedido_credito = 1
                                     If var_clave_movimiento = "FT" Then
                                        rsaux.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + CStr(var_orden_surtido), cnn, adOpenDynamic, adLockOptimistic
                                        var_pedido_credito = 1
                                        If Not rsaux.EOF Then
                                           var_pedido_credito = IIf(IsNull(rsaux!inte_ors_pedido_credito), 1, rsaux!inte_ors_pedido_credito)
                                        End If
                                        rsaux.Close
                                     End If
                         
                         
                                     rsaux.Open "update tb_encabezado_movimientos set VCHA_CAN_CANAL_VENTA_ID = '" + var_canal_venta + "', inte_emo_pedido_credito = " + CStr(var_pedido_credito) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                                     var_inserta = False
                                     var_inserta = TB_DET_EMBARQUE_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_numero_embarque, var_clave_movimiento, var_numero_folio, "")
                                     txt_folio = var_numero_folio
                                     var_primera_vez = False
                                  End If
                                  rsaux4.Close
                               End If
                               rs.MoveNext
                         Wend
                         rs.Close
                      End If
                   End If
                   lv_salidas.selectedItem.SubItems(6) = "S"
                   var_renglon = lv_salidas.selectedItem.Index
                   If var_empresa = "03" Or var_agente_coppel = "00143" Then
                      If rsaux5.State = 1 Then
                         rsaux5.Close
                      End If
                      rsaux5.Open "select vcha_emp_empresa_id, min(inte_ors_orden_surtido) as inte_ors_orden_surtido from tb_detalle_cajas where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque + " group by vcha_emp_empresa_id", cnn, adOpenDynamic, adLockOptimistic
                      var_orden_surtido_03 = rsaux5!INTE_ORS_ORDEN_SURTIDO
                      rsaux5.Close
                   End If
                   rsaux4.Open "select * from tb_Detalle_cajas with (nolock)  where INTE_emb_embarque = " + txt_embarque + " and inte_paq_caja = " + CStr(var_caja) + " and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                   var_primera_vez = False
                   While Not rsaux4.EOF
                         If var_empresa = "03" Or var_agente_coppel = "00143" Then
                            var_orden_surtido = var_orden_surtido_03
                         Else
                            var_orden_surtido = rsaux4!INTE_ORS_ORDEN_SURTIDO
                         End If
                         Cadena = "SELECT  dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID, dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID, dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID, dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN, dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_DETALLE_EMBARQUES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID AND"
                         Cadena = Cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID AND "
                         Cadena = Cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO WHERE dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + txt_embarque + " AND dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN = " + CStr(var_orden_surtido)
                         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                         var_numero_folio = rs!INTE_SAL_NUMERO
                         rs.Close
                         If var_empresa = "03" Or var_agente_coppel = "00143" Then
                            var_orden_surtido = rsaux4!INTE_ORS_ORDEN_SURTIDO
                         End If
            
                         var_codigo = rsaux4!VCHA_ART_ARTICULO_ID
                         Cadena = "select * from tb_det_orden_surtido where inte_ors_orden_surtido = " + CStr(var_orden_surtido) + " and vcha_art_articulo_id = '" + var_codigo + "'"
                         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                         If Not rs.EOF Then
                            var_promocion_1 = IIf(IsNull(rs!floa_ors_promocion_1), 0, rs!floa_ors_promocion_1)
                            var_promocion_2 = IIf(IsNull(rs!floa_ors_promocion_2), 0, rs!floa_ors_promocion_2)
                            var_costo = IIf(IsNull(rs!floa_ors_costo), 0, rs!floa_ors_costo)
                            var_precio = IIf(IsNull(rs!floa_ors_precio), 0, rs!floa_ors_precio)
                            var_cantidad_leida = IIf(IsNull(rsaux4!floa_paq_cantidad), 0, rsaux4!floa_paq_cantidad)
                            If Trim(lbl_recibidos) <> "" Then
                               lbl_recibidos = Format(Int(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                            Else
                               lbl_recibidos = Format(var_cantidad_leida, "###,###,##0.00")
                            End If
                            var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                            ' se inibe porque ya fue leido al pasar las cajas la primera vez
                            'var_actualiza = TB_DET_ORDEN_SURTIDO_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_orden_surtido, CStr(var_codigo), var_cantidad_leida, 0 - var_cantidad_leida, var_precio, var_tipo_pedido)
               
                            If var_factura_ceros = 1 Then
                               var_precio = 0
                            End If
                            Cadena = "select * from TB_TEMPORAL_SALIDAS with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + var_codigo + "'"
                            rs.Close
               
                            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                            If Not rs.EOF Then
                               var_inserta = False
                               'rsaux.Open "update tb_temporal_salidas set floa_sal_cantidad = floa_sal_cantidad +" + CStr(var_cantidad_leida) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_Sal_Numero = " + CStr(var_numero_folio) + " and vcha_art_articulo_id = '" + var_codigo + "' and round(floa_sal_precio,2) = round(" + CStr(var_precio) + ",2) and char_ped_tipo = '" + var_tipo_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
                               rsaux.Open "update tb_temporal_salidas set floa_sal_cantidad = floa_sal_cantidad +" + CStr(var_cantidad_leida) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_Sal_Numero = " + CStr(var_numero_folio) + " and vcha_art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                               rs.Close
                            Else
                               var_inserta = False
                               rsaux.Open "INSERT INTO TB_TEMPORAL_SALIDAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, CHAR_PED_TIPO) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + var_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ",  " + CStr(var_precio) + ", 0,  " + CStr(var_promocion_1) + ", " + CStr(var_promocion_2) + ",'" + var_tipo_pedido + "') ", cnn, adOpenDynamic, adLockOptimistic
                               rs.Close
                            End If
                         Else
                            rs.Close
                         End If
                         rsaux4.MoveNext
                   Wend
                   rsaux4.Close
                   rsaux4.Open "update tb_detalle_cajas set char_paq_estatus = 'S' where INTE_emb_embarque = " + txt_embarque + " and inte_paq_caja = " + CStr(var_caja) + " AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                End If
            Next var_jj 'fin de la correccion
            Me.txt_codigo = ""
            MsgBox "El embarque se a corregido, vuelva a cerrar el embarque", vbOKOnly, "ATENCION"
         End If ' fin de la validacion de que la cantidad leida sea igual a la cantidad en cajas
         var_clave_movimiento = var_clave_movimiento_anterior
         
      Else
         MsgBox "El cerrado del embarque a sido cancelado", vbOKOnly, "ATENCION"
      End If
   End If
   frm_sellos.Visible = False
   Exit Sub
   frm_sellos.Visible = False
archivo_ocupado:
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If

   Exit Sub
   frm_sellos.Visible = False
End Sub

Private Sub cmd_cerrar_embarque_Click()
   Dim var_existen_cajas As Integer
   Dim var_numero_items As Integer
   lv_sellos.ListItems.Clear
   txt_sello = ""
   Text1 = Cadena
   If var_empresa <> "03" Then
      Cadena = "SELECT dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO, dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR, dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIDA , IsNull(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_NEGADA, 0) AS FLOA_ORS_CANTIDAD_NEGADA, dbo.VW_ORDENES_SURTIDO_DISTINTAS_CAJAS.VCHA_EMP_EMPRESA_ID, dbo.VW_ORDENES_SURTIDO_DISTINTAS_CAJAS.INTE_EMB_EMBARQUE FROM dbo.TB_DET_ORDEN_SURTIDO INNER JOIN dbo.VW_ORDENES_SURTIDO_DISTINTAS_CAJAS ON dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.VW_ORDENES_SURTIDO_DISTINTAS_CAJAS.INTE_ORS_ORDEN_SURTIDO"
      Cadena = Cadena + " WHERE (dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIDA + ISNULL(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_NEGADA, 0) < dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR) AND (dbo.VW_ORDENES_SURTIDO_DISTINTAS_CAJAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.VW_ORDENES_SURTIDO_DISTINTAS_CAJAS.INTE_EMB_EMBARQUE = " + Me.txt_embarque + ")"
   Else
      Cadena = "SELECT INTE_EMB_EMBARQUE, VCHA_EMP_EMPRESA_ID, INTE_ORS_ORDEN_SURTIDO, VCHA_ART_ARTICULO_ID, FLOA_ORS_CANTIDAD_SURTIR, FLOA_ORS_CANTIDAD_SURTIDA , FLOA_ORS_CANTIDAD_NEGADA From dbo.VW_ORDENES_SURTIDO_EMBARQUE WHERE     (FLOA_ORS_CANTIDAD_SURTIR > FLOA_ORS_CANTIDAD_SURTIDA + ISNULL(FLOA_ORS_CANTIDAD_NEGADA, 0)) AND (INTE_EMB_EMBARQUE = " + Me.txt_embarque + ") AND (VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')"
   End If
   Text1 = Cadena
   rsaux4.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux4.EOF Then
      rsaux4.Close
      frmasignacion_negado.txt_numero_embarque = Me.txt_embarque
      frmasignacion_negado.txt_agente = Me.txt_agente
      var_activa_forma_asignacion_negado = Me.Name
      var_negado_desde = 2
      frmasignacion_negado.Show
      Me.Enabled = False
   Else
      rsaux4.Close
      rs.Open "select * from tb_encabezado_embarques where inte_emb_embarque = " + Me.txt_embarque + " AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_embarque_cerrado = IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", Trim(rs!CHAR_EMB_ESTATUS))
      End If
      rs.Close
      rs.Open "select DISTINCT INTE_ORS_ORDEN_SURTIDO,INTE_PAQ_CAJA from tb_detalle_cajas where inte_emb_embarque = " + Me.txt_embarque + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and (char_paq_estatus <> 'S' and char_paq_estatus <> 'C')", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_existen_cajas = 1
      Else
         var_existen_cajas = 0
      End If
      rs.Close
      If var_existen_cajas = 0 Then
         If Trim(var_embarque_cerrado) = "E" Then
            rs.Open "select * from tb_Sellos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Str(var_numero_embarque), cnn, adOpenDynamic, adLockOptimistic
            var_numero_items = 0
            If Not rs.EOF Then
               While Not rs.EOF
                     Set list_item = lv_sellos.ListItems.Add(, , rs!vcha_sel_Sello)
                     rs.MoveNext
                     var_numero_items = var_numero_items + 1
               Wend
            End If
            If var_numero_items > 5 Then
               lv_sellos.ColumnHeaders(1).Width = 2650
            Else
               lv_sellos.ColumnHeaders(1).Width = 2850
            End If
            rs.Close
            frm_sellos.Visible = True
            txt_sello.SetFocus
         Else
            MsgBox "El embarque ya habia sido cerrado con anterioridad", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Faltan cajas sin subir", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_posible_kanban = 0
   Top = 0
   Left = 1500
   var_cantidad_enviada = 0
   var_cantidad_recibida = 0
   Me.frm_sellos.Visible = False
   Me.txt_embarque = frmnumero_embarque.txt_embarque
   Dim var_posible As Boolean
   var_posible = True
   'If var_posible_paqueteria = 1 Then
   '   rsaux10.Open "select isnull(vcha_paq_clave_id,'') as vcha_paq_clave_id, isnull(vcha_paq_guia,'') as vcha_paq_guia from tb_detalle_cajas where inte_Emb_embarque = " + txt_embarque + " and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   '   While Not rsaux10.EOF
   '         var_si_paqueteria = IIf(IsNull(rsaux10(0).Value), "", rsaux10(0).Value)
   '         var_si_guia = IIf(IsNull(rsaux10(1).Value), "", rsaux10(1).Value)
   '         If var_si_paqueteria = "" Then
   '            var_posible = False
   '         End If
   '         rsaux10.MoveNext
   '   Wend
   '   rsaux10.Close
   'Else
   '   var_posible = True
   'End If
   If var_posible = True Then
      rs.Open "select * from tb_encabezado_embarques where inte_emb_embarque = " + txt_embarque + " and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_estatus_embarque = Trim(IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", rs!CHAR_EMB_ESTATUS))
         rs.Close
         rs.Open "select a.char_paq_estatus,a.CHAR_PED_TIPO, a.inte_paq_caja, a.inte_ors_orden_surtido, sum(a.floa_paq_cantidad) as cantidad, a.vcha_alm_almacen_id, b.inte_ors_factura_ceros from tb_detalle_cajas a, tb_enc_orden_surtido b where a.vcha_emp_empresa_ID = '" + var_empresa + "' and a.inte_emb_embarque =  " + txt_embarque + " and a.inte_ors_orden_surtido = b.inte_ors_orden_surtido and char_paq_estatus <> 'C' group by a.vcha_emp_empresa_id, a.vcha_uor_unidad_id, a.vcha_alm_almacen_id, a.inte_emb_embarque, a.inte_ors_orden_surtido, a.inte_paq_caja, b.inte_ors_factura_ceros, a.CHAR_PED_TIPO, a.char_paq_estatus", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux.Open "select * from tb_detalle_cajas with (nolock)  where inte_emb_embarque = " + txt_embarque + " and vcha_emp_empresa_id = '" + var_empresa + "' and inte_paq_caja = " + CStr(rs!inte_paq_caja), cnn, adOpenDynamic, adLockOptimistic
            rsaux2.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               var_almacen_origen = rs!VCHA_ALM_ALMACEN_ID
               Me.txt_origen = rsaux2!VCHA_ALM_NOMBRE
               rsaux3.Open "SELECT * FROM TB_ENC_ORDEN_SURTIDO WHERE INTE_ORS_ORDEN_SURTIDO = " + CStr(rsaux!INTE_ORS_ORDEN_SURTIDO), cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux3.EOF Then
                  rsaux4.Open "select * from tb_encabezado_pedidos where inte_ped_numero = " + CStr(rsaux3!inte_ped_numero), cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux4.EOF Then
                     var_clave_agente = IIf(IsNull(rsaux4!VCHA_AGE_AGENTE_ID), "", rsaux4!VCHA_AGE_AGENTE_ID)
                  End If
                  rsaux4.Close
                  rsaux4.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + var_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux4.EOF Then
                     txt_agente = IIf(IsNull(rsaux4!VCHA_AGE_NOMBRE), "", rsaux4!VCHA_AGE_NOMBRE)
                     var_i = 1
                     While Not rs.EOF
                           var_numero_caja = IIf(IsNull(inte_paq_caja), 0, rs!inte_paq_caja)
                           If Len(Trim(Str(var_numero_caja))) = 1 Then
                              var_referencia_caja = "00" + Trim(Str(var_numero_caja))
                           End If
                           If Len(Trim(Str(var_numero_caja))) = 2 Then
                              var_referencia_caja = "0" + Trim(Str(var_numero_caja))
                           End If
                           If Len(Trim(Str(var_numero_caja))) = 3 Then
                              var_referencia_caja = Trim(Str(var_numero_caja))
                           End If
                           If Len(Trim(Str(txt_embarque))) = 1 Then
                              var_referencia_embarque = "00000" + Trim(Str(txt_embarque))
                           End If
                           If Len(Trim(Str(txt_embarque))) = 2 Then
                              var_referencia_embarque = "0000" + Trim(Str(txt_embarque))
                           End If
                           If Len(Trim(Str(txt_embarque))) = 3 Then
                              var_referencia_embarque = "000" + Trim(Str(txt_embarque))
                           End If
                           If Len(Trim(Str(txt_embarque))) = 4 Then
                              var_referencia_embarque = "00" + Trim(Str(txt_embarque))
                           End If
                           If Len(Trim(Str(txt_embarque))) = 5 Then
                              var_referencia_embarque = "0" + Trim(Str(txt_embarque))
                           End If
                           var_codigo_caja = "C" + var_referencia_embarque + var_referencia_caja
                           Set list_item = lv_salidas.ListItems.Add(, , var_codigo_caja)
                           list_item.SubItems(1) = IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                           list_item.SubItems(2) = IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), 0, rs!INTE_ORS_ORDEN_SURTIDO)
                           list_item.SubItems(3) = IIf(IsNull(rs!inte_paq_caja), 0, rs!inte_paq_caja)
                           list_item.SubItems(4) = IIf(IsNull(rs!inte_ors_factura_ceros), 0, rs!inte_ors_factura_ceros)
                           list_item.SubItems(5) = IIf(IsNull(rs!char_ped_tipo), "", rs!char_ped_tipo)
                           list_item.SubItems(6) = IIf(IsNull(rs!char_paq_estatus), "", rs!char_paq_estatus)
                           var_estatus_caja = IIf(IsNull(rs!char_paq_estatus), "", rs!char_paq_estatus)
                           var_cantidad_enviada = var_cantidad_enviada + rs!Cantidad
                           If var_estatus_caja = "S" Then
                              var_cantidad_recibida = var_cantidad_recibida + rs!Cantidad
                           End If
                           var_i = var_i + 1
                           rs.MoveNext
                     Wend
                     var_renglon = -1
                     Call ilumina_grid
                     lbl_recibidos = Format(var_cantidad_recibida, "###,###,##0.00")
                     lbl_enviados = Format(var_cantidad_enviada, "###,###,##0.00")
                  Else
                     MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
                  End If
                  rsaux4.Close
               Else
                  MsgBox "Orden de surtido incorrecta", vbOKOnly, "ATENCION"
               End If
               rsaux3.Close
            Else
               MsgBox "Clave de almacen incorrecta", vbOKOnly, "ATENCION"
            End If
            rsaux.Close
            rsaux2.Close
         End If
         rs.Close
         If var_estatus_embarque = "E" Then
            rs.Open "select * from tb_detalle_embarques where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_embarque, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_primera_vez = False
            Else
               var_primera_vez = True
            End If
            rs.Close
            txt_codigo.Enabled = True
            txt_foco.Enabled = False
         Else
            If var_estatus_embarque = "" Then
               MsgBox "El embarque no a sido cerrado en el modulo de de creación de cajas", vbOKOnly, "ATENCION"
            Else
               If var_estatus_embarque = "I" Then
                  MsgBox "El embarque ya fue cerrado", vbOKOnly
               Else
                  MsgBox "El embarque ya fue facturado"
               End If
            End If
            var_primera_vez = False
            txt_codigo.Enabled = False
            txt_foco.Enabled = False
         End If
      Else
         rs.Close
         MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
      End If
   Else
      Me.txt_codigo.Enabled = False
      Me.txt_foco.Enabled = False
      MsgBox "No se a indicado la paqueteria", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_salidas_cajas)
End Sub

Private Sub txt_cantidad_eliminar_Change()

End Sub

Private Sub txt_cantidad_eliminar_GotFocus()
   txt_cantidad_eliminar = ""
End Sub


Private Sub lv_salidas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_salidas, ColumnHeader)
End Sub

Private Sub lv_salidas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If var_estatus_embarque = "E" Then
         If lv_salidas.ListItems.Count > 0 Then
            txt_cantidad_eliminar = lv_salidas.selectedItem
            If Trim(txt_cantidad_eliminar) <> "" Then
               x = Mid(txt_cantidad_eliminar, 2, 6)
               If IsNumeric(x) Then
                  var_embarque_paquete = x
                  x = Mid(txt_cantidad_eliminar, 8, 3)
                  If IsNumeric(x) Then
                     var_embarque_caja = x
                     var_posible_caja = True
                  Else
                     var_posible_caja = False
                  End If
               Else
                  var_posible_caja = False
               End If
               If var_posible_caja = True Then
                  var_si = MsgBox("     ¿Deseas eliminar la caja?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     var_estatus_caja = lv_salidas.selectedItem.SubItems(6)
                     If var_estatus_caja = "S" Then
                        var_embarque_paquete = txt_embarque
                        rsaux3.Open "select * from tb_detalle_cajas with (nolock)  where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and INTE_EMB_EMBARQUE = " + CStr(var_numero_embarque) + " and inte_paq_caja = " + Str(var_embarque_caja), cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           If rsaux3!char_paq_estatus = "S" Then
                              'cnn.BeginTrans
                              Set TB_DETALLE_CAJAS_M = New TB_DETALLE_CAJAS_M
                              ok = False
                              txt_archivo = lv_salidas.selectedItem.SubItems(2)
                              var_orden_surtido = lv_salidas.selectedItem.SubItems(2) * 1
                              ok = TB_DETALLE_CAJAS_M.Anadir(CDbl(txt_archivo), CInt(var_embarque_caja), var_empresa, var_unidad_organizacional, var_almacen_origen, "I", "", 0, var_numero_embarque)
                              
                              Cadena = "SELECT  dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID, dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID, dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID, dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN, dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_DETALLE_EMBARQUES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID AND"
                              Cadena = Cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID AND "
                              Cadena = Cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO WHERE dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + txt_embarque + " AND dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN = " + CStr(var_orden_surtido)
                              rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                              var_numero_folio = rs!INTE_SAL_NUMERO
                              rs.Close
                              While Not rsaux3.EOF
                                    Set TB_DET_ORDEN_SURTIDO_M = New TB_DET_ORDEN_SURTIDO_M
                                    Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
                                    var_codigo = rsaux3!VCHA_ART_ARTICULO_ID
                                    var_precio = rsaux3!floa_paq_precio
                                    var_tipo_pedido = rsaux3!char_ped_tipo
                                    var_n = lv_salidas.ListItems.Count
                                    var_encontro = 0
                                    var_i = 1
                                    var_cantidad_eliminar = rsaux3!floa_paq_cantidad
                                    var_actualiza = TB_DET_ORDEN_SURTIDO_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, CDbl(var_orden_surtido), CStr(var_codigo), 0 - var_cantidad_eliminar, var_cantidad_eliminar, var_precio, var_tipo_pedido)
                                    rsaux.Open "update tb_temporal_salidas set floa_sal_cantidad = floa_sal_cantidad -" + CStr(var_cantidad_eliminar) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_Sal_Numero = " + CStr(var_numero_folio) + " and vcha_art_articulo_id = '" + var_codigo + "' and round(floa_sal_precio,2) = round(" + CStr(var_precio) + ",2) AND CHAR_PED_TIPO = '" + var_tipo_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
                                    lbl_recibidos = Format(Int(lbl_recibidos) - var_cantidad_eliminar, "###,###,##0.00")
                                    txt_codigo.SetFocus
                                    rsaux3.MoveNext
                              Wend
                              'cnn.CommitTrans
                              var_renglon = lv_salidas.selectedItem.Index
                              lv_salidas.selectedItem.SubItems(6) = "I"
                              Call ilumina_grid
                           End If
                        End If
                        rsaux3.Close
                     Else
                        MsgBox "La caja no a sido surtida", vbOKOnly, "ATENCION"
                     End If
                  End If
               End If
            End If
         End If
      Else
         MsgBox "El embarque ya no puede ser modificado", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_codigo_GotFocus()
   txt_codigo = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(txt_codigo) <> "" Then
         Set itmfound = lv_salidas.findItem(Trim(txt_codigo), lvwText, , lvwPartial)
         If itmfound Is Nothing Then
            txt_codigo = ""
            frmmensaje.lbl_mensaje = "La caja no se encuentra en el embarque"
            frmmensaje.Show 1
            txt_codigo.SetFocus
            var_orden_surtido = 0
            var_caja = 0
            var_factura_ceros = 0
            var_tipo_pedido = ""
         Else
            itmfound.EnsureVisible
            itmfound.Selected = True
            If lv_salidas.selectedItem.SubItems(6) = "S" Then
               frmmensaje.lbl_mensaje = "La caja ya fue surtida"
               frmmensaje.Show 1
               txt_codigo.SetFocus
               var_orden_surtido = 0
               var_caja = 0
               var_factura_ceros = 0
               var_tipo_pedido = ""
            Else
               var_orden_surtido = lv_salidas.selectedItem.SubItems(2)
               var_caja = lv_salidas.selectedItem.SubItems(3)
               var_factura_ceros = lv_salidas.selectedItem.SubItems(4)
               var_tipo_pedido = lv_salidas.selectedItem.SubItems(5)
               txt_foco.Enabled = True
               txt_foco.SetFocus
            End If
         End If
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   Dim var_actualiza As Boolean
   Dim var_inserta As Boolean
   Dim bandera_suma As Boolean
   Dim var_cantidad As Variant
   Dim var_costo As Variant
   Dim var_precio As Variant
   Dim var_posible_caja As Boolean
   Dim var_cantidad_posible As Variant
   Dim var_embarque_paquete As Integer
   Dim var_embarque_caja As Integer
   Dim var_estatus_caja As String
   Dim var_orden_surtido_caja As Double
   Dim var_posible_empaque As Boolean
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   Dim var_encontrado As Integer
   Dim var_canal_venta As String
   Dim var_i As Integer
   Dim var_n As Integer
   Dim var_j As Integer
   Dim var_numero_pedido As Double
   Dim var_orden_surtido_03 As Double
   Dim var_agente_coppel As String
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_DET_ORDEN_SURTIDO_M = New TB_DET_ORDEN_SURTIDO_M
   Set TB_DET_EMBARQUE_I = New TB_DET_EMBARQUE_I
   Set TB_DETALLE_CAJAS_M = New TB_DETALLE_CAJAS_M
   'On Error GoTo salir:
   z = 0
   cnn.CommandTimeout = 360
   If Trim(txt_codigo.Text) <> "" Then
      'var_primera_vez = True
      rs.Open "select vcha_age_Agente_id from tb_encabezado_embarques where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque
      If Not rs.EOF Then
         var_agente_coppel = rs!VCHA_AGE_AGENTE_ID
      End If
      rs.Close
      If var_clave_titular <> "T000001474" Then
         var_agente_coppel = ""
      End If
      If var_primera_vez = True Then
         var_inserta = False
         var_agente_coppel = ""
         If var_empresa = "03" Or var_agente_coppel = "00143" Then
            rs.Open "select vcha_emp_empresa_id, min(inte_ors_orden_surtido) as inte_ors_orden_surtido from tb_detalle_cajas where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque + " group by vcha_emp_empresa_id", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux4.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + CStr(rs!INTE_ORS_ORDEN_SURTIDO), cnn, adOpenDynamic, adLockOptimistic
                  var_orden_surtido = 0
                  var_clave_cliente = ""
                  var_almacen_Destino = ""
                  var_clave_establecimiento = ""
                  var_clave_titular = ""
                  var_descuento_1 = 0
                  var_descuento_2 = 0
                  var_clave_moneda = 0
                  var_orden_surtido = rs!INTE_ORS_ORDEN_SURTIDO
                  var_clave_cliente = rsaux4!vcha_cli_clave_id
                  var_almacen_Destino = ""
                  var_clave_establecimiento = rsaux4!vcha_ESB_ESTABLECIMIENTO_id
                  var_clave_titular = rsaux4!vcha_tit_titular_id
                  var_descuento_1 = IIf(IsNull(rsaux4!FLOA_ORS_DESCUENTO_1), 0, rsaux4!FLOA_ORS_DESCUENTO_1)
                  var_descuento_2 = IIf(IsNull(rsaux4!FLOA_ORS_DESCUENTO_2), 0, rsaux4!FLOA_ORS_DESCUENTO_2)
                  var_clave_moneda = IIf(IsNull(rsaux4!vcha_mon_moneda_id), "", rsaux4!vcha_mon_moneda_id)
                  rsaux4.Close
                  rsaux.Open "select vcha_can_canal_venta_id from tb_agentes where vcha_age_agente_id = '" + var_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_canal_venta = IIf(IsNull(rsaux!vcha_can_canal_venta_id), "", rsaux!vcha_can_canal_venta_id)
                  rsaux.Close
                  If var_empresa = "03" Or var_agente_coppel = "00143" Then
                      var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, var_orden_surtido, var_clave_cliente, "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", CStr(rs!INTE_ORS_ORDEN_SURTIDO), var_clave_establecimiento, "B", var_clave_titular, var_clave_agente, var_descuento_1, var_descuento_2, var_descuento_3, var_clave_moneda, 0)
                      var_numero_folio = var_numero_folio_regreso
                      If var_factura_ceros = 1 Then
                         rsaux.Open "update tb_encabezado_movimientos set inte_emo_factura_ceros = 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                      End If
                      var_pedido_credito = 1
                      
                      If var_clave_movimiento = "FT" Then
                         rsaux.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + CStr(var_orden_surtido), cnn, adOpenDynamic, adLockOptimistic
                         var_pedido_credito = 1
                         If Not rsaux.EOF Then
                            var_pedido_credito = IIf(IsNull(rsaux!inte_ors_pedido_credito), 1, rsaux!inte_ors_pedido_credito)
                         End If
                         rsaux.Close
                      End If
                      
                      rsaux.Open "update tb_encabezado_movimientos set VCHA_CAN_CANAL_VENTA_ID = '" + var_canal_venta + "', inte_emo_pedido_credito = " + CStr(var_pedido_credito) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                      var_inserta = False
                      var_inserta = TB_DET_EMBARQUE_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_numero_embarque, var_clave_movimiento, var_numero_folio, "")
                      txt_folio = var_numero_folio
                      var_primera_vez = False
                     'End If
                     'rsaux4.Close
                  Else
                     rsaux4.Open "select * from tb_encabezado_movimientos where inte_emo_numero_origen = " + CStr(var_orden_surtido) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                     If rsaux4.EOF Then
                        var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, var_orden_surtido, var_clave_cliente, "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", CStr(rs!INTE_ORS_ORDEN_SURTIDO), var_clave_establecimiento, "B", var_clave_titular, var_clave_agente, var_descuento_1, var_descuento_2, var_descuento_3, var_clave_moneda, 0)
                        var_numero_folio = var_numero_folio_regreso
                        If var_factura_ceros = 1 Then
                            rsaux.Open "update tb_encabezado_movimientos set inte_emo_factura_ceros = 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        var_pedido_credito = 1
                        If var_clave_movimiento = "FT" Then
                           rsaux.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + CStr(var_orden_surtido), cnn, adOpenDynamic, adLockOptimistic
                           var_pedido_credito = 1
                           If Not rsaux.EOF Then
                              var_pedido_credito = IIf(IsNull(rsaux!inte_ors_pedido_credito), 1, rsaux!inte_ors_pedido_credito)
                           End If
                           rsaux.Close
                        End If
                        
                        rsaux.Open "update tb_encabezado_movimientos set VCHA_CAN_CANAL_VENTA_ID = '" + var_canal_venta + "', inte_emo_pedido_credito = " + CStr(var_pedido_credito) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_inserta = False
                        var_inserta = TB_DET_EMBARQUE_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_numero_embarque, var_clave_movimiento, var_numero_folio, "")
                        txt_folio = var_numero_folio
                        var_primera_vez = False
                     End If
                     rsaux4.Close
                  End If
                  rs.MoveNext
            Wend
            rs.Close
         Else
            rs.Open "select distinct vcha_emp_empresa_id, inte_ors_orden_surtido from tb_detalle_cajas where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux4.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + CStr(rs!INTE_ORS_ORDEN_SURTIDO), cnn, adOpenDynamic, adLockOptimistic
                  var_orden_surtido = 0
                  var_clave_cliente = ""
                  var_almacen_Destino = ""
                  var_clave_establecimiento = ""
                  var_clave_titular = ""
                  var_descuento_1 = 0
                  var_descuento_2 = 0
                  var_clave_moneda = 0
                  var_orden_surtido = rs!INTE_ORS_ORDEN_SURTIDO
                  var_clave_cliente = rsaux4!vcha_cli_clave_id
                  var_almacen_Destino = ""
                  var_clave_establecimiento = rsaux4!vcha_ESB_ESTABLECIMIENTO_id
                  var_clave_titular = rsaux4!vcha_tit_titular_id
                  var_descuento_1 = IIf(IsNull(rsaux4!FLOA_ORS_DESCUENTO_1), 0, rsaux4!FLOA_ORS_DESCUENTO_1)
                  var_descuento_2 = IIf(IsNull(rsaux4!FLOA_ORS_DESCUENTO_2), 0, rsaux4!FLOA_ORS_DESCUENTO_2)
                  var_clave_moneda = IIf(IsNull(rsaux4!vcha_mon_moneda_id), "", rsaux4!vcha_mon_moneda_id)
                  rsaux4.Close
                  rsaux.Open "select vcha_can_canal_venta_id from tb_agentes where vcha_age_agente_id = '" + var_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_canal_venta = IIf(IsNull(rsaux!vcha_can_canal_venta_id), "", rsaux!vcha_can_canal_venta_id)
                  rsaux.Close
                  If var_clave_titular <> "T000001474" Then
                     var_agente_coppel = ""
                  End If
                  If var_empresa = "03" Or var_agente_coppel = "00143" Then
                      var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, var_orden_surtido, var_clave_cliente, "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", CStr(rs!INTE_ORS_ORDEN_SURTIDO), var_clave_establecimiento, "B", var_clave_titular, var_clave_agente, var_descuento_1, var_descuento_2, var_descuento_3, var_clave_moneda, 0)
                      var_numero_folio = var_numero_folio_regreso
                      If var_factura_ceros = 1 Then
                         rsaux.Open "update tb_encabezado_movimientos set inte_emo_factura_ceros = 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                      End If
                      var_pedido_credito = 1
                      If var_clave_movimiento = "FT" Then
                         rsaux.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + CStr(var_orden_surtido), cnn, adOpenDynamic, adLockOptimistic
                         var_pedido_credito = 1
                         If Not rsaux.EOF Then
                            var_pedido_credito = IIf(IsNull(rsaux!inte_ors_pedido_credito), 1, rsaux!inte_ors_pedido_credito)
                         End If
                         rsaux.Close
                      End If
                      rsaux.Open "update tb_encabezado_movimientos set VCHA_CAN_CANAL_VENTA_ID = '" + var_canal_venta + "', inte_emo_pedido_credito = " + CStr(var_pedido_credito) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                      var_inserta = False
                      var_inserta = TB_DET_EMBARQUE_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_numero_embarque, var_clave_movimiento, var_numero_folio, "")
                      txt_folio = var_numero_folio
                      var_primera_vez = False
                  Else
                     rsaux4.Open "select * from tb_encabezado_movimientos where inte_emo_numero_origen = " + CStr(var_orden_surtido) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                     If rsaux4.EOF Then
                        var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, var_orden_surtido, var_clave_cliente, "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", CStr(rs!INTE_ORS_ORDEN_SURTIDO), var_clave_establecimiento, "B", var_clave_titular, var_clave_agente, var_descuento_1, var_descuento_2, var_descuento_3, var_clave_moneda, 0)
                        var_numero_folio = var_numero_folio_regreso
                        If var_factura_ceros = 1 Then
                            rsaux.Open "update tb_encabezado_movimientos set inte_emo_factura_ceros = 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        var_pedido_credito = 1
                        If var_clave_movimiento = "FT" Then
                           rsaux.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + CStr(var_orden_surtido), cnn, adOpenDynamic, adLockOptimistic
                           var_pedido_credito = 1
                           If Not rsaux.EOF Then
                              var_pedido_credito = IIf(IsNull(rsaux!inte_ors_pedido_credito), 1, rsaux!inte_ors_pedido_credito)
                           End If
                           rsaux.Close
                        End If
                        rsaux.Open "update tb_encabezado_movimientos set VCHA_CAN_CANAL_VENTA_ID = '" + var_canal_venta + "', inte_emo_pedido_credito = " + CStr(var_pedido_credito) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_inserta = False
                        var_inserta = TB_DET_EMBARQUE_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_numero_embarque, var_clave_movimiento, var_numero_folio, "")
                        txt_folio = var_numero_folio
                        var_primera_vez = False
                     End If
                     rsaux4.Close
                  End If
                  rs.MoveNext
            Wend
            rs.Close
         End If
      End If
      lv_salidas.selectedItem.SubItems(6) = "S"
      var_renglon = lv_salidas.selectedItem.Index
      Call ilumina_grid
      If var_empresa = "03" Or var_agente_coppel = "00143" Then
         If rsaux5.State = 1 Then
            rsaux5.Close
         End If
         rsaux5.Open "select vcha_emp_empresa_id, min(inte_ors_orden_surtido) as inte_ors_orden_surtido from tb_detalle_cajas where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque + " group by vcha_emp_empresa_id", cnn, adOpenDynamic, adLockOptimistic
         var_orden_surtido_03 = rsaux5!INTE_ORS_ORDEN_SURTIDO
         rsaux5.Close
      End If
      rsaux4.Open "select * from tb_Detalle_cajas with (nolock)  where INTE_emb_embarque = " + txt_embarque + " and inte_paq_caja = " + CStr(var_caja) + " and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      var_primera_vez = False
      While Not rsaux4.EOF
            If var_empresa = "03" Or var_agente_coppel = "00143" Then
               var_orden_surtido = var_orden_surtido_03
            Else
               var_orden_surtido = rsaux4!INTE_ORS_ORDEN_SURTIDO
            End If
            Cadena = "SELECT  dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID, dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID, dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID, dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN, dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_DETALLE_EMBARQUES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID AND"
            Cadena = Cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID AND "
            Cadena = Cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO WHERE dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + txt_embarque + " AND dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN = " + CStr(var_orden_surtido)
            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
           ' MsgBox CStr(rsaux4!inte_ors_orden_surtido)
            var_numero_folio = rs!INTE_SAL_NUMERO
            rs.Close
            If var_empresa = "03" Or var_agente_coppel = "00143" Then
               var_orden_surtido = rsaux4!INTE_ORS_ORDEN_SURTIDO
            End If
            var_codigo = rsaux4!VCHA_ART_ARTICULO_ID
            Cadena = "select * from tb_det_orden_surtido where inte_ors_orden_surtido = " + CStr(var_orden_surtido) + " and vcha_art_articulo_id = '" + var_codigo + "'"
            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_promocion_1 = IIf(IsNull(rs!floa_ors_promocion_1), 0, rs!floa_ors_promocion_1)
               var_promocion_2 = IIf(IsNull(rs!floa_ors_promocion_2), 0, rs!floa_ors_promocion_2)
               var_costo = IIf(IsNull(rs!floa_ors_costo), 0, rs!floa_ors_costo)
               var_precio = IIf(IsNull(rs!floa_ors_precio), 0, rs!floa_ors_precio)
               var_cantidad_leida = IIf(IsNull(rsaux4!floa_paq_cantidad), 0, rsaux4!floa_paq_cantidad)
               If Trim(lbl_recibidos) <> "" Then
                  lbl_recibidos = Format(Int(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
               Else
                  lbl_recibidos = Format(var_cantidad_leida, "###,###,##0.00")
               End If
               var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
               var_actualiza = TB_DET_ORDEN_SURTIDO_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_orden_surtido, CStr(var_codigo), var_cantidad_leida, 0 - var_cantidad_leida, var_precio, var_tipo_pedido)
               
               If var_factura_ceros = 1 Then
                  var_precio = 0
               End If
               Cadena = "select * from TB_TEMPORAL_SALIDAS with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + var_codigo + "'"
               rs.Close
               
               rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_inserta = False
                  rsaux.Open "update tb_temporal_salidas set floa_sal_cantidad = floa_sal_cantidad +" + CStr(var_cantidad_leida) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_Sal_Numero = " + CStr(var_numero_folio) + " and vcha_art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  rs.Close
               Else
                  var_inserta = False
                  rsaux.Open "INSERT INTO TB_TEMPORAL_SALIDAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, CHAR_PED_TIPO) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + var_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ",  " + CStr(var_precio) + ", 0,  " + CStr(var_promocion_1) + ", " + CStr(var_promocion_2) + ",'" + var_tipo_pedido + "') ", cnn, adOpenDynamic, adLockOptimistic
                  rs.Close
               End If
               
            Else
               rs.Close
            End If
            rsaux4.MoveNext
      Wend
      rsaux4.Close
      rsaux4.Open "update tb_detalle_cajas set char_paq_estatus = 'S' where INTE_emb_embarque = " + txt_embarque + " and inte_paq_caja = " + CStr(var_caja) + " AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   If txt_codigo.Enabled = True Then
      txt_codigo.SetFocus
   End If
   Exit Sub
salir:
Resume
End Sub

Private Sub txt_sello_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmd_aceptar_sello.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_sellos.Visible = False
   End If
End Sub
