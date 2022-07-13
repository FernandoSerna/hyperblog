VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmfactura_empresas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11250
      Picture         =   "frmfactura_empresas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Salir"
      Top             =   585
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   780
      Picture         =   "frmfactura_empresas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Imprimir Movimiento Alt + I"
      Top             =   585
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   450
      Picture         =   "frmfactura_empresas.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Buscar Movimiento Alt + B"
      Top             =   585
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmfactura_empresas.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Nuevo Movimiento Alt + N"
      Top             =   585
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   1545
      Index           =   2
      Left            =   90
      TabIndex        =   22
      Top             =   3975
      Width           =   5610
      Begin VB.ComboBox cmb_series 
         Height          =   315
         Left            =   585
         TabIndex        =   34
         Top             =   630
         Width           =   795
      End
      Begin VB.TextBox txt_de 
         Height          =   315
         Left            =   1950
         TabIndex        =   25
         Top             =   630
         Width           =   1410
      End
      Begin VB.TextBox txt_a 
         Height          =   315
         Left            =   4065
         TabIndex        =   24
         Top             =   630
         Width           =   1410
      End
      Begin VB.TextBox txt_renglones 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4065
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   975
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   35
         Top             =   690
         Width           =   405
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   " Facturas Sugeridas"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   29
         Top             =   120
         Width           =   5535
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   1650
         TabIndex        =   28
         Top             =   690
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   195
         Left            =   3765
         TabIndex        =   27
         Top             =   690
         Width           =   150
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Total de Facturas a Imprimir:"
         Height          =   195
         Left            =   1920
         TabIndex        =   26
         Top             =   1035
         Width           =   1995
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   60
      TabIndex        =   18
      Top             =   450
      Width           =   11580
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   840
      Width           =   11580
   End
   Begin VB.Frame Frame2 
      Height          =   4590
      Left            =   5730
      TabIndex        =   15
      Top             =   930
      Width           =   5850
      Begin MSComctlLib.ListView lv_piezas 
         Height          =   4125
         Left            =   45
         TabIndex        =   16
         Top             =   390
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   7276
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
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Costo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Precio"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Relación de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   17
         Top             =   120
         Width           =   5775
      End
   End
   Begin VB.TextBox txt_clave_movimiento 
      Height          =   285
      Left            =   2205
      TabIndex        =   14
      Top             =   675
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame3 
      Height          =   1500
      Index           =   0
      Left            =   90
      TabIndex        =   12
      Top             =   930
      Width           =   5595
      Begin VB.TextBox txt_archivo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2175
         TabIndex        =   1
         Top             =   900
         Width           =   2010
      End
      Begin VB.TextBox txt_numero_folio 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2175
         TabIndex        =   0
         Top             =   495
         Width           =   2010
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Archivo:"
         Height          =   195
         Index           =   2
         Left            =   375
         TabIndex        =   21
         Top             =   1020
         Width           =   585
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Número de Movimiento:"
         Height          =   195
         Index           =   1
         Left            =   345
         TabIndex        =   20
         Top             =   585
         Width           =   1680
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   1
         Left            =   30
         TabIndex        =   13
         Top             =   120
         Width           =   5520
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1500
      Index           =   1
      Left            =   90
      TabIndex        =   7
      Top             =   2445
      Width           =   5610
      Begin VB.ComboBox cmb_clientes 
         Height          =   315
         Left            =   2010
         TabIndex        =   4
         Top             =   735
         Width           =   3495
      End
      Begin VB.TextBox txt_descuentos 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   5
         Top             =   1065
         Width           =   2370
      End
      Begin VB.TextBox txt_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   3
         Top             =   735
         Width           =   960
      End
      Begin VB.TextBox txt_almacen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   2
         Top             =   390
         Width           =   4485
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descuentos:"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   1125
         Width           =   900
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   795
         Width           =   525
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   9
         Top             =   120
         Width           =   5535
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   8
         Top             =   450
         Width           =   660
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   1215
      Top             =   30
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
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   75
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":0940
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":121A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":1AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":2090
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":296C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":3246
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":3B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":3C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":3D44
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":3E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":3F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":407A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":418C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":432E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":5180
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":5356
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":5468
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":66EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":67FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfactura_empresas.frx":7A7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog comdialog 
      Left            =   -15
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   165
      TabIndex        =   19
      Top             =   15
      Width           =   11445
   End
End
Attribute VB_Name = "frmfactura_empresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tabla As ADODB.Connection
Dim var_primera_vez As Boolean
Dim var_numero_folio As Double
Dim var_clave_cliente As String
Dim var_titular As String
Dim var_agente As String
Dim var_almacen As String
Dim var_numero_renglones As Integer
Dim var_descuento_1 As Double
Dim var_descuento_2 As Double
Dim var_descuento_3 As Double
Dim var_lista_precios As String
Dim var_numero_factura As Double
Dim var_serie As String

Private Sub cmb_clientes_Click()
   txt_cliente = Obtener_llave(cnn, rsaux, "TB_clientes", "VCHA_cli_NOMBRE", cmb_clientes, 0, "T")
   var_clave_cliente = txt_cliente
   If rs.State = 1 Then
      rs.Close
   End If
   var_descuento_1 = 0
   var_descuento_2 = 0
   var_descuento_3 = 0
   var_titular = ""
   var_agente = ""
   var_lista_precios = 0
   rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      var_descuento_1 = IIf(IsNull(rs!floa_gac_Descuento_1), 0, rs!floa_gac_Descuento_1)
      var_descuento_2 = IIf(IsNull(rs!FLOA_GAC_DESCUENTO_2), 0, rs!FLOA_GAC_DESCUENTO_2)
      var_descuento_3 = IIf(IsNull(rs!floa_gac_descuento_3), 0, rs!floa_gac_descuento_3)
      txt_descuentos = Str(var_descuento_1) + " % " + Str(var_descuento_2) + " % " + Str(var_descuento_3) + " %"
      var_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
      var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
      var_lista_precios = IIf(IsNull(rs!vcha_LIS_LISTA_iD), "", rs!vcha_LIS_LISTA_iD)
   rs.Close
End Sub

Private Sub cmb_clientes_KeyPress(KeyAscii As Integer)
   cmb_clientes.Enabled = False
End Sub

Private Sub cmb_clientes_LostFocus()
   cmb_clientes.Enabled = False
End Sub

Private Sub cmb_series_Click()
   var_serie = cmb_series
   var_descuento_1 = 0
   var_descuento_2 = 0
   var_descuento_3 = 0
   If Trim(txt_archivo) <> "" Then
      rs.Open "select inte_ser_factura from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
      var_numero_factura = IIf(IsNull(rs!inte_ser_factura), 0, rs!inte_ser_factura) + 1
      rs.Close
      rs.Open "select * from tb_archivo_comparacion where vcha_com_referencia = '" + txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
      rsaux2.Open "select * from tb_almacenes where vcha_alm_almacen_id ='" + rs!VCHA_ALM_ALMACEN_ID + "'", cnn, adOpenDynamic, adLockOptimistic
      txt_almacen = rsaux2!VCHA_ALM_NOMBRE
      var_almacen = rsaux2!VCHA_ALM_ALMACEN_ID
      rsaux2.Close
      If Not rs.EOF Then
         var_contador = 0
         var_contador_facturas = 1
         lv_piezas.ListItems.Clear
         While Not rs.EOF
            var_contador = var_contador + 1
            Dim list_item As ListItem
            Set list_item = lv_piezas.ListItems.Add(, , Trim(rs!vcha_Art_articulo_id))
            rsaux2.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
            list_item.SubItems(1) = IIf(IsNull(rsaux2!vcha_art_nombre_español), "", Trim(rsaux2!vcha_art_nombre_español))
            rsaux2.Close
            list_item.SubItems(2) = IIf(IsNull(rs!FLOA_COM_CANTIDAD_ENVIADA), 0, rs!FLOA_COM_CANTIDAD_ENVIADA)
            rs.MoveNext
            If var_contador < var_numero_renglones And rs.EOF = True And var_contador_facturas <> 1 Then
               var_contador_facturas = var_contador_facturas + 1
            End If
            If var_contador = var_numero_renglones Then
               var_contador = 0
               If Not rs.EOF Then
                  var_contador_facturas = var_contador_facturas + 1
               End If
            End If
         Wend
         txt_de = var_numero_factura
         txt_a = var_numero_factura + var_contador_facturas - 1
         txt_renglones = var_contador_facturas
      Else
         MsgBox "El archivo no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub cmd_buscar_Click()
   x = 1 + 1
End Sub

Private Sub cmd_imprimir_Click()
Dim var_precio As Double
Dim var_posible As Boolean
Dim var_clave_moneda As String
Dim var_tipo_Cambio As Double
Dim var_promocion_1 As Double
Dim var_promocion_2 As Double
Dim var_marca_promocion As String
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
If Trim(var_lista_precios) <> "" Then
   If Trim(txt_archivo) <> "" Then
      si = MsgBox("¿Deseas cerrar el movimiento?", vbYesNo, "ATENCION")
      If si = 6 Then
         var_posible = True
         rsaux2.Open "select * from TB_ARCHIVO_COMPARACION where vcha_com_referencia = '" + txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux2.EOF
            rsaux3.Open "select * from vw_detalle_lista_precios where vcha_lis_lista_precios_id = '" + var_lista_precios + "' and vcha_art_articulo_id = '" + rsaux2!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If rsaux3.EOF Then
               var_posible = False
            End If
            rsaux3.Close
            rsaux2.MoveNext
         Wend
         rsaux2.Close
         If var_posible = True Then
            If rs.State = 1 Then
               rs.Close
            End If
            cnn.BeginTrans
            rsaux2.Open "select * from TB_ARCHIVO_COMPARACION where vcha_com_referencia = '" + txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               rsaux3.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emo_referencia = '" + txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux3.EOF Then
                  MsgBox "El movimiento ya fue impreso", vbOKOnly, "ATENCION"
               Else
                  Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
                  Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
                  If var_primera_vez = True Then
                     rsaux.Open "select * from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_moneda_local = IIf(IsNull(rsaux!inte_mon_moneda_local), 0, rsaux!inte_mon_moneda_local)
                     var_clave_moneda = IIf(IsNull(rsaux!vcha_mon_moneda_id), "", rsaux!vcha_mon_moneda_id)
                     If var_moneda_local = 0 Then
                        rsaux1.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'"
                        If rsaux1.EOF Then
                           MsgBox "No se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
                           GoTo no_tipo_cambio:
                        End If
                     Else
                        var_tipo_Cambio = 1
                     End If
                    
                     var_inserta = False
                     var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen, var_clave_movimiento, Now, var_numero_folio, 0, var_clave_cliente, "", var_almacen, "", "", var_clave_usuario_global, fun_NombrePc, 0, "", txt_archivo, "", "B", var_titular, var_agente, var_descuento_1, var_descuento_2, var_descuento_3, var_clave_moneda, var_tipo_Cambio)
                     var_numero_folio = var_numero_folio_regreso
                     txt_numero_folio = var_numero_folio
                     var_primera_vez = False
                  End If
                  While Not rsaux2.EOF
                     var_inserta = False
                     rsaux1.Open "select * from vw_detalle_lista_precios where vcha_lis_lista_precios_id = '" + var_lista_precios + "' and vcha_art_articulo_id = '" + rsaux2!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_precio = IIf(IsNull(rsaux1!floa_dli_precio), 0, rsaux1!floa_dli_precio)
                     rsaux1.Close
                     var_inserta = TB_SALIDAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen, var_clave_movimiento, var_numero_folio, rsaux2!vcha_Art_articulo_id, rsaux2!FLOA_COM_CANTIDAD_ENVIADA, rsaux2!FLOA_COM_COSTO, var_precio, "0")
                     rsaux2.MoveNext
                  Wend
                  rsaux1.Open "execute FACTURA_MERCANCIA_VISTAS '" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen + "', '" + var_clave_movimiento + "', " + Str(var_numero_folio) + ",'" + var_clave_usuario_global + "' ,'" + fun_NombrePc + "', '" + var_serie + "', 'FA'", cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux3.Close
               rsaux2.Close
            Else
               rsaux2.Close
            End If
            cnn.CommitTrans
         Else
            MsgBox "Existen articulos que no estan dentro de la lista de precios relacionada al cliente", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "No se a seleccionado ningun archivo", vbOKOnly, "ATENCION"
   End If
Else
   MsgBox "No existe una lista de precios relacionadas a este cliente", vbOKOnly, "ATENCION"
End If
Exit Sub
no_tipo_cambio:
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If rsaux1.State = 1 Then
      rsaux1.Close
   End If
End Sub

Private Sub cmd_nuevo_Click()
         txt_cliente = ""
         txt_de = ""
         txt_a = ""
         txt_renglones = ""
         var_primera_vez = True

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

Private Sub Form_Load()
   Dim var_contador_series As Integer
   var_cadena_seguridad = ""
   Left = 0
   Top = 500
   Set var_tabla = CreateObject("ADODB.connection")
   rs.Open "select * from tb_clientes where VCHA_TCL_TIPO_CLIENTE_ID  = 'I'", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_clientes.hwnd, rs, 1)
   rs.Close
   var_primera_vez = True
   rs.Open "select vcha_ser_serie_id from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_contador_serie = 0
      While Not rs.EOF
         var_contador_serie = var_contador_serie + 1
         rs.MoveNext
      Wend
      rs.MoveFirst
      txt_numero_folio.Enabled = True
      txt_archivo.Enabled = True
      Call RecsetToCombo(cmb_series.hwnd, rs, 0)
      If var_contador_serie > 1 Then
         cmb_series.Enabled = True
      Else
         cmb_series.Enabled = False
      End If
      rs.MoveFirst
      cmb_series = rs!VCHA_SER_SERIE_ID
      var_serie = rs!VCHA_SER_SERIE_ID
   Else
      MsgBox "No se a indicado una serie para esta Unidad organizacional", vbOKOnly, "ATENCION"
      txt_numero_folio.Enabled = False
      txt_archivo.Enabled = False
   End If
   rs.Close

End Sub

Private Sub salir_ButtonClick(ByVal Button As MSComctlLib.Button)
   Unload Me
End Sub

Private Sub Text1_Change()
  
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
End Sub



Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_factura_empresas)
End Sub

Private Sub txt_archivo_KeyPress(KeyAscii As Integer)
'On Error GoTo er:
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(txt_archivo) <> "" Then
         Dim var_ruta As String
         Dim var_costo As Double
         Dim var_contador As Double
         Dim var_contador_facturas As Double
         rs.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
         var_ruta = IIf(IsNull(rs!VCHA_PRI_RUTA_ALMACENES), "", rs!VCHA_PRI_RUTA_ALMACENES)
         var_numero_renglones = IIf(IsNull(rs!INTE_PRI_RENGLONES_FACTURA), "", rs!INTE_PRI_RENGLONES_FACTURA)
         rs.Close
         rs.Open "select * from tb_archivo_comparacion where vcha_com_referencia = '" + txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rs.Close
         Else
            rs.Close
            If var_tabla.State = 1 Then
               var_tabla.Close
            End If
            cnn.BeginTrans
            var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
            Cadena = "select * from " + txt_archivo + ""
            rs.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                  var_codigo = ""
                  var_costo = 0
                  rsaux3.Open "select * from tb_articulos where vcha_art_articulo_id = '" + Trim(rs!codigo) + "'", cnn, adOpenDynamic, adLockBatchOptimistic
                  If Not rsaux3.EOF Then
                     var_codigo = rsaux3!vcha_Art_articulo_id
                     var_costo = IIf(IsNull(rsaux3!mone_Art_costo_estandar), 0, rsaux3!mone_Art_costo_estandar)
                     rsaux3.Close
                  Else
                     rsaux3.Close
                     rsaux3.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + rs!codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux3.EOF Then
                        var_codigo = rsaux3!vcha_Art_articulo_id
                        rsaux3.Close
                        rsaux3.Open "select * from tb_articulos where vcha_art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockBatchOptimistic
                        var_costo = IIf(IsNull(rsaux3!mone_Art_costo_estandar), 0, rsaux3!mone_Art_costo_estandar)
                        rsaux3.Close
                     Else
                        rsaux3.Close
                        GoTo salir:
                     End If
                  End If
                  rsaux2.Open " INSERT INTO TB_ARCHIVO_COMPARACION  ( [VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID], [INTE_COM_NUMERO],[CHAR_COM_TIPO_PROVEEDOR],[VCHA_COM_PROVEEDOR], [VCHA_ART_ARTICULO_ID], [FLOA_COM_COSTO],[FLOA_COM_CANTIDAD_ENVIADA],[FLOA_COM_CANTIDAD_RECIBIDA],[VCHA_COM_TRANSPORTO],  [VCHA_COM_REFERENCIA],[DTIM_COM_FECHA]) values ( '" + var_empresa + "', '" + var_unidad_organizacional + "', '" + Trim(rs!cve_empres) + "', '" + var_clave_movimiento + "'," + Str(rs!NUMERO) + ", 'I', '" + rs!cve_empres + "','" + var_codigo + "', " + Str(var_costo) + "," + Str(rs!Cantidad) + ", 0,' ' ,'" + txt_archivo + "'," + CStr(Date) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
               Wend
            End If
            rs.Close
            cnn.CommitTrans
         End If
         var_descuento_1 = 0
         var_descuento_2 = 0
         var_descuento_3 = 0
         rs.Open "select inte_ser_factura from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
         var_numero_factura = IIf(IsNull(rs!inte_ser_factura), 0, rs!inte_ser_factura) + 1
         rs.Close
         rs.Open "select * from tb_archivo_comparacion where vcha_com_referencia = '" + txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
         rsaux2.Open "select * from tb_almacenes where vcha_alm_almacen_id ='" + rs!VCHA_ALM_ALMACEN_ID + "'", cnn, adOpenDynamic, adLockOptimistic
         txt_almacen = rsaux2!VCHA_ALM_NOMBRE
         var_almacen = rsaux2!VCHA_ALM_ALMACEN_ID
         rsaux2.Close
         If Not rs.EOF Then
            var_contador = 0
            var_contador_facturas = 1
            lv_piezas.ListItems.Clear
            While Not rs.EOF
               var_contador = var_contador + 1
               Dim list_item As ListItem
               Set list_item = lv_piezas.ListItems.Add(, , Trim(rs!vcha_Art_articulo_id))
               rsaux2.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rsaux2!vcha_art_nombre_español), "", Trim(rsaux2!vcha_art_nombre_español))
               rsaux2.Close
               list_item.SubItems(2) = IIf(IsNull(rs!FLOA_COM_CANTIDAD_ENVIADA), 0, rs!FLOA_COM_CANTIDAD_ENVIADA)
               rs.MoveNext
               If var_contador < var_numero_renglones And rs.EOF = True And var_contador_facturas <> 1 Then
                  var_contador_facturas = var_contador_facturas + 1
               End If
               If var_contador = var_numero_renglones Then
                  var_contador = 0
                  If Not rs.EOF Then
                     var_contador_facturas = var_contador_facturas + 1
                  End If
               End If
            Wend
            txt_de = var_numero_factura
            txt_a = var_numero_factura + var_contador_facturas - 1
            txt_renglones = var_contador_facturas
         Else
            MsgBox "El archivo no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
   End If
   Exit Sub
er:
   MsgBox "El archivo no existe", vbOKOnly, "ATENCIO"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   Exit Sub
salir:
   MsgBox "El archivo contiene artículos que no existen", vbOKOnly, "ATENCION"
   cnn.RollbackTrans
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   Exit Sub
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If

End Sub
