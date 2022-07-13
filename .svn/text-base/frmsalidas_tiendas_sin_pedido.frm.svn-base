VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmsalidas_tiendas_sin_pedido 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1155
      Picture         =   "frmsalidas_tiendas_sin_pedido.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancelar Movimiento"
      Top             =   615
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   165
      Picture         =   "frmsalidas_tiendas_sin_pedido.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   615
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   495
      Picture         =   "frmsalidas_tiendas_sin_pedido.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Buscar Movimiento"
      Top             =   615
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   825
      Picture         =   "frmsalidas_tiendas_sin_pedido.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   615
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7215
      Picture         =   "frmsalidas_tiendas_sin_pedido.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   630
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1230
      TabIndex        =   17
      Top             =   975
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   18
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
         TabIndex        =   19
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   1200
      TabIndex        =   14
      Top             =   1005
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         TabIndex        =   15
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
         TabIndex        =   16
         Top             =   120
         Width           =   3075
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   90
      TabIndex        =   12
      Top             =   495
      Width           =   7455
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   8430
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3750
      Width           =   1125
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   90
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
            Picture         =   "frmsalidas_tiendas_sin_pedido.frx":0A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_tiendas_sin_pedido.frx":131C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_tiendas_sin_pedido.frx":1BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_tiendas_sin_pedido.frx":2192
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_tiendas_sin_pedido.frx":2A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_tiendas_sin_pedido.frx":3348
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_tiendas_sin_pedido.frx":3C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_tiendas_sin_pedido.frx":3D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_tiendas_sin_pedido.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_tiendas_sin_pedido.frx":3F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_tiendas_sin_pedido.frx":406A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_tiendas_sin_pedido.frx":417C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   690
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
      Left            =   90
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Frame Frame3 
      Height          =   1485
      Index           =   1
      Left            =   135
      TabIndex        =   33
      Top             =   945
      Width           =   5580
      Begin VB.TextBox txt_almacen_origen 
         Height          =   315
         Left            =   825
         TabIndex        =   5
         Top             =   420
         Width           =   1395
      End
      Begin VB.TextBox txt_nombre_almacen_origen 
         Height          =   315
         Left            =   2235
         TabIndex        =   6
         Top             =   420
         Width           =   3270
      End
      Begin VB.TextBox txt_almacen_destino 
         Height          =   315
         Left            =   825
         TabIndex        =   7
         Top             =   765
         Width           =   1395
      End
      Begin VB.TextBox txt_nombre_almacen_destino 
         Height          =   315
         Left            =   2235
         TabIndex        =   8
         Top             =   765
         Width           =   3270
      End
      Begin VB.TextBox txt_descuento 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1065
         TabIndex        =   9
         Top             =   1110
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   37
         Top             =   825
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   36
         Top             =   480
         Width           =   660
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   35
         Top             =   120
         Width           =   5505
      End
      Begin VB.Label Label4 
         Caption         =   "Descuento:"
         Height          =   225
         Left            =   150
         TabIndex        =   34
         Top             =   1155
         Width           =   870
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1485
      Index           =   0
      Left            =   5760
      TabIndex        =   30
      Top             =   945
      Width           =   1800
      Begin VB.TextBox txt_folio 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
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
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   555
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   32
         Top             =   120
         Width           =   1725
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4860
      Left            =   135
      TabIndex        =   20
      Top             =   2430
      Width           =   7425
      Begin VB.TextBox txt_cantidad 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
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
         TabIndex        =   11
         Top             =   495
         Width           =   1890
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
         TabIndex        =   10
         Top             =   450
         Width           =   2640
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   2175
         TabIndex        =   21
         Top             =   2160
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   60
            TabIndex        =   22
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
            TabIndex        =   23
            Top             =   15
            Width           =   2895
         End
      End
      Begin MSComctlLib.ListView lv_traspasossalidas 
         Height          =   3270
         Left            =   45
         TabIndex        =   24
         Top             =   1035
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   5768
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
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   8264
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label lbl_cantidad_total 
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
         Height          =   390
         Left            =   5310
         TabIndex        =   28
         Top             =   4395
         Width           =   1965
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   4410
         TabIndex        =   27
         Top             =   615
         Width           =   675
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   26
         Top             =   120
         Width           =   7350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   615
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   90
      TabIndex        =   29
      Top             =   855
      Width           =   7455
   End
   Begin VB.Label lblnombremovimiento 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   38
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmsalidas_tiendas_sin_pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_kanban As String
Dim var_clave_agente As String
Dim var_clave_establecimiento As String
Dim var_clave_titular As String
Dim var_clave_cliente As String
Dim var_clave_ruta As String
Dim var_descuento_1 As Double
Dim var_descuento_2 As Double
Dim var_almacen_Destino As String
Dim var_almacen_origen As String
Dim var_primera_vez As Boolean
Dim var_numero_folio As Double
Dim var_cantidad_leida As Double
Dim var_costo As Double
Dim var_precio As Double
Dim var_descripcion_articulo As String
Dim var_estatus_movimiento As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_tipo_almacen As String
Dim var_correo_electronico As String
Dim var_tabla As ADODB.Connection
Dim var_ruta As String
Dim var_ventana As Integer
Dim var_clave_moneda As String
Dim var_tipo_lista As Integer
Dim var_año As Integer
Dim var_renglon As Double

Sub ilumina_grid()
   var_n = lv_traspasossalidas.ListItems.Count
   For var_i = 1 To var_n
       If var_i = var_renglon Then
          lv_traspasossalidas.ListItems.Item(var_i).Bold = True
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_traspasossalidas.ListItems.Item(var_i).ForeColor = &H8000&
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
       Else
          lv_traspasossalidas.ListItems.Item(var_i).Bold = False
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(1).Bold = False
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(2).Bold = False
          lv_traspasossalidas.ListItems.Item(var_i).ForeColor = &H80000012
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000012
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000012
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_traspasossalidas.ListItems.Item(var_renglon).Selected = True
      lv_traspasossalidas.selectedItem.EnsureVisible
   End If
   lv_traspasossalidas.Refresh
End Sub


Private Sub cmd_buscar_Click()
   frm_busqueda.Visible = True: var_ventana = 1
   txt_busqueda_folio.SetFocus
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
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_DET_ORDEN_SURTIDO_M = New TB_DET_ORDEN_SURTIDO_M
   Set TB_DET_EMBARQUE_I = New TB_DET_EMBARQUE_I
   Set TB_DETALLE_CAJAS_M = New TB_DETALLE_CAJAS_M
   
   Set TB_ENC_PEDIDOS_I = New TB_ENC_PEDIDOS_I
   Set TB_DETALLE_PEDIDOS_I = New TB_DETALLE_PEDIDOS_I
   Set TB_DETALLE_PEDIDOS_M = New TB_DETALLE_PEDIDOS_M
   
   
   Set TB_ENC_ORDEN_SURTIDO = New TB_ENC_ORDEN_SURTIDO
   Set TB_DET_ORDEN_SURTIDO_I = New TB_DET_ORDEN_SURTIDO_I
   
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
   Set TB_ENTRADAS_INSERTA = New TB_ENTRADAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_ARCH_COMPARACION_I = New TB_ARCH_COMPARACION_I
   Dim var_correo_electronico As String
   Dim var_canal_venta As String
   If var_numero_folio > 0 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         Set reporte = appl.OpenReport(App.Path + "\rep_nota_ENVIO_TIENDAS.rpt")
         reporte.RecordSelectionFormula = "{VW_ENVIO_TIENDAS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_ENVIO_TIENDAS.inte_emo_numero} = " + Str(var_numero_folio) + " and {VW_ENVIO_TIENDAS.FLOA_SAL_CANTIDAD} > 0 and {VW_ENVIO_TIENDAS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_ENVIO_TIENDAS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "'"
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Movimientos"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
         If var_tipo_almacen = "T" Then
            Call pro_envio_correo_app(var_correo_electronico, "Nota de Envio " & var_numero_folio, "Se anexa nota de envio", App.Path & "\dev_tien.dbf")
         End If
         If var_tabla.State = 1 Then
            var_tabla.Close
         End If
     Else
         var_posible_Cantidad = 1
         If var_empresa = "18" Or var_empresa = "31" Then
            Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and floa_Sal_cantidad > 0"
            rsaux10.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux10.EOF
                  rsaux9.Open "select * from tb_existencias where vcha_Alm_almacen_id = '" + var_almacen_Destino + "' and vcha_Art_Articulo_id = '" + rsaux10!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux9.EOF Then
                     var_cantidad = IIf(IsNull(rsaux9!floa_Exi_Cantidad_disponible), 0, rsaux9!floa_Exi_Cantidad_disponible)
                     If var_empresa = "18" Then
                        If rsaux10!vcha_Art_Articulo_id = "360010000002" Or rsaux10!vcha_Art_Articulo_id = "360020000009" Or rsaux10!vcha_Art_Articulo_id = "900000000003" Or rsaux10!vcha_Art_Articulo_id = "911110000005" Or rsaux10!vcha_Art_Articulo_id = "801020000007" Or rsaux10!vcha_Art_Articulo_id = "801010000000" Or rsaux10!vcha_Art_Articulo_id = "801040000001" Or rsaux10!vcha_Art_Articulo_id = "801050000008" Or rsaux10!vcha_Art_Articulo_id = "790020020002" Then
                           var_cantidad = Round(IIf(IsNull(rsaux10!floa_Sal_Cantidad), 0, rsaux10!floa_Sal_Cantidad), 4) + 1
                        End If
                     End If
                     
                     If Round(var_cantidad, 4) < Round(IIf(IsNull(rsaux10!floa_Sal_Cantidad), 0, rsaux10!floa_Sal_Cantidad), 4) Then
                        var_posible_Cantidad = 0
                        If var_cadena_articulos = "" Then
                           rsaux8.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux10!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux8.EOF Then
                              var_nombre_articulo = IIf(IsNull(rsaux8!vcha_Art_nombre_español), "", rsaux8!vcha_Art_nombre_español)
                           Else
                              var_nombre_articulo = ""
                           End If
                           rsaux8.Close
                           var_cadena_articulos = rsaux10!vcha_Art_Articulo_id + " " + var_nombre_articulo + " Existen [" + CStr(var_cantidad) + "] y salen [" + CStr(rsaux10!floa_Sal_Cantidad) + "]"
                        Else
                           rsaux8.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux10!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux8.EOF Then
                              var_nombre_articulo = IIf(IsNull(rsaux8!vcha_Art_nombre_español), "", rsaux8!vcha_Art_nombre_español)
                           Else
                              var_nombre_articulo = ""
                           End If
                           rsaux8.Close
                           var_cadena_articulos = var_cadena_articulos + ", " + rsaux10!vcha_Art_Articulo_id + " " + var_nombre_articulo + " Existen [" + CStr(var_cantidad) + "] y salen [" + CStr(rsaux10!floa_Sal_Cantidad) + "]"
                        End If
                     
                     
                     End If
                  Else
                     If var_cadena_articulos = "" Then
                        rsaux8.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux10!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux8.EOF Then
                           var_nombre_articulo = IIf(IsNull(rsaux8!vcha_Art_nombre_español), "", rsaux8!vcha_Art_nombre_español)
                        Else
                           var_nombre_articulo = ""
                        End If
                        rsaux8.Close
                        var_cadena_articulos = rsaux10!vcha_Art_Articulo_id + " " + var_nombre_articulo
                     Else
                        rsaux8.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux10!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux8.EOF Then
                           var_nombre_articulo = IIf(IsNull(rsaux8!vcha_Art_nombre_español), "", rsaux8!vcha_Art_nombre_español)
                        Else
                           var_nombre_articulo = ""
                        End If
                        rsaux8.Close
                        var_cadena_articulos = var_cadena_articulos + ", " + rsaux10!vcha_Art_Articulo_id + " " + var_nombre_articulo
                     End If
                     var_posible_Cantidad = 0
                  End If
                  rsaux9.Close
                  rsaux10.MoveNext
            Wend
            rsaux10.Close
         End If
         If var_posible_Cantidad = 1 Then
            var_si = MsgBox("¿Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
            If var_si = 1 Then
               cnn.BeginTrans
               var_posible_cerrar_KANBAN = True
               If var_posible_kanban = 1 Then
                  Set TB_PROC_KANBANS_EN_MOVIMIENTO = New TB_PROC_KANBANS_EN_MOVIMIENTO
                  var_inserta = TB_PROC_KANBANS_EN_MOVIMIENTO.Anadir(Me.txt_almacen_origen, var_clave_movimiento, CDbl(Me.txt_folio), "", "")
                  If var_kanban_exito = "N" Then
                     var_posible_cerrar_KANBAN = False
                  End If
               Else
                  var_posible_cerrar_KANBAN = True
               End If
               If var_posible_cerrar_KANBAN = True Then
                  rs.Open "select * from VW_clientes where vcha_cli_clave_id = '" + Me.txt_almacen_destino + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     txt_agente = rs!VCHA_AGE_AGENTE_ID
                     var_descuento_1 = IIf(IsNull(rs!floa_gac_Descuento_1), 0, rs!floa_gac_Descuento_1)
                     var_descuento_2 = IIf(IsNull(rs!FLOA_GAC_DESCUENTO_2), 0, rs!FLOA_GAC_DESCUENTO_2)
                     var_dias_condiciones = IIf(IsNull(rs!inte_pla_dias), 0, rs!inte_pla_dias)
                     var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                     txt_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
                  End If
                  rs.Close
                  var_maximo_orden = 0
                  ok = TB_ENC_PEDIDOS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, "T", maximo_pedido, 0, Date, Date, CStr(txt_agente), CStr(txt_titular), Me.txt_almacen_destino, "", 0, 0, "", var_descuento_1, var_descuento_2, 0, CDbl(var_dias_condiciones), 0, var_clave_usuario_global, fun_NombrePc, Date, var_clave_moneda, 0)
                  rsaux.Open "select * from vw_maximo_orden_surtido ", cnn, adOpenDynamic, adLockOptimistic
                  If rsaux.EOF Then
                     var_maximo_orden = 1
                  Else
                     If IsNull(rsaux!maximo) Then
                        var_maximo_orden = 1
                     Else
                        var_maximo_orden = rsaux!maximo + 1
                     End If
                  End If
                  rsaux.Close
                  rsaux5.Open "select * from tb_temporal_salidas with (nolock) where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux5.EOF
                        rsaux4.Open "update tb_existencias set floa_exi_temporal_cantidad_salida = isnull(floa_exi_temporal_cantidad_salida,0) - " + CStr(rsaux5!floa_Sal_Cantidad) + " where vcha_alm_almacen_id = '" + rsaux5!VCHA_ALM_ALMACEN_ID + "' and vcha_Art_Articulo_id = '" + rsaux5!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        rsaux5.MoveNext
                  Wend
                  ok = TB_ENC_ORDEN_SURTIDO.Anadir(var_empresa, var_unidad_organizacional, "T", maximo_pedido, Me.txt_almacen_origen, CDbl(var_maximo_orden), Date, Date + 0, "", CStr(txt_titular), Me.txt_almacen_destino, "", var_descuento_1, var_descuento_2, 0, "", "", Date, 0, var_clave_moneda, Date)
                  rs.Open "update tb_encabezado_movimientos set inte_emo_numero_origen = " + CStr(var_maximo_orden) + " where inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  cnn.CommandTimeout = 360
                  rsaux4.Open "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", 1,'','','" + var_clave_titular + "','" + Me.txt_almacen_destino + "',0,0,0", cnn, adOpenDynamic, adLockOptimistic
          
                  rs.Open "select * from vw_maximo_embarque where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                  If rs.EOF Then
                     var_numero_embarque = 1
                  Else
                     var_numero_embarque = rs!maximo_embarque + 1
                  End If
                  rs.Close
                
                  Set TB_ENC_EMBARQUE_I = New TB_ENC_EMBARQUE_I
                  ok = False
                  rs.Open "insert into tb_encabezado_embarques (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, INTE_EMB_EMBARQUE, INTE_JAU_JAULA_ID, VCHA_VEH_VEHICULO_ID, VCHA_AGE_AGENTE_ID, DTIM_EMB_FECHA_INICIO, DTIM_EMB_FECHA_FINAL, CHAR_EMB_ESTATUS, VCHA_CHO_CHOFER_ID, FLOA_EMB_CUBICAJE, CHAR_EMB_TIPO, INTE_EMB_BLOQUEADO, VCHA_EMB_BLOQUEADO_POR, VCHA_AUD_MAQUINA, VCHA_AUD_USUARIO) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', " + CStr(var_numero_embarque) + ", 0, '', '" + txt_agente + "', getdate(), getdate(), 'I', '', 0,'',0, '','" + fun_NombrePc + "','" + var_clave_usuario_global + "')", cnn, adOpenDynamic, adLockOptimistic
                  var_inserta = TB_DET_EMBARQUE_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_numero_embarque, var_clave_movimiento, var_numero_folio, "")
                  txt_numero_embarque = CStr(var_numero_embarque)
                  var_estatus_embarque = "I"
                  cnn.CommitTrans
                  
                  Set reporte = appl.OpenReport(App.Path + "\rep_nota_ENVIO_TIENDAS.rpt")
                  reporte.RecordSelectionFormula = "{VW_ENVIO_TIENDAS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_ENVIO_TIENDAS.inte_emo_numero} = " + Str(var_numero_folio) + " and {VW_ENVIO_TIENDAS.FLOA_SAL_CANTIDAD} > 0 and {VW_ENVIO_TIENDAS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_ENVIO_TIENDAS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "'"
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Movimientos"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  var_estatus_movimiento = "I"
                  txt_codigo.Enabled = False
                  txt_foco.Enabled = False
               Else
                  cnn.RollbackTrans
                  MsgBox "No se pudo cerrar el movimiento kanban", vbOKOnly, "ATENCION"
               End If
            End If
         Else
            MsgBox "El movimiento no se puede imprimir ya que las existencias de los siguientes artículos exceden a la cantidad disponible en el almacen " + var_cadena_articulos
         End If
      End If
   Else
      MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   If var_numero_folio > 0 Then
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 where inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   txt_codigo.Enabled = False
   var_primera_vez = True
   frm_busqueda.Visible = False
   var_ventana = 0
   lv_traspasossalidas.ListItems.Clear
   var_numero_folio = 0
   txt_folio = ""
   txt_codigo = ""
   var_estatus_movimiento = ""
   txt_almacen_origen = ""
   txt_almacen_destino = ""
   txt_nombre_almacen_origen = ""
   txt_nombre_almacen_destino = ""
   txt_almacen_origen.Enabled = True
   txt_almacen_origen.SetFocus
   lbl_cantidad_total = "0"
End Sub

Private Sub Command1_Click()
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
   If Shift = 4 And KeyCode = 67 Then
   End If
End Sub

Private Sub Form_Load()
   var_posible_kanban = 0
   lbl_cantidad_total = "0"
   var_cadena_seguridad = ""
   Top = 0
   Left = 2000
   frm_lista.Visible = False
   Set var_tabla = CreateObject("ADODB.connection")
   rs.Open "select VCHA_PRI_RUTA_NOTAS_ENVIO from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_ruta = IIf(IsNull(rs!VCHA_PRI_RUTA_NOTAS_ENVIO), "", rs!VCHA_PRI_RUTA_NOTAS_ENVIO)
   End If
   rs.Close
   rs.Open "select * from tb_monedas where inte_mon_moneda_local = 1", cnn, adOpenDynamic
   var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
   rs.Close
   var_estatus_movimiento = ""
   var_ventana = 0
   frm_busqueda.Visible = False
   frm_eliminar.Visible = False
   lbl_cantidad.Visible = False
   txt_cantidad.Visible = False
   txt_codigo.Enabled = False
   var_primera_vez = True
   var_cantidad_leida = 1#
   txt_almacen_destino.Enabled = False
   txt_almacen_origen.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_numero_folio > 0 Then
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 where inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   'Call activa_forma(var_activa_forma_salidas_proveedor)
   Call activa_forma(var_activa_forma_salidas_sin_comparacion)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 0 Then
         If var_tipo_lista = 1 Then
            txt_almacen_origen = lv_lista.selectedItem
            txt_nombre_almacen_origen = lv_lista.selectedItem.SubItems(1)
            txt_almacen_origen.SetFocus
         End If
         If var_tipo_lista = 2 Then
            txt_almacen_destino = lv_lista.selectedItem
            txt_nombre_almacen_destino = lv_lista.selectedItem.SubItems(1)
            txt_almacen_destino.SetFocus
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

Private Sub lv_traspasossalidas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         MsgBox "Imporsible modificar el movimiento", vbOKOnly, "ATENCION"
      Else
         var_ventana = 1
         frm_eliminar.Visible = True
         txt_cantidad_eliminar.SetFocus
      End If
   End If
End Sub

Private Sub Toolbar1_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
End Sub

Private Sub txt_almacen_destino_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_almacen_destino_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "SELECT * FROM VW_CLIENTES WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and vcha_tcl_tipo_cliente_id = 'T'", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
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
End Sub

Private Sub txt_almacen_destino_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_nombre_almacen_destino.SetFocus
   End If
End Sub

Private Sub txt_almacen_destino_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_almacen_destino) <> "" Then
      rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + Me.txt_almacen_destino + "' and vcha_emp_empresa_id  = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_descuento = IIf(IsNull(rs!floa_gac_Descuento_1), 0, rs!floa_gac_Descuento_1)
         txt_almacen_destino = rs!vcha_cli_clave_id
         txt_nombre_almacen_destino = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         var_almacen_Destino = ""
         txt_almacen_destino.Enabled = False
         var_tipo_almacen = ""
         var_correo_electronico = ""
         txt_codigo.Enabled = True
      Else
         txt_codigo.Enabled = False
         txt_almacen_destino = ""
         txt_nombre_almacen_destino = ""
         MsgBox "Clave de almacen incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      Me.txt_descuento = ""
      txt_almacen_destino = ""
      txt_nombre_almacen_destino = ""
   End If
End Sub

Private Sub txt_almacen_origen_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_almacen_origen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from vw_movimientos_almacenes WHERE VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and char_alm_tipo = 'A' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Almacenes"
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
End Sub

Private Sub txt_almacen_origen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_nombre_almacen_origen.SetFocus
   End If
End Sub

Private Sub txt_almacen_origen_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_almacen_origen) <> "" Then
      rs.Open "select * from vw_movimientos_almacenes  where vcha_alm_almacen_id = '" + txt_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and char_alm_tipo = 'A'", cnn, adOpenDynamic, adLockBatchOptimistic
      If Not rs.EOF Then
         var_almacen_origen = rs!VCHA_ALM_ALMACEN_ID
         txt_nombre_almacen_origen = rs!VCHA_ALM_NOMBRE
         txt_almacen_destino.Enabled = True
         txt_almacen_origen.Enabled = False
      Else
         var_almacen_origen = ""
         txt_almacen_origen = ""
         txt_nombre_almacen_origen = ""
         MsgBox "Clave de almacen incorrecto", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
        KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_busqueda_folio) <> "" Then
         If var_numero_folio = CDbl(txt_busqueda_folio) Then
            rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 where inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
         rs.Open "select * from tb_encabezado_movimientos where inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If var_numero_folio > 0 Then
               rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 where inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            var_movimiento_bloqueado = IIf(IsNull(rs!INTE_EMO_BLOQUEADO), 0, rs!INTE_EMO_BLOQUEADO)
            If var_movimiento_bloqueado = 0 Then
               var_almacen_destino_tem = rs!VCHA_EMO_ALMACEN_DESTINO
               var_almacen_origen_tem = rs!VCHA_ALM_ALMACEN_ID
               var_posible = 1
               If var_tipo_permiso = 1 Then
                  rsaux.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_per_almacen_1 = '" + var_almacen_origen_tem + "'", cnn, adOpenDynamic, adLockBatchOptimistic
                  If rsaux.EOF Then
                     var_posible = 0
                  End If
                  rsaux.Close
               End If
               If var_posible = 1 Then
                  lbl_cantidad_total = "0"
                  var_estatus_movimiento = rs!char_Emo_estatus
                  var_almacen_Destino = rs!VCHA_EMO_ALMACEN_DESTINO
                  Me.txt_descuento = IIf(IsNull(rs!floa_emo_descuento_1), 0, rs!floa_emo_descuento_1)
                  var_almacen_origen = rs!VCHA_ALM_ALMACEN_ID
                  lv_traspasossalidas.ListItems.Clear
                  var_primera_vez = False
                  var_numero_folio = rs!INTE_EMO_NUMERO
                  txt_folio = var_numero_folio
                  rsaux.Open "select * from vw_clientes where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_almacen_destino = rsaux!vcha_cli_clave_id
                  txt_nombre_almacen_destino = rsaux!VCHA_CLI_NOMBRE
                  rsaux.Close
                  rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_almacen_origen = rsaux(2).Value
                  txt_nombre_almacen_origen = rsaux(3).Value
                  rsaux.Close
                  var_tipo_almacen = ""
                  var_correo_electronico = ""
                  rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 1 where inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                  If rsaux.State = 1 Then
                     rsaux.Close
                  End If
                  rsaux.Open "select * from tb_temporal_salidas with (nolock) where inte_sal_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     While Not rsaux.EOF
                        rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           Set list_item = lv_traspasossalidas.ListItems.Add(, , rsaux!vcha_Art_Articulo_id)
                           list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
                           list_item.SubItems(2) = IIf(IsNull(rsaux!floa_Sal_Cantidad), 0, rsaux!floa_Sal_Cantidad)
                           lbl_cantidad_total = CStr(CDbl(lbl_cantidad_total) + IIf(IsNull(rsaux!floa_Sal_Cantidad), 0, rsaux!floa_Sal_Cantidad))
                           rsaux2.Close
                           rsaux.MoveNext:
                        End If
                     Wend
                  End If
                  rsaux.Close
                  If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
                     txt_codigo.Enabled = False
                     txt_cantidad.Visible = False
                     lbl_cantidad.Visible = False
                     txt_foco.Enabled = False
                  Else
                     txt_foco.Enabled = False
                     txt_codigo.Enabled = True
                     txt_cantidad.Visible = False
                     lbl_cantidad.Visible = False
                  End If
               Else
                  MsgBox "No esta autorizado para modificar este movimiento", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El movimiento esta siendo usado por otro usuario", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El número de movimiento no existe ", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
      var_ventana = 0
      frm_busqueda.Visible = False
   End If
   If KeyAscii = 27 Then
      var_ventana = 0
      frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   'Select Case KeyAscii
   'Case 48 To 57, 52, 13, 8, 46, 27
   'Case Else
   '     KeyAscii = 0
   'End Select
   If KeyAscii = 13 Then
      If var_posible_kanban = 1 Then
         If IsNumeric(Me.txt_cantidad_eliminar) Then
            Set TB_CANCELAR_RES_FUERA_DE_KANBAN = New TB_CANCELAR_RES_FUERA_DE_KANBAN
            var_inserta = TB_CANCELAR_RES_FUERA_DE_KANBAN.Anadir(Me.txt_almacen_origen, var_clave_movimiento, var_numero_folio, Me.lv_traspasossalidas.selectedItem, CDbl(Me.txt_cantidad_eliminar), "", "")
            var_kanban_es_un_kanban = var_kanban_es_un_kanban
            var_kanban_almacen_id = var_kanban_almacen_id
            var_kanban_articulo_id = var_kanban_articulo_id
            var_kanban_exito = var_kanban_exito
            var_kanban_mensaje = var_kanban_mensaje
            If var_kanban_exito = "S" Then
               var_posible = True
            Else
               frmmensaje.lbl_mensaje = var_kanban_mensaje
               frmmensaje.Show 1
               var_posible = False
            End If
         Else
            Set TB_ES_UN_KANBAN = New TB_ES_UN_KANBAN
            var_kanban = Me.txt_codigo
            var_inserta = TB_ES_UN_KANBAN.Anadir(Me.txt_cantidad_eliminar, "", "", "", "", "")
            var_kanban_es_un_kanban = var_kanban_es_un_kanban
            var_kanban_almacen_id = var_kanban_almacen_id
            var_kanban_articulo_id = var_kanban_articulo_id
            var_kanban_exito = var_kanban_exito
            var_kanban_mensaje = var_kanban_mensaje
            If var_kanban_es_un_kanban = "S" Then
               If Me.lv_traspasossalidas.selectedItem = var_kanban_articulo_id Then
                  Set TB_CANCELAR_RESERVACION_KANBAN = New TB_CANCELAR_RESERVACION_KANBAN
                  var_kanban = Me.txt_codigo
                  var_inserta = TB_CANCELAR_RESERVACION_KANBAN.Anadir(Me.txt_almacen_origen, var_clave_movimiento, var_numero_folio, Me.txt_cantidad_eliminar, "", "")
                  var_kanban_es_un_kanban = var_kanban_es_un_kanban
                  var_kanban_almacen_id = var_kanban_almacen_id
                  var_kanban_articulo_id = var_kanban_articulo_id
                  var_kanban_exito = var_kanban_exito
                  var_kanban_mensaje = var_kanban_mensaje
                  If var_kanban_exito = "S" Then
                     var_posible = True
                  Else
                     frmmensaje.lbl_mensaje = var_kanban_mensaje
                     frmmensaje.Show 1
                     var_posible = False
                  End If
               Else
                  frmmensaje.lbl_mensaje = "El codigo de kanban no corresponde al del artículo seleccionado"
                  frmmensaje.Show 1
                  var_posible = False
               End If
            Else
               frmmensaje.lbl_mensaje = var_kanban_mensaje
               frmmensaje.Show 1
               var_posible = False
            End If
         End If
      Else
         var_posible = True
      End If
         
      If var_posible = True Then
         If var_posible_kanban = 1 Then
            If Not IsNumeric(txt_cantidad_eliminar) Then
               Me.txt_cantidad_eliminar = 1
            End If
         End If
      
         
         If IsNumeric(txt_cantidad_eliminar) Then
            If CDbl(txt_cantidad_eliminar) <= lv_traspasossalidas.selectedItem.SubItems(2) * 1 Then
               Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
               Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
               var_cantidad_eliminar = Val(txt_cantidad_eliminar)
               rs.Open "update tb_existencias set floa_Exi_temporal_cantidad_salida = isnull(floa_Exi_temporal_cantidad_salida,0) - " + CStr(var_cantidad_eliminar) + " where vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_Art_articulo_id = '" + lv_traspasossalidas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
               var_inserta = False
               var_inserta = TB_TEMPORAL_SALIDAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, lv_traspasossalidas.selectedItem, 0 - Val(txt_cantidad_eliminar))
               lbl_cantidad_total = CStr(CDbl(lbl_cantidad_total) - var_cantidad_eliminar)
               lv_traspasossalidas.selectedItem.SubItems(2) = lv_traspasossalidas.selectedItem.SubItems(2) - Val(txt_cantidad_eliminar)
               var_renglon = lv_traspasossalidas.selectedItem.Index
               Call ilumina_grid
               var_ventana = 0
               frm_eliminar.Visible = False
               txt_codigo.SetFocus
            Else
               MsgBox "La cantidad no puede superar a la leida", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Cantidad Incorrecta", vbOKOnly, "ATENCION"
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
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
        KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_cantidad) <> "" Then
         If Not IsNumeric(txt_cantidad) Then
            txt_cantidad = "0"
         End If
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

Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_codigo_seleccionado = ""
      frmbusqueda_articulo.Show 1
      Me.txt_codigo = var_codigo_seleccionado
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   Dim var_recontable As Integer
   Dim var_caja As String
   Dim var_cantidad_caja As Integer
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   txt_codigo = Trim(txt_codigo)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If var_posible_kanban = 1 Then
         Set TB_ES_UN_KANBAN = New TB_ES_UN_KANBAN
         var_kanban = Me.txt_codigo
         var_inserta = TB_ES_UN_KANBAN.Anadir(Me.txt_codigo, "", "", "", "", "")
         var_kanban_es_un_kanban = var_kanban_es_un_kanban
         var_kanban_almacen_id = var_kanban_almacen_id
         var_kanban_articulo_id = var_kanban_articulo_id
         var_kanban_exito = var_kanban_exito
         var_kanban_mensaje = var_kanban_mensaje
         
         If var_kanban_es_un_kanban = "S" Then
            Me.txt_codigo = var_kanban_articulo_id
         Else
            var_kanban_almacen_id = Me.txt_almacen_origen
         End If
         If var_kanban_almacen_id = Me.txt_almacen_origen Then
            If var_empresa = 16 Then
               If Len(Me.txt_codigo) = 6 Then
                  Me.txt_codigo = Mid(Me.txt_codigo, 1, 3) + "-" + Mid(Me.txt_codigo, 4, 3) + "-"
               Else
                  If Len(Me.txt_codigo) = 7 Then
                     Me.txt_codigo = Mid(Me.txt_codigo, 1, 3) + "-" + Mid(Me.txt_codigo, 4, 3) + "-" + Mid(Me.txt_codigo, 7, 1)
                  End If
               End If
            End If
            
            var_verificador = True
            If Len(Trim(txt_codigo)) = 12 Then
               Call calcula_verificador(Trim(txt_codigo))
            End If
            If var_verificador = True Then
               var_es_caja = False
               If Trim(txt_codigo) <> "" Then
                  If Left(Trim(txt_codigo), 1) = "C" Then
                     x = Mid(txt_codigo, 2, 6)
                     var_embarque_caja = 0
                     If IsNumeric(x) Then
                        var_embarque_caja = CDbl(x)
                        If var_embarque_caja = var_numero_embarque Then
                           var_es_caja = True
                        Else
                           frmmensaje.lbl_mensaje = "La caja pertenece a otro embarque"
                           frmmensaje.Show 1
                           'MsgBox "La caja pertenece al embarque " + CStr(var_embarque_caja)
                           var_es_caja = False
                        End If
                     Else
                        frmmensaje.lbl_mensaje = "Caja incorrecta"
                        frmmensaje.Show 1
                        'MsgBox "Caja incorrecta", vbOKOnly, "ATENCION"
                        var_es_caja = False
                     End If
                  Else
                     var_es_caja = False
                  End If
                  If var_es_caja = True Then
                     txt_foco.Enabled = True
                     txt_foco.SetFocus
                  Else
                     var_caja = Left(txt_codigo, 6)
                     If var_caja = "000005" Or var_caja = "000010" Or var_caja = "000015" Or var_caja = "000020" Or var_caja = "000001" Or var_caja = "000002" Or var_caja = "000003" Or var_caja = "000004" Or var_caja = "000006" Or var_caja = "000007" Or var_caja = "000008" Or var_caja = "000009" Or var_caja = "000010" Or var_caja = "000011" Or var_caja = "000012" Or var_caja = "000013" Or var_caja = "000014" Or var_caja = "000015" Or var_caja = "000016" Or var_caja = "000017" Or var_caja = "000018" Or var_caja = "000019" Or var_caja = "000020" Or var_caja = "000021" Or var_caja = "000022" Or var_caja = "000023" Or var_caja = "000024" Or var_caja = "000025" Or var_caja = "000026" Or var_caja = "000027" Or var_caja = "000028" Or var_caja = "000029" Or var_caja = "000030" Then
                        var_cantidad_caja = CInt(var_caja)
                        txt_codigo = Mid(txt_codigo, 7, 5)
                     End If
                     rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_descripcion_articulo = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
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
                              var_descripcion_articulo = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
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
                              frmmensaje.lbl_mensaje = "El artículo no existe"
                              frmmensaje.Show 1
                              'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                              txt_codigo = ""
                           End If
                        Else
                           frmmensaje.lbl_mensaje = "El artículo no existe"
                           frmmensaje.Show 1
                          'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                           txt_codigo = ""
                           rs.Close
                        End If
                     End If
                  End If
               End If
            Else
               txt_codigo = ""
               frmmensaje.lbl_mensaje = "Error en Código"
               frmmensaje.Show 1
               ' MsgBox "Error en Código", vbOKOnly, "ATENCION"
            End If
         Else
            txt_codigo = ""
            frmmensaje.lbl_mensaje = "El almacén del Kanban no pertenece al almacén del movimiento"
            frmmensaje.Show 1
         End If
      Else
''' FIN KANBAN
         var_verificador = True
         If Len(Trim(txt_codigo)) = 12 Then
            Call calcula_verificador(Trim(txt_codigo))
         End If
         If var_verificador = True Then
            var_caja = Left(txt_codigo, 6)
            If var_caja = "000005" Or var_caja = "000010" Or var_caja = "000015" Or var_caja = "000020" Then
               var_cantidad_caja = CInt(var_caja)
               txt_codigo = Mid(txt_codigo, 7, 5)
            End If
            var_costo = 0
            var_precio = 0
            If Trim(txt_codigo) <> "" Then
               rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  If IsNull(rs(43).Value) Then
                     var_recontable = 0
                  Else
                     var_recontable = rs(43).Value
                  End If
                  'var_recontable = 1
                  var_descripcion_articulo = rs(1).Value
                  var_costo = rs(3).Value
                  var_precio = rs(2).Value
                  rs.Close
                  rs.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_costo = IIf(IsNull(rs(4).Value), 0, rs(4).Value)
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
                           var_recontable = 1
                        Else
                           var_recontable = 0
                        End If
                     
                        var_descripcion_articulo = rs(1).Value
                        var_costo = IIf(IsNull(rs(3).Value), 0, rs(3).Value)
                        var_precio = rs(2).Value
                        rs.Close
                        rs.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_costo = rs(4).Value
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
                        frmmensaje.Show 1
                        'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                     End If
                  Else
                     txt_codigo = ""
                     frmmensaje.lbl_mensaje = "El artículo no existe"
                     frmmensaje.Show 1
                     'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                     rs.Close
                  End If
               End If
            Else
            End If
         Else
            txt_codigo = ""
            frmmensaje.lbl_mensaje = "Error en Código"
            frmmensaje.Show 1
            'MsgBox "Error en Código", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   Dim var_inserta As Boolean
   var_año = 2005
   If Trim(txt_codigo.Text) <> "" Then
      var_pase_existencias = 1
      If var_empresa = "18" Or var_empresa = "31" Then
         If var_numero_folio = 0 Or Trim(Me.txt_folio) = "" Then
            var_cantidad_temporal = 0
         Else
            rsaux.Open "select isnull(floa_sal_cantidad,0) from tb_Temporal_salidas where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + CStr(var_numero_folio) + " and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_cantidad_temporal = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
            Else
               var_cantidad_temporal = 0
            End If
            rsaux.Close
         End If
         'MsgBox CStr(var_cantidad_temporal)
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "select floa_exi_Cantidad_disponible from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_cantidad_Existencias = IIf(IsNull(rsaux!floa_Exi_Cantidad_disponible), 0, rsaux!floa_Exi_Cantidad_disponible)
         Else
            var_cantidad_Existencias = 0
         End If
         rsaux.Close
         var_cantidad_posible = var_cantidad_Existencias - (var_cantidad_temporal + var_cantidad_leida)
         If var_cantidad_posible < 0 Then
            var_pase_existencias = 0
         End If
      End If
      If var_empresa = "18" Then
         If Me.txt_codigo = "360010000002" Or Me.txt_codigo = "360020000009" Or Me.txt_codigo = "900000000003" Or Me.txt_codigo = "911110000005" Or Me.txt_codigo = "801020000007" Or Me.txt_codigo = "801010000000" Or Me.txt_codigo = "801040000001" Or Me.txt_codigo = "801050000008" Or Me.txt_codigo = "790020020002" Then
            var_pase_existencias = 1
         End If
      End If
      If var_pase_existencias = 1 Then
         If rsaux5.State = 1 Then
            rsaux5.Close
         End If
         rsaux5.Open "select isnull(floa_Exi_cantidad_disponible,0) - isnull(floa_Exi_Temporal_cantidad_salida,0) from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux5.EOF Then
            'var_cantidad_disponible = IIf(IsNull(rsaux5(0).Value), 0, rsaux5(0).Value)
            'If var_cantidad_disponible >= var_cantidad_leida Then
               bandera_suma = False
               If var_primera_vez = True Then
                  var_primera_vez = False
                  rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + Me.txt_almacen_destino + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                     var_clave_titular = rs!vcha_tit_titular_id
                     var_canal_venta = rs!vcha_can_canal_venta_id
                     var_descuento_1 = IIf(IsNull(rs!floa_gac_Descuento_1), 0, rs!floa_gac_Descuento_1)
                     var_descuento_2 = IIf(IsNull(rs!FLOA_GAC_DESCUENTO_2), 0, rs!FLOA_GAC_DESCUENTO_2)
                     txt_agente = rs!VCHA_AGE_AGENTE_ID
                  End If
                  rs.Close
                  rs.Open "select * from tb_Detalle_Establecimientos where vcha_cli_clave_id = '" + txt_almacen_destino + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_clave_establecimiento = ""
                  If Not rs.EOF Then
                     var_clave_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
                  End If
                  rs.Close
                  var_numero_folio = 0
                  var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, CDbl(var_numero_folio), 0, CStr(txt_almacen_destino), "", var_almacen_origen, "", "", var_clave_usuario_global, fun_NombrePc, 0, "", "", var_clave_establecimiento, "", var_clave_titular, CStr(txt_agente), var_descuento_1, var_descuento_2, 0, var_clave_moneda, 0)
                  var_numero_folio = var_numero_folio_regreso
                  Me.txt_folio = var_numero_folio
                  rsaux.Open "update tb_encabezado_movimientos set VCHA_CAN_CANAL_VENTA_ID = '" + var_canal_venta + "' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
               End If
         
         
               If var_posible_kanban = 1 Then
                  Set TB_RESERVAR_FUERA_DE_KANBAN = New TB_RESERVAR_FUERA_DE_KANBAN
                  Set TB_RESERVAR_KANBAN = New TB_RESERVAR_KANBAN
                  If var_kanban_es_un_kanban = "S" Then
                     var_inserta = TB_RESERVAR_KANBAN.Anadir(var_kanban, var_clave_movimiento, var_numero_folio, Me.txt_almacen_origen, Me.txt_codigo, "", "")
                     If var_kanban_exito = "S" Then
                        var_posible_leido = 1
                     Else
                        var_posible_leido = 0
                     End If
                  Else
                     var_inserta = TB_RESERVAR_FUERA_DE_KANBAN.Anadir(var_numero_folio, var_clave_movimiento, Me.txt_almacen_origen, Me.txt_codigo, "", "")
                     If var_kanban_exito = "S" Then
                        var_posible_leido = 1
                     Else
                        var_posible_leido = 0
                     End If
                  End If
               Else
                  var_kanban_mensaje = ""
                  var_posible_leido = 1
               End If
               If var_posible_leido = 1 Then
                  rsaux4.Open "select floa_dli_precio from tb_detalle_lista_precios where vcha_Art_Articulo_id = '" + txt_codigo + "' and vcha_lis_lista_precios_id = '01'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux4.EOF Then
                     var_precio = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                  End If
                  rsaux4.Close
                  rs.Open "update tb_existencias set floa_Exi_temporal_cantidad_salida = isnull(floa_Exi_temporal_cantidad_salida,0) + " + CStr(var_cantidad_leida) + " where vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               
                  Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = " + txt_codigo
                  
                  rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_inserta = False
                     var_inserta = TB_TEMPORAL_SALIDAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida)
                     rs.Close
                     valor = Trim(txt_codigo)
                     Set itmfound = lv_traspasossalidas.findItem(valor, lvwText, , lvwPartial)
                     itmfound.EnsureVisible
                     itmfound.Selected = True
                     lv_traspasossalidas.selectedItem.SubItems(2) = lv_traspasossalidas.selectedItem.SubItems(2) + var_cantidad_leida
                     var_renglon = lv_traspasossalidas.selectedItem.Index
                     lbl_cantidad_total = CStr(CDbl(lbl_cantidad_total) + var_cantidad_leida)
                     Call ilumina_grid
                  Else
                     var_inserta = False
                     var_inserta = TB_TEMPORAL_SALIDAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", 0, 0)
                     rs.Close
                     Set list_item = lv_traspasossalidas.ListItems.Add(, , Trim(txt_codigo))
                     list_item.SubItems(1) = var_descripcion_articulo
                     list_item.SubItems(2) = var_cantidad_leida
                     var_renglon = lv_traspasossalidas.ListItems.Count
                     lbl_cantidad_total = CStr(CDbl(lbl_cantidad_total) + var_cantidad_leida)
                     Call ilumina_grid
                  End If
               Else
               
                  frmmensaje.lbl_mensaje = var_kanban_mensaje
                  frmmensaje.Show 1
                  txt_codigo = ""
               
               
               End If
            'Else
            '   rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            '   If Not rsaux4.EOF Then
            '      frmmensaje.lbl_articulo = IIf(IsNull(rsaux4!vcha_art_nombre_español), "", rsaux4!vcha_art_nombre_español)
            '   End If
            '   txt_codigo = ""
            '   frmmensaje.lbl_mensaje = "La cantidad supera al disponible en el almacen"
            '   frmmensaje.Show 1
            '   rsaux4.Close
            'End If
         Else
            rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux4.EOF Then
               frmmensaje.lbl_articulo = IIf(IsNull(rsaux4!vcha_Art_nombre_español), "", rsaux4!vcha_Art_nombre_español)
            End If
            txt_codigo = ""
            frmmensaje.lbl_mensaje = "El artículo no se encuentra en el inventario del almacén"
            frmmensaje.Show 1
            rsaux4.Close
         End If
         rsaux5.Close
      Else
         Me.txt_codigo = ""
         frmmensaje.lbl_mensaje = "La cantidad excede a la cantidad en existencias"
         frmmensaje.Show 1
      End If
      txt_codigo.SetFocus
   End If
End Sub


Private Sub txt_nombre_almacen_destino_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_almacen_destino_KeyDown(KeyCode As Integer, Shift As Integer)
   If txt_almacen_destino.Enabled = True Then
      If KeyCode = 116 Then
         lv_lista.ListItems.Clear
         If var_tipo_permiso = 1 Then
            rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and char_alm_tipo = 'G' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
         Else
            rs.Open "select * from vw_movimientos_almacenes WHERE VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and char_alm_tipo = 'G' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
         End If
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "Agentes"
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
   End If
End Sub

Private Sub txt_nombre_almacen_destino_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txt_codigo.Enabled = True Then
         txt_codigo.SetFocus
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_almacen_destino_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_almacen_origen_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_almacen_origen_KeyDown(KeyCode As Integer, Shift As Integer)
   If txt_almacen_origen.Enabled = True Then
      If KeyCode = 116 Then
         lv_lista.ListItems.Clear
         rs.Open "select * from vw_movimientos_almacenes WHERE VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and char_alm_tipo = 'A' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "Almacenes"
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
   End If
End Sub

Private Sub txt_nombre_almacen_origen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txt_almacen_destino.Enabled = True Then
         txt_almacen_destino.SetFocus
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_almacen_origen_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub



