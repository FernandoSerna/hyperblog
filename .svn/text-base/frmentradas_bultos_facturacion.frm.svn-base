VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmentradas_bultos_facturacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   345
      Picture         =   "frmentradas_bultos_facturacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Buscar Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   15
      Picture         =   "frmentradas_bultos_facturacion.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11130
      Picture         =   "frmentradas_bultos_facturacion.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Lectura de cajas "
      Height          =   840
      Left            =   15
      TabIndex        =   15
      Top             =   6345
      Width           =   11505
      Begin VB.CommandButton cmd_pasar_todos 
         Height          =   360
         Left            =   7185
         Picture         =   "frmentradas_bultos_facturacion.frx":083E
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   255
         Width           =   435
      End
      Begin VB.TextBox txt_caja 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   510
         Left            =   3705
         TabIndex        =   8
         Top             =   195
         Width           =   3405
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   675
      Picture         =   "frmentradas_bultos_facturacion.frx":0940
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   330
      TabIndex        =   12
      Top             =   1110
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         MaxLength       =   10
         TabIndex        =   13
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
         TabIndex        =   14
         Top             =   120
         Width           =   3060
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2625
      Left            =   420
      TabIndex        =   9
      Top             =   2265
      Width           =   5820
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2145
         Left            =   30
         TabIndex        =   10
         Top             =   420
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   3784
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
         Height          =   270
         Left            =   30
         TabIndex        =   11
         Top             =   120
         Width           =   5745
      End
   End
   Begin VB.Frame Frame3 
      Height          =   120
      Left            =   0
      TabIndex        =   16
      Top             =   255
      Width           =   11505
   End
   Begin VB.Frame Frame1 
      Height          =   1140
      Left            =   15
      TabIndex        =   19
      Top             =   360
      Width           =   11490
      Begin VB.TextBox txt_facturas 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   885
         TabIndex        =   6
         Top             =   600
         Width           =   2235
      End
      Begin VB.CommandButton cmd_buscar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3150
         Picture         =   "frmentradas_bultos_facturacion.frx":0A42
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Buscar Movimiento Alt + B"
         Top             =   630
         Width           =   330
      End
      Begin VB.TextBox txt_folio 
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
         Height          =   585
         Left            =   8595
         TabIndex        =   20
         Top             =   180
         Width           =   2805
      End
      Begin VB.TextBox txt_almacen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   885
         TabIndex        =   4
         Top             =   210
         Width           =   780
      End
      Begin VB.TextBox txt_nombre_almacen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   5
         Top             =   210
         Width           =   5295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Factura:"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   23
         Top             =   690
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
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
         Left            =   7680
         TabIndex        =   22
         Top             =   292
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Left            =   135
         TabIndex        =   21
         Top             =   300
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4785
      Left            =   15
      TabIndex        =   17
      Top             =   1500
      Width           =   11490
      Begin MSComctlLib.ListView lv_detalle_cajas 
         Height          =   4200
         Left            =   60
         TabIndex        =   18
         Top             =   135
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   7408
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código caja"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripción"
            Object.Width           =   8379
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cantidad Caja"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Caja leida"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Movimiento"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Costo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Precio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Estatus"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_cantidad_recibida 
         AutoSize        =   -1  'True
         Caption         =   "0.00"
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
         Left            =   7500
         TabIndex        =   27
         Top             =   4365
         Width           =   1605
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Piezas recibidas:"
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
         Left            =   5010
         TabIndex        =   26
         Top             =   4365
         Width           =   2385
      End
      Begin VB.Label lbl_cantidad_enviada 
         AutoSize        =   -1  'True
         Caption         =   "0.00"
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
         Left            =   3045
         TabIndex        =   25
         Top             =   4365
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Piezas enviadas:"
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
         Left            =   600
         TabIndex        =   24
         Top             =   4365
         Width           =   2355
      End
   End
End
Attribute VB_Name = "frmentradas_bultos_facturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_primera_vez As Integer
Dim var_numero_folio As Double
Dim var_estatus_movimiento As String
    
Dim var_tipo_lista As Integer
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report


Private Sub cmd_buscar_Click()
   If Trim(Me.txt_almacen) <> "" Then
      If Trim(Me.txt_facturas) <> "" Then
         var_serie_FACTURA = ""
         var_numero_factura = ""
         For var_j = 1 To Len(Me.txt_facturas)
             If IsNumeric(Mid(Me.txt_facturas, var_j, 1)) Then
                var_numero_factura = var_numero_factura + Mid(Me.txt_facturas, var_j, 1)
             Else
                var_serie_FACTURA = var_serie_FACTURA + Mid(Me.txt_facturas, var_j, 1)
             End If
         Next var_j
         If IsNumeric(var_numero_factura) Then
            Me.lv_detalle_cajas.ListItems.Clear
            var_cadena = "SELECT dbo.TB_ENTRADAS_BULTOS_FACTURACION.*, dbo.tb_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL AS vcha_art_nombre_español FROM dbo.TB_ENTRADAS_BULTOS_FACTURACION INNER JOIN dbo.tb_ARTICULOS ON dbo.TB_ENTRADAS_BULTOS_FACTURACION.VCHA_ART_ARTICULO_ID = dbo.tb_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_ENTRADAS_BULTOS_FACTURACION.VCHA_ENT_FACTURA = '" + Me.txt_facturas + "') order by vcha_Ent_caja"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               Me.cmd_buscar.Enabled = False
               Me.txt_facturas.Enabled = False
               Me.lbl_cantidad_enviada = 0#
               Me.lbl_cantidad_recibida = 0#
               While Not rs.EOF
                     Set list_item = Me.lv_detalle_cajas.ListItems.Add(, , rs!vcha_ent_Caja)
                     list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ART_ARTICULO_ID), "", rs!VCHA_ART_ARTICULO_ID)
                     list_item.SubItems(2) = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
                     list_item.SubItems(3) = IIf(IsNull(rs!FLOA_ENT_CANTIDAD_enviada), "", rs!FLOA_ENT_CANTIDAD_enviada)
                     list_item.SubItems(4) = IIf(IsNull(rs!FLOA_ENT_CANTIDAD_LEIDA), "", rs!FLOA_ENT_CANTIDAD_LEIDA)
                     list_item.SubItems(5) = 0
                     list_item.SubItems(6) = IIf(IsNull(rs!floa_ent_costo), "", rs!floa_ent_costo)
                     list_item.SubItems(7) = IIf(IsNull(rs!floa_ent_precio), "", rs!floa_ent_precio)
                     Me.lbl_cantidad_enviada = Format(CDbl(Me.lbl_cantidad_enviada) + IIf(IsNull(rs!FLOA_ENT_CANTIDAD_enviada), "", rs!FLOA_ENT_CANTIDAD_enviada), "###,###,##0.00")
                     Me.lbl_cantidad_recibida = Format(CDbl(Me.lbl_cantidad_recibida) + IIf(IsNull(rs!FLOA_ENT_CANTIDAD_LEIDA), "", rs!FLOA_ENT_CANTIDAD_LEIDA), "###,###,##0.00")
                     list_item.SubItems(8) = ""
                     rs.MoveNext:
               Wend
               var_primera_vez = 1
            Else
                Me.lv_detalle_cajas.ListItems.Clear
                var_conexion_facturas_ei = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=vianney;Data Source=DISTRIBUCION"
                If var_conexion_facturas_ei <> "" Then
                   If cnn_facturas_ei.State = 1 Then
                      cnn_facturas_ei.Close
                   End If
                   cnn_facturas_ei.Open var_conexion_facturas_ei
                   var_cadena = "SELECT floa_Emo_Descuento_1, floa_emo_Descuento_2, floa_emo_tipo_cambio, max(dbo.TB_DETALLE_CAJAS.FLOA_PAQ_COSTO) as costo, max(FLOA_PAQ_Precio) as precio, Dbo.TB_DETALLE_CAJAS.VCHA_ART_ARTICULO_ID, SUM(dbo.TB_DETALLE_CAJAS.FLOA_PAQ_CANTIDAD) AS CANTIDAD, dbo.TB_DETALLE_CAJAS.INTE_EMB_EMBARQUE, dbo.TB_DETALLE_CAJAS.INTE_PAQ_CAJA, dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID + CAST(dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO AS VARCHAR(50)) AS FACTURA, dbo.TB_Articulos.VCHA_ART_NOMBRE_ESPAÑOL FROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID "
                   var_cadena = var_cadena + " AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN "
                   var_cadena = var_cadena + " dbo.TB_DETALLE_CAJAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_CAJAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = dbo.TB_DETALLE_CAJAS.INTE_EMB_EMBARQUE INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO INNER JOIN dbo.tb_ARTICULOS ON dbo.TB_DETALLE_CAJAS.VCHA_ART_ARTICULO_ID = dbo.tb_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_CLI_CLAVE_ID = 'C000005397') AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID + CAST(dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO AS VARCHAR(50)) = '" + Me.txt_facturas + "') "
                   var_cadena = var_cadena + " GROUP BY  floa_Emo_Descuento_1, floa_emo_Descuento_2, floa_emo_tipo_cambio, dbo.TB_DETALLE_CAJAS.VCHA_ART_ARTICULO_ID, dbo.TB_DETALLE_CAJAS.INTE_EMB_EMBARQUE, dbo.TB_DETALLE_CAJAS.INTE_PAQ_CAJA, dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID + CAST(dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO AS VARCHAR(50)), dbo.TB_Articulos.VCHA_ART_NOMBRE_ESPAÑOL ORDER BY dbo.TB_DETALLE_CAJAS.INTE_PAQ_CAJA"
                   rsaux1.Open var_cadena, cnn_facturas_ei, adOpenDynamic, adLockOptimistic
                   If Not rsaux1.EOF Then
                      Me.cmd_buscar.Enabled = False
                      Me.txt_facturas.Enabled = False
                      var_primera_vez = 1
                      Me.lbl_cantidad_enviada = 0
                      Me.lbl_cantidad_enviada = 0
                      While Not rsaux1.EOF
                            txt_embarque = rsaux1!inte_emb_embarque
                            var_numero_caja = rsaux1!inte_paq_caja
                            var_referencia_caja = ""
                            var_contador = 0
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
                            If Len(Trim(Str(txt_embarque))) = 6 Then
                               var_referencia_embarque = Trim(Str(txt_embarque))
                            End If
                            var_referencia_embarque = "C" + var_referencia_embarque + var_referencia_caja
                            Set list_item = Me.lv_detalle_cajas.ListItems.Add(, , var_referencia_embarque)
                            list_item.SubItems(1) = IIf(IsNull(rsaux1!VCHA_ART_ARTICULO_ID), "", rsaux1!VCHA_ART_ARTICULO_ID)
                            list_item.SubItems(2) = IIf(IsNull(rsaux1!vcha_Art_nombre_español), "", rsaux1!vcha_Art_nombre_español)
                            list_item.SubItems(3) = IIf(IsNull(rsaux1!Cantidad), "", rsaux1!Cantidad)
                            Me.lbl_cantidad_enviada = Format(CDbl(Me.lbl_cantidad_enviada) + IIf(IsNull(rsaux1!Cantidad), 0, rsaux1!Cantidad), "###,###,##0.00")
                            list_item.SubItems(4) = 0
                            list_item.SubItems(5) = 0
                            rsaux10.Open "select * from tb_detalle_lista_precios where vcha_lis_lista_precios_id = '02' and vcha_Art_Articulo_id = '" + IIf(IsNull(rsaux1!VCHA_ART_ARTICULO_ID), "", rsaux1!VCHA_ART_ARTICULO_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
                            list_item.SubItems(6) = IIf(IsNull(rsaux10!floa_dli_Precio), 0, rsaux10!floa_dli_Precio)
                            rsaux10.Close
                            rsaux10.Open "select * from tb_detalle_lista_precios where vcha_lis_lista_precios_id = '50' and vcha_Art_Articulo_id = '" + IIf(IsNull(rsaux1!VCHA_ART_ARTICULO_ID), "", rsaux1!VCHA_ART_ARTICULO_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
                            list_item.SubItems(7) = IIf(IsNull(rsaux10!floa_dli_Precio), "", rsaux10!floa_dli_Precio)
                            rsaux10.Close
                            list_item.SubItems(8) = ""
                            var_cadena = "insert into TB_ENTRADAS_BULTOS_FACTURACION (vcha_emp_empresa_id, vcha_ser_serie_id,inte_Car_numero, VCHA_ENT_FACTURA, vcha_ent_Caja, vcha_Art_Articulo_id, floa_Ent_Cantidad_enviada, floa_Ent_cantidad_leida, floa_ent_precio, floa_ent_costo, vcha_ent_maquina, vcha_Ent_usuario, dtim_ent_fecha, VCHA_ALM_ALMACEN_ID, floa_ent_descuento_1, floa_ent_descuento_2, floa_Ent_descuento_3, floa_ent_tipo_cambio) "
                            var_cadena = var_cadena + "                          values ('" + var_empresa + "','" + var_serie_FACTURA + "'," + CStr(var_numero_factura) + ",'" + rsaux1!FACTURA + "','" + var_referencia_embarque + "', '" + rsaux1!VCHA_ART_ARTICULO_ID + "'," + CStr(rsaux1!Cantidad) + ", 0," + CStr(rsaux1!Precio) + " ," + CStr(rsaux1!Costo) + ",'" + fun_NombrePc + "','" + var_clave_usuario_global + "', getdate(),'" + Me.txt_almacen + "','" + CStr(IIf(IsNull(rsaux1!floa_emo_descuento_1), 0, rsaux1!floa_emo_descuento_1)) + "'," + CStr(IIf(IsNull(rsaux1!floa_emo_descuento_2), 0, rsaux1!floa_emo_descuento_2)) + ",0," + CStr(IIf(IsNull(rsaux1!floa_emo_tipo_cambio), 1, rsaux1!floa_emo_tipo_cambio)) + ")"
                            rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                           rsaux1.MoveNext:
                      Wend
                   Else
                      MsgBox "La factura no existe", vbOKOnly, "ATENCION"
                   End If
                   rsaux1.Close
                Else
                   MsgBox "La planta no cuenta con conexión", vbOKOnly, "ATENCION"
                End If
            End If
            rs.Close
         Else
            MsgBox "Factura incorrecta", vbOKOnly, "ATENCION"
        End If
      Else
         MsgBox "No se a indicado ninguna factura", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a indicado un almacén", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_cerrar_Click()
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
   If Me.lv_detalle_cajas.ListItems.Count > 0 Then
      If Trim(Me.txt_facturas) <> "" Then
         var_serie_FACTURA = ""
         var_numero_factura = ""
         For var_j = 1 To Len(Me.txt_facturas)
             If IsNumeric(Mid(Me.txt_facturas, var_j, 1)) Then
                var_numero_factura = var_numero_factura + Mid(Me.txt_facturas, var_j, 1)
             Else
                var_serie_FACTURA = var_serie_FACTURA + Mid(Me.txt_facturas, var_j, 1)
             End If
         Next var_j
         If IsNumeric(var_numero_factura) Then
            rs.Open "SELECT * FROM  TB_ENTRADAS_BULTOS_INTERCOMPAÑIAS  WHERE VCHA_ENT_PLANTA_PROVEEDOR = '" + txt_planta + "' AND VCHA_SER_SERIE_ID = '" + var_serie_FACTURA + "' AND INTE_CAR_NUMERO = " + var_numero_factura + " ", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_estatus = IIf(IsNull(rs!vcha_Ent_Estatus), "", rs!vcha_Ent_Estatus)
               If var_estatus = "" Then
                  var_si = MsgBox("Desea cerrar la lectura de las cajas", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     If var_si = 6 Then
                        VAR_FALTAN_CAJAS = ""
                        For var_j = 1 To Me.lv_detalle_cajas.ListItems.Count
                            Me.lv_detalle_cajas.ListItems.item(var_j).Selected = True
                            If CDbl(Me.lv_detalle_cajas.selectedItem.SubItems(4)) = 0 Then
                               VAR_FALTAN_CAJAS = "FALTAN"
                            End If
                        Next var_j
                        If VAR_FALTAN_CAJAS = "" Then
                           rsaux.Open "update TB_ENTRADAS_BULTOS_INTERCOMPAÑIAS  set VCHA_ent_estatus = 'I', DTIM_ENT_FECHA = GETDATE()  WHERE VCHA_ENT_PLANTA_PROVEEDOR = '" + txt_planta + "' AND VCHA_SER_SERIE_ID = '" + var_serie_FACTURA + "' AND INTE_CAR_NUMERO = " + var_numero_factura + " ", cnn, adOpenDynamic, adLockOptimistic
                           For var_j = 1 To Me.lv_detalle_cajas.ListItems.Count
                               Me.lv_detalle_cajas.ListItems.item(var_j).Selected = True
                               Me.lv_detalle_cajas.selectedItem.SubItems(9) = "I"
                           Next var_j
                           var_numero_factura = CStr(CDbl(var_numero_factura))
                           Set reporte = appl.OpenReport(App.Path + "\REP_BULTOS_ENTRADAS.rpt")
                           reporte.RecordSelectionFormula = "{VW_BULTOS_ENTRADAS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_BULTOS_ENTRADAS.VCHA_ENT_PLANTA_PROVEEDOR} = '" + txt_planta + "' AND {VW_BULTOS_ENTRADAS.VCHA_SER_SERIE_ID} = '" + var_serie_FACTURA + "' AND {VW_BULTOS_ENTRADAS.INTE_CAR_NUMERO} = '" + var_numero_factura + "'"
                           frmvistasprevias.cr.ReportSource = reporte
                           For ntablas = 1 To reporte.Database.Tables.Count
                               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                           Next ntablas
                           frmvistasprevias.cr.ViewReport
                           frmvistasprevias.Caption = "Reporte de Movimientos"
                           frmvistasprevias.Show 1
                           Set reporte = Nothing
                        Else
                           MsgBox "Faltan cajas por leer", vbOKOnly, "ATENCION"
                        End If
                     End If
                  End If
               Else
                  var_numero_factura = CStr(CDbl(var_numero_factura))
                  Set reporte = appl.OpenReport(App.Path + "\REP_BULTOS_ENTRADAS.rpt")
                  reporte.RecordSelectionFormula = "{VW_BULTOS_ENTRADAS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_BULTOS_ENTRADAS.VCHA_ENT_PLANTA_PROVEEDOR} = '" + txt_planta + "' AND {VW_BULTOS_ENTRADAS.VCHA_SER_SERIE_ID} = '" + var_serie_FACTURA + "' AND {VW_BULTOS_ENTRADAS.INTE_CAR_NUMERO} = '" + var_numero_factura + "'"
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Movimientos"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
               End If
            End If
            rs.Close
         End If
      End If
   End If
End Sub

Private Sub cmd_imprimir_Click()
   Dim pError As ADODB.Error
   Dim var_codigo_barras_caja As String
   Dim var_actualiza As Boolean
   Dim var_inserta As Boolean
   Dim bandera_suma As Boolean
   Dim var_cantidad As Variant
   Dim var_costo As Variant
   Dim var_precio As Variant
   Dim var_consecutivo_serie  As Double
   Dim var_posible As Boolean
   Dim var_P_RC_LINEA_ID As Double
   Dim var_P_RC_NUMERO_LINEA As Double
   Set TB_ARCH_COMPARACION_M = New TB_ARCH_COMPARACION_M
   Set TB_ARCH_COMPARACION_I = New TB_ARCH_COMPARACION_I
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   Dim var_posible_lectura_kanban As Boolean
   'On Error GoTo salir:
   cnn.CommandTimeout = 360
   
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_ENTRADAS_I = New TB_ENTRADAS_I
   Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_ENTRADAS_VISTAS_I = New TB_ENTRADAS_VISTAS_I
   'On Error GoTo salir:
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
   For var_j = 1 To Me.lv_detalle_cajas.ListItems.Count
       Me.lv_detalle_cajas.ListItems.item(var_j).Selected = True
   Next var_j
   var_almacen_Destino = Me.txt_almacen
   If IsNumeric(Me.txt_folio) Then
      If var_estatus_movimiento = "I" Then
         Set reporte = appl.OpenReport(App.Path + "\REP_ENTRADAS_BULTOS_FACTURACION.rpt")
         reporte.RecordSelectionFormula = "{VW_ENTRADAS_BULTOS_FACTURACION.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_ENTRADAS_BULTOS_FACTURACION.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' AND {VW_ENTRADAS_BULTOS_FACTURACION.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_ENTRADAS_BULTOS_FACTURACION.INTE_ENT_NUMERO} = " + Me.txt_folio
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Movimientos"
         frmvistasprevias.Show 1
         Set reporte = Nothing
      Else
         var_si = MsgBox("¿Desea cerrar el movimiento", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar el cerrado del movimiento", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_almacen_Destino = Me.txt_almacen
               var_almacen_origen = Me.txt_almacen
               var_numero_folio = CDbl(Me.txt_folio)
               x = 0
               Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
               rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
               VAR_ZZ = rs.RecordCount
               var_posible_entrada = True
               If rs.RecordCount > 0 Then
                 rs.MoveFirst
               End If
               If var_posible_entrada = True Then
                  cnn.BeginTrans
                  While Not rs.EOF
                        rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, inte_ent_año, vcha_ent_referencia) values ('" + CStr(var_empresa) + "', '" + CStr(var_unidad_organizacional) + "', '" + CStr(var_almacen_Destino) + "', '" + CStr(var_clave_movimiento) + "', " + CStr(CDbl(var_numero_folio)) + ", '" + rs!VCHA_ART_ARTICULO_ID + "'," + CStr(rs!floa_ent_cantidaD) + "," + CStr(rs!floa_ent_costo) + "," + CStr(rs!floa_ent_precio) + ",2005,'" + rs!vcha_ent_Caja + "') ", cnn, adOpenDynamic, adLockOptimistic
                        'var_inserta = TB_ENTRADAS_I.Anadir(CStr(var_empresa), CStr(var_unidad_organizacional), CStr(var_almacen_Destino), CStr(var_clave_movimiento), CDbl(var_numero_folio), 1, CStr(var_almacen_origen), 0)
                        rs.MoveNext
                  Wend
                  rs.Close
                  var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(CStr(var_empresa), CStr(var_unidad_organizacional), CStr(var_almacen_Destino), CStr(var_clave_movimiento), CDbl(var_numero_folio), "", Now, 1)
                  var_estatus_movimiento = "I"
                  var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(CStr(var_empresa), CStr(var_unidad_organizacional), CStr(var_almacen_Destino), CStr(var_clave_movimiento), CDbl(var_numero_folio), "I", Now, 1)
                  cnn.CommitTrans
                  Me.txt_caja.Enabled = False
                  Set reporte = appl.OpenReport(App.Path + "\REP_ENTRADAS_BULTOS_FACTURACION.rpt")
                  reporte.RecordSelectionFormula = "{VW_ENTRADAS_BULTOS_FACTURACION.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_ENTRADAS_BULTOS_FACTURACION.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' AND {VW_ENTRADAS_BULTOS_FACTURACION.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_ENTRADAS_BULTOS_FACTURACION.INTE_ENT_NUMERO} = " + Me.txt_folio
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                     reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Movimientos"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
               End If
               If rs.State = 1 Then
                  rs.Close
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub cmd_nuevo_Click()
  var_primera_vez = 1
  Me.txt_almacen = ""
  Me.txt_nombre_almacen = ""
  Me.txt_facturas = ""
  Me.lv_detalle_cajas.ListItems.Clear
  Me.txt_almacen.Enabled = True
  Me.txt_facturas.Enabled = True
  Me.txt_almacen.SetFocus
  var_estatus_movimiento = ""
  Me.cmd_buscar.Enabled = True
  Me.txt_facturas.Enabled = True
  Me.txt_caja.Enabled = True
End Sub

Private Sub cmd_pasar_todos_Click()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   var_si = MsgBox("¿Desea pasar todas las cajas?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_si = MsgBox("Confirmar pasar todas las cajas", vbYesNo, "ATENCION")
      If var_si = 6 Then
         For var_j = 1 To Len(Me.txt_facturas)
             If IsNumeric(Mid(Me.txt_facturas, var_j, 1)) Then
                var_numero_factura = var_numero_factura + Mid(Me.txt_facturas, var_j, 1)
             Else
                var_serie_FACTURA = var_serie_FACTURA + Mid(Me.txt_facturas, var_j, 1)
             End If
         Next var_j

         rsaux11.Open "SELECT DISTINCT VCHA_eNT_CAJA FROM TB_ENTRADAS_BULTOS_facturacion WHERE vcha_ent_factura = '" + Me.txt_facturas + "' AND VCHA_SER_sERIE_ID = '" + var_serie_FACTURA + "' AND INTE_cAR_NUMERO = " + var_numero_factura, cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux11.EOF
               Me.txt_caja = rsaux11(0).Value
               var_almacen_Destino = Me.txt_almacen
               If Me.lv_detalle_cajas.ListItems.Count > 0 Then
                  If Mid(Me.txt_caja, 1, 1) = "C" Then
                     var_serie_FACTURA = ""
                     var_numero_factura = ""
                     For var_j = 1 To Len(Me.txt_facturas)
                         If IsNumeric(Mid(Me.txt_facturas, var_j, 1)) Then
                            var_numero_factura = var_numero_factura + Mid(Me.txt_facturas, var_j, 1)
                         Else
                            var_serie_FACTURA = var_serie_FACTURA + Mid(Me.txt_facturas, var_j, 1)
                         End If
                     Next var_j
                     If IsNumeric(var_numero_factura) Then
                        rs.Open "SELECT * FROM TB_ENTRADAS_BULTOS_facturacion WHERE vcha_ent_factura = '" + Me.txt_facturas + "' AND VCHA_SER_sERIE_ID = '" + var_serie_FACTURA + "' AND INTE_cAR_NUMERO = " + var_numero_factura + " AND VCHA_ENT_CAJA = '" + Me.txt_caja + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_cantidad_caja = IIf(IsNull(rs!FLOA_ENT_CANTIDAD_LEIDA), 0, rs!FLOA_ENT_CANTIDAD_LEIDA)
                           If var_cantidad_caja > 0 Then
                              Me.txt_caja = ""
                              frmmensaje.lbl_mensaje = "La caja ya fue leida"
                              frmmensaje.Show
                           Else
                              For var_j = 1 To Me.lv_detalle_cajas.ListItems.Count
                                  Me.lv_detalle_cajas.ListItems.item(var_j).Selected = True
                                  If Me.lv_detalle_cajas.selectedItem = Me.txt_caja Then
                                     rsaux.Open "update TB_ENTRADAS_BULTOS_facturacion set floa_ent_Cantidad_leida = floa_ent_cantidad_enviada where  vcha_Ent_factura = '" + Me.txt_facturas + "' AND VCHA_SER_sERIE_ID = '" + var_serie_FACTURA + "' AND INTE_cAR_NUMERO = " + var_numero_factura + " AND VCHA_ENT_CAJA = '" + Me.txt_caja + "'", cnn, adOpenDynamic, adLockOptimistic
                                     If var_primera_vez = 1 Then
                                        var_numero_folio = 0
                                        var_folio_enviado = 0
                                        var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, CStr(var_almacen_Destino), var_clave_movimiento, Now, CDbl(var_numero_folio), CDbl(var_folio_enviado), "", CStr(var_proveedor), CStr(var_almacen_origen), CStr(var_almacen_Destino), "", var_clave_usuario_global, fun_NombrePc, CStr(var_factura), "", CStr(Me.txt_facturas), "", "B", "", "", 0, 0, 0, CStr(var_clave_moneda), CDbl(var_tipo_Cambio))
                                        var_numero_folio = var_numero_folio_regreso
                                        Me.txt_folio = var_numero_folio
                                        var_primera_vez = 0
                                     End If
                                     rsaux3.Open "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(Me.txt_folio) + " and vcha_Art_articulo_id = '" + Me.lv_detalle_cajas.selectedItem.SubItems(1) + "' AND vcha_ent_caja = '" + Me.lv_detalle_cajas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                                     If Not rsaux3.EOF Then
                                        rsaux.Open "update tb_temporal_Entradas set floa_ent_cantidad = floa_ent_cantidad + " + Me.lv_detalle_cajas.selectedItem.SubItems(3) + " where vcha_emp_empresa_id  = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(Me.txt_folio) + " and vcha_Art_articulo_id = '" + Me.lv_detalle_cajas.selectedItem.SubItems(1) + "' and vcha_ent_caja = '" + Me.lv_detalle_cajas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                                     Else
                                        var_costo = CDbl(Me.lv_detalle_cajas.selectedItem.SubItems(6))
                                        var_precio = CDbl(Me.lv_detalle_cajas.selectedItem.SubItems(7))
                                        rsaux.Open "insert into tb_temporal_Entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_precio, floa_ent_costo, floa_ent_Cantidad, vcha_ent_Caja) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "'," + CStr(Me.txt_folio) + ",'" + Me.lv_detalle_cajas.selectedItem.SubItems(1) + "', " + CStr(var_precio) + "," + CStr(var_costo) + ", " + Me.lv_detalle_cajas.selectedItem.SubItems(3) + ",'" + Me.lv_detalle_cajas.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                                     End If
                                     rsaux3.Close
                                     Me.lbl_cantidad_recibida = Format(CDbl(Me.lbl_cantidad_recibida) + CDbl(Me.lv_detalle_cajas.selectedItem.SubItems(3)), "###,###,##0.00")
                                     Me.lv_detalle_cajas.selectedItem.SubItems(5) = Me.lv_detalle_cajas.selectedItem.SubItems(3)
                                     Me.lv_detalle_cajas.selectedItem.SubItems(4) = Me.lv_detalle_cajas.selectedItem.SubItems(3)
                                     Me.lv_detalle_cajas.ListItems.item(var_j).EnsureVisible
                                  End If
                              Next var_j
                              Me.txt_caja = ""
                           End If
                        Else
                           frmmensaje.lbl_mensaje = "La caja no se encuentra en la nota"
                           frmmensaje.Show 1
                           Me.txt_caja = ""
                        End If
                        rs.Close
                     Else
                        frmmensaje.lbl_mensaje = "Factura incorrecta"
                        frmmensaje.Show 1
                     End If
                  Else
                     MsgBox "Código de caja invalida", vbOKOnly, "ATENCION"
                  End If
               End If
               rsaux11.MoveNext
         Wend
         rsaux11.Close
      End If
   End If
   
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command2_Click()
   Me.frm_busqueda.Visible = True
   Me.txt_busqueda_folio = ""
   Me.txt_busqueda_folio.SetFocus
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
   var_primera_vez = 1
   Me.frm_busqueda.Visible = False
   Me.frm_lista.Visible = False
   var_estatus_movimiento = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_entradas_sin_comparacion)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If var_tipo_lista = 1 Then
      If Me.lv_lista.ListItems.Count > 0 Then
         Me.txt_almacen = lv_lista.selectedItem
         Me.txt_nombre_almacen = Me.lv_lista.selectedItem.SubItems(1)
      End If
      Me.txt_almacen.SetFocus
   End If
   
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      End If
      'rs.Open "select distinct vcha_cli_nombre from vw_establecimientos where vcha_esb_establecimiento_id = '" + txt_establecimiento + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
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

Private Sub txt_almacen_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.txt_facturas.SetFocus
   End If
End Sub

Private Sub txt_almacen_LostFocus()
   If Trim(Me.txt_almacen) <> "" Then
      rs.Open "SELECT dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_MOVIMIENTOS_ALMACENES.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ALMACENES.VCHA_EMP_EMPRESA_ID FROM dbo.TB_MOVIMIENTOS_ALMACENES INNER JOIN dbo.TB_ALMACENES ON dbo.TB_MOVIMIENTOS_ALMACENES.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID WHERE (dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID = '" + Me.txt_almacen + "') AND (dbo.TB_ALMACENES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_MOVIMIENTOS_ALMACENES.VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "')", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_almacen = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
      Else
         MsgBox "El almacén no existe o no tiene permisos para este movimiento", vbOKOnly, "ATENCION"
         Me.txt_almacen = ""
         Me.txt_nombre_almacen = ""
      End If
      rs.Close
      Me.txt_almacen.Enabled = False
      Me.txt_nombre_almacen.Enabled = False
   Else
      Me.txt_nombre_almacen = ""
   End If
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_busqueda_folio) Then
         rs.Open "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + Me.txt_busqueda_folio, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_estatus_movimiento = IIf(IsNull(rs!char_Emo_estatus), "", rs!char_Emo_estatus)
            var_almacen_Destino = rs!VCHA_ALM_ALMACEN_ID
            Me.txt_almacen = rs!VCHA_ALM_ALMACEN_ID
            rsaux.Open "SELECT * FROM TB_ALMACENES WHERE VCHA_ALM_ALMACEN_ID = '" + Me.txt_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_nombre_almacen = rsaux!VCHA_ALM_NOMBRE
            End If
            rsaux.Close
            
            Me.txt_folio = CStr(rs!INTE_EMO_NUMERO)
            Me.txt_facturas = rs!vcha_Emo_referencia
            
            var_serie_FACTURA = ""
            var_numero_factura = ""
            For var_j = 1 To Len(Me.txt_facturas)
                If IsNumeric(Mid(Me.txt_facturas, var_j, 1)) Then
                   var_numero_factura = var_numero_factura + Mid(Me.txt_facturas, var_j, 1)
                Else
                   var_serie_FACTURA = var_serie_FACTURA + Mid(Me.txt_facturas, var_j, 1)
                End If
            Next var_j
            Me.lbl_cantidad_enviada = 0#
            Me.lbl_cantidad_recibida = 0#
            If IsNumeric(var_numero_factura) Then
               Me.lv_detalle_cajas.ListItems.Clear
               rsaux.Open "SELECT dbo.TB_ENTRADAS_BULTOS_FACTURACION.*, dbo.tb_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL AS vcha_art_nombre_español FROM dbo.TB_ENTRADAS_BULTOS_FACTURACION INNER JOIN dbo.tb_ARTICULOS ON dbo.TB_ENTRADAS_BULTOS_FACTURACION.VCHA_ART_ARTICULO_ID = dbo.tb_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_ENTRADAS_BULTOS_FACTURACION.VCHA_ENT_FACTURA = '" + Me.txt_facturas + "') order by vcha_Ent_caja", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  While Not rsaux.EOF
                        Set list_item = Me.lv_detalle_cajas.ListItems.Add(, , rsaux!vcha_ent_Caja)
                        list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_ART_ARTICULO_ID), "", rsaux!VCHA_ART_ARTICULO_ID)
                        list_item.SubItems(2) = IIf(IsNull(rsaux!vcha_Art_nombre_español), "", rsaux!vcha_Art_nombre_español)
                        list_item.SubItems(3) = IIf(IsNull(rsaux!FLOA_ENT_CANTIDAD_enviada), "", rsaux!FLOA_ENT_CANTIDAD_enviada)
                        list_item.SubItems(4) = IIf(IsNull(rsaux!FLOA_ENT_CANTIDAD_LEIDA), "", rsaux!FLOA_ENT_CANTIDAD_LEIDA)
                        list_item.SubItems(5) = 0
                        list_item.SubItems(6) = IIf(IsNull(rsaux!floa_ent_costo), "", rsaux!floa_ent_costo)
                        list_item.SubItems(7) = IIf(IsNull(rsaux!floa_ent_precio), "", rsaux!floa_ent_precio)
                        list_item.SubItems(8) = ""
                        Me.lbl_cantidad_enviada = Format(CDbl(Me.lbl_cantidad_enviada) + IIf(IsNull(rsaux!FLOA_ENT_CANTIDAD_enviada), "", rsaux!FLOA_ENT_CANTIDAD_enviada), "###,###,##0.00")
                        Me.lbl_cantidad_recibida = Format(CDbl(Me.lbl_cantidad_recibida) + IIf(IsNull(rsaux!FLOA_ENT_CANTIDAD_LEIDA), "", rsaux!FLOA_ENT_CANTIDAD_LEIDA), "###,###,##0.00")
                        rsaux.MoveNext:
                  Wend
               End If
               rsaux.Close
               rsaux.Open "select * from tb_temporal_Entradas where  vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + Me.txt_busqueda_folio, cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux.EOF
                     For var_j = 1 To Me.lv_detalle_cajas.ListItems.Count
                         Me.lv_detalle_cajas.ListItems.item(var_j).Selected = True
                         If rsaux!VCHA_ART_ARTICULO_ID = Me.lv_detalle_cajas.selectedItem.SubItems(1) And rsaux!vcha_ent_Caja = Me.lv_detalle_cajas.selectedItem Then
                            Me.lv_detalle_cajas.selectedItem.SubItems(5) = Me.lv_detalle_cajas.selectedItem.SubItems(5) + rsaux!floa_ent_cantidaD
                         End If
                     Next var_j
                     rsaux.MoveNext
               Wend
               rsaux.Close
            End If
            
            
            
            Me.frm_busqueda.Visible = False
            Me.txt_almacen.Enabled = False
            Me.txt_nombre_almacen.Enabled = False
            Me.txt_facturas.Enabled = False
            var_primera_vez = 0
            Me.cmd_buscar.Enabled = False
            If var_estatus_movimiento = "I" Then
               Me.txt_caja.Enabled = False
            Else
               Me.txt_caja.Enabled = True
            End If
         Else
            MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
            Me.txt_almacen = ""
            Me.txt_folio = ""
            Me.txt_nombre_almacen = ""
            Me.lv_detalle_cajas.ListItems.Clear
            Me.txt_facturas = ""
            
         End If
         rs.Close
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_busqueda_folio_LostFocus()
   Me.frm_busqueda.Visible = False
End Sub

Private Sub txt_caja_KeyPress(KeyAscii As Integer)
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      var_almacen_Destino = Me.txt_almacen
      If Me.lv_detalle_cajas.ListItems.Count > 0 Then
         If Mid(Me.txt_caja, 1, 1) = "C" Then
            var_serie_FACTURA = ""
            var_numero_factura = ""
            For var_j = 1 To Len(Me.txt_facturas)
                If IsNumeric(Mid(Me.txt_facturas, var_j, 1)) Then
                   var_numero_factura = var_numero_factura + Mid(Me.txt_facturas, var_j, 1)
                Else
                   var_serie_FACTURA = var_serie_FACTURA + Mid(Me.txt_facturas, var_j, 1)
                End If
            Next var_j
            If IsNumeric(var_numero_factura) Then
               rs.Open "SELECT * FROM TB_ENTRADAS_BULTOS_facturacion WHERE vcha_ent_factura = '" + Me.txt_facturas + "' AND VCHA_SER_sERIE_ID = '" + var_serie_FACTURA + "' AND INTE_cAR_NUMERO = " + var_numero_factura + " AND VCHA_ENT_CAJA = '" + Me.txt_caja + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_cantidad_caja = IIf(IsNull(rs!FLOA_ENT_CANTIDAD_LEIDA), 0, rs!FLOA_ENT_CANTIDAD_LEIDA)
                  If var_cantidad_caja > 0 Then
                     Me.txt_caja = ""
                     frmmensaje.lbl_mensaje = "La caja ya fue leida"
                     frmmensaje.Show
                  Else
                     For var_j = 1 To Me.lv_detalle_cajas.ListItems.Count
                         Me.lv_detalle_cajas.ListItems.item(var_j).Selected = True
                         If Me.lv_detalle_cajas.selectedItem = Me.txt_caja Then
                            rsaux.Open "update TB_ENTRADAS_BULTOS_facturacion set floa_ent_Cantidad_leida = floa_ent_cantidad_enviada where  vcha_Ent_factura = '" + Me.txt_facturas + "' AND VCHA_SER_sERIE_ID = '" + var_serie_FACTURA + "' AND INTE_cAR_NUMERO = " + var_numero_factura + " AND VCHA_ENT_CAJA = '" + Me.txt_caja + "'", cnn, adOpenDynamic, adLockOptimistic
                            
                            If var_primera_vez = 1 Then
                               var_numero_folio = 0
                               var_folio_enviado = 0
                               var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, CStr(var_almacen_Destino), var_clave_movimiento, Now, CDbl(var_numero_folio), CDbl(var_folio_enviado), "", CStr(var_proveedor), CStr(var_almacen_origen), CStr(var_almacen_Destino), "", var_clave_usuario_global, fun_NombrePc, CStr(var_factura), "", CStr(Me.txt_facturas), "", "B", "", "", 0, 0, 0, CStr(var_clave_moneda), CDbl(var_tipo_Cambio))
                               var_numero_folio = var_numero_folio_regreso
                               Me.txt_folio = var_numero_folio
                               var_primera_vez = 0
                            End If
                            
                            rsaux3.Open "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(Me.txt_folio) + " and vcha_Art_articulo_id = '" + Me.lv_detalle_cajas.selectedItem.SubItems(1) + "' AND vcha_ent_caja = '" + Me.lv_detalle_cajas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                            If Not rsaux3.EOF Then
                               rsaux.Open "update tb_temporal_Entradas set floa_ent_cantidad = floa_ent_cantidad + " + Me.lv_detalle_cajas.selectedItem.SubItems(3) + " where vcha_emp_empresa_id  = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(Me.txt_folio) + " and vcha_Art_articulo_id = '" + Me.lv_detalle_cajas.selectedItem.SubItems(1) + "' and vcha_ent_caja = '" + Me.lv_detalle_cajas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                               
                            Else
                               var_costo = CDbl(Me.lv_detalle_cajas.selectedItem.SubItems(6))
                               var_precio = CDbl(Me.lv_detalle_cajas.selectedItem.SubItems(7))
                               rsaux.Open "insert into tb_temporal_Entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_precio, floa_ent_costo, floa_ent_Cantidad, vcha_ent_Caja) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "'," + CStr(Me.txt_folio) + ",'" + Me.lv_detalle_cajas.selectedItem.SubItems(1) + "', " + CStr(var_precio) + "," + CStr(var_costo) + ", " + Me.lv_detalle_cajas.selectedItem.SubItems(3) + ",'" + Me.lv_detalle_cajas.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                            End If
                            rsaux3.Close
                            Me.lbl_cantidad_recibida = Format(CDbl(Me.lbl_cantidad_recibida) + CDbl(Me.lv_detalle_cajas.selectedItem.SubItems(3)), "###,###,##0.00")
                            Me.lv_detalle_cajas.selectedItem.SubItems(5) = Me.lv_detalle_cajas.selectedItem.SubItems(3)
                            Me.lv_detalle_cajas.selectedItem.SubItems(4) = Me.lv_detalle_cajas.selectedItem.SubItems(3)
                            Me.lv_detalle_cajas.ListItems.item(var_j).EnsureVisible
                         End If
                     Next var_j
                     Me.txt_caja = ""
                  End If
               Else
                  frmmensaje.lbl_mensaje = "La caja no se encuentra en la nota"
                  frmmensaje.Show 1
                  Me.txt_caja = ""
               End If
               rs.Close
            Else
               frmmensaje.lbl_mensaje = "Factura incorrecta"
               frmmensaje.Show 1
            End If
         Else
            MsgBox "Código de caja invalida", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   If KeyAscii = 13 Then
      If txt_codigo <> "" Then
         rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            rsaux.Open "SELECT * FROM TB_EQUIVALENCIAS WHERE VCHA_EQU_CODIGO_EQUIVALENTE = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               rsaux1.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + rsaux!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  txt_codigo = rsaux1!VCHA_ART_ARTICULO_ID
               End If
               rsaux1.Close
            End If
            rsaux.Close
         End If
         rs.Close
         If txt_codigo = Me.lv_detalle_cajas.selectedItem.SubItems(1) Then
            If CDbl(Me.lv_detalle_cajas.selectedItem.SubItems(4)) >= CDbl(Me.lv_detalle_cajas.selectedItem.SubItems(5)) + 1 Then
               var_serie_FACTURA = ""
               var_numero_factura = ""
               For var_j = 1 To Len(Me.txt_facturas)
                   If IsNumeric(Mid(Me.txt_facturas, var_j, 1)) Then
                      var_numero_factura = var_numero_factura + Mid(Me.txt_facturas, var_j, 1)
                   Else
                      var_serie_FACTURA = var_serie_FACTURA + Mid(Me.txt_facturas, var_j, 1)
                   End If
               Next var_j
               var_almacen_Destino = Me.txt_almacen
               var_almacen_origen = Me.txt_almacen
               var_factura = var_numero_factura
               var_clave_moneda = "1"
               var_tipo_Cambio = 1
               If IsNumeric(var_numero_factura) Then
                  If var_primera_vez = 1 Then
                     var_numero_folio = 0
                     var_folio_enviado = 0
                     var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, CStr(var_almacen_Destino), var_clave_movimiento, Now, CDbl(var_numero_folio), CDbl(var_folio_enviado), "", CStr(var_proveedor), CStr(var_almacen_origen), CStr(var_almacen_Destino), "", var_clave_usuario_global, fun_NombrePc, CStr(var_factura), "", CStr(Me.txt_facturas), "", "B", "", "", 0, 0, 0, CStr(var_clave_moneda), CDbl(var_tipo_Cambio))
                     var_numero_folio = var_numero_folio_regreso
                     Me.txt_folio = var_numero_folio
                     var_primera_vez = 0
                  End If
                  rs.Open "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(Me.txt_folio) + " and vcha_Art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     rsaux.Open "update tb_temporal_Entradas set floa_ent_cantidad = floa_ent_cantidad + 1 where vcha_emp_empresa_id  = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(Me.txt_folio) + " and vcha_Art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     var_costo = CDbl(Me.lv_detalle_cajas.selectedItem.SubItems(7))
                     var_precio = CDbl(Me.lv_detalle_cajas.selectedItem.SubItems(8))
                     rsaux.Open "insert into tb_temporal_Entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_precio, floa_ent_costo, floa_ent_Cantidad, vcha_ent_Caja) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "'," + CStr(Me.txt_folio) + ",'" + txt_codigo + "', " + CStr(var_precio) + "," + CStr(var_costo) + ", 1,'" + Me.lv_detalle_cajas.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rs.Close
                  'rsaux.Open "update TB_ENTRADAS_BULTOS_INTERCOMPAÑIAS set floa_Ent_Cantidad_piezas_leida = isnull(floa_Ent_Cantidad_piezas_leida,0) + 1 where  VCHA_ENT_PLANTA_PROVEEDOR = '" + txt_planta + "' AND VCHA_SER_sERIE_ID = '" + VAR_SERIE_FACTURA + "' AND INTE_cAR_NUMERO = " + var_numero_factura + " AND VCHA_ENT_CAJA = '" + Me.lbl_caja + "' AND VCHA_aRT_ARTICULO_ID = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  Me.lv_detalle_cajas.selectedItem.SubItems(5) = CDbl(Me.lv_detalle_cajas.selectedItem.SubItems(5)) + 1
                  Me.lv_detalle_cajas.selectedItem.SubItems(6) = CDbl(Me.lv_detalle_cajas.selectedItem.SubItems(6)) + 1
                  txt_codigo = ""
               Else
               End If
            Else
               frmmensaje.lbl_mensaje = "La cantidad excede a la cantidad que viene en la caja"
               frmmensaje.Show 1
               Me.txt_caja = ""
            End If
         Else
            rs.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_DEscripcion = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
            End If
            rs.Close
            frmmensaje.lbl_mensaje = "El código no viene en la caja"
            frmmensaje.lbl_articulo = var_DEscripcion
            frmmensaje.Show 1
            Me.txt_caja = ""
         End If
      End If
   End If
   If KeyAscii = 27 Then
      Me.lv_detalle_cajas.SetFocus
   End If
End Sub


Private Sub txt_facturas_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.cmd_buscar.SetFocus
   End If
End Sub

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_facturas.SetFocus
   Else
      If KeyAscii = 27 Then
      Else
         KeyAscii = 0
      End If
   End If
End Sub

