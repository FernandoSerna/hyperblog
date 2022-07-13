VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmentradas_bultos_intercompañias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entradas bultos intercompañias"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2625
      Left            =   2160
      TabIndex        =   28
      Top             =   165
      Width           =   5820
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2145
         Left            =   30
         TabIndex        =   29
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
         TabIndex        =   30
         Top             =   120
         Width           =   5745
      End
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmentradas_bultos_intercompañias.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Buscar Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   405
      TabIndex        =   25
      Top             =   240
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         MaxLength       =   10
         TabIndex        =   26
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
         TabIndex        =   27
         Top             =   120
         Width           =   3060
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1095
      Picture         =   "frmentradas_bultos_intercompañias.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Lectura de cajas "
      Height          =   840
      Left            =   105
      TabIndex        =   17
      Top             =   6345
      Width           =   11505
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
         TabIndex        =   18
         Top             =   195
         Width           =   3405
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11220
      Picture         =   "frmentradas_bultos_intercompañias.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_cerrar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   765
      Picture         =   "frmentradas_bultos_intercompañias.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cerrar lectura de cajas"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmentradas_bultos_intercompañias.frx":0940
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   120
      Left            =   90
      TabIndex        =   16
      Top             =   255
      Width           =   11505
   End
   Begin VB.Frame Frame2 
      Height          =   4785
      Left            =   105
      TabIndex        =   13
      Top             =   1500
      Width           =   11490
      Begin VB.Frame frm_codigo 
         Height          =   1245
         Left            =   3180
         TabIndex        =   19
         Top             =   1590
         Width           =   4920
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
            ForeColor       =   &H00C00000&
            Height          =   510
            Left            =   765
            TabIndex        =   21
            Top             =   525
            Width           =   3405
         End
         Begin VB.Label lbl_caja 
            BackColor       =   &H8000000D&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   345
            Left            =   0
            TabIndex        =   20
            Top             =   15
            Width           =   4905
         End
      End
      Begin MSComctlLib.ListView lv_detalle_cajas 
         Height          =   4590
         Left            =   60
         TabIndex        =   11
         Top             =   135
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   8096
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
            Object.Width           =   6703
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
            Text            =   "Piezas"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Movimiento"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Costo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Precio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Estatus"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1140
      Left            =   105
      TabIndex        =   12
      Top             =   360
      Width           =   11490
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
         Left            =   1605
         TabIndex        =   5
         Top             =   210
         Width           =   5295
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
         Left            =   810
         TabIndex        =   4
         Top             =   210
         Width           =   780
      End
      Begin VB.TextBox txt_folio 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8130
         TabIndex        =   6
         Top             =   135
         Width           =   2805
      End
      Begin VB.CommandButton cmd_buscar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10995
         Picture         =   "frmentradas_bultos_intercompañias.frx":0A42
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Buscar Movimiento Alt + B"
         Top             =   630
         Width           =   330
      End
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
         Left            =   8130
         TabIndex        =   9
         Top             =   622
         Width           =   2805
      End
      Begin VB.TextBox txt_planta 
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
         Left            =   810
         TabIndex        =   7
         Top             =   622
         Width           =   780
      End
      Begin VB.TextBox txt_nombre_planta 
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
         Left            =   1605
         TabIndex        =   8
         Top             =   622
         Width           =   5295
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Left            =   135
         TabIndex        =   23
         Top             =   300
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         Height          =   195
         Left            =   7365
         TabIndex        =   22
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Factura:"
         Height          =   195
         Index           =   0
         Left            =   7365
         TabIndex        =   15
         Top             =   705
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Planta:"
         Height          =   195
         Left            =   135
         TabIndex        =   14
         Top             =   705
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmentradas_bultos_intercompañias"
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
         rs.Open "SELECT VCHA_ENT_PLANTA_PROVEEDOR, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_ENT_cAJA, A.VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ENT_CANTIDAD_CAJA, FLOA_ENT_CANTIDAD_CAJA_LEIDA, FLOA_ENT_CANTIDAD_PIEZAS_LEIDA, floa_Ent_costo, floa_ent_precio, isnull(vcha_ent_Estatus,'') as vcha_ent_estatus  FROM  TB_ENTRADAS_BULTOS_INTERCOMPAÑIAS A, TB_ARTICULOS B WHERE VCHA_ENT_PLANTA_PROVEEDOR = '" + Me.txt_planta + "' AND VCHA_SER_SERIE_ID = '" + var_serie_FACTURA + "' AND INTE_CAR_NUMERO = " + var_numero_factura + " AND A.VCHA_aRT_ARTICULO_ID = B.VCHA_ART_ARTICULO_ID", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Set list_item = Me.lv_detalle_cajas.ListItems.Add(, , rs!vcha_ent_Caja)
                  list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_Articulo_id), "", rs!vcha_Art_Articulo_id)
                  list_item.SubItems(2) = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
                  list_item.SubItems(3) = IIf(IsNull(rs!FLOA_ENT_CANTIDAD_CAJA), "", rs!FLOA_ENT_CANTIDAD_CAJA)
                  list_item.SubItems(4) = IIf(IsNull(rs!floa_ent_cantidad_caja_leida), "", rs!floa_ent_cantidad_caja_leida)
                  list_item.SubItems(5) = IIf(IsNull(rs!FLOA_ENT_CANTIDAD_PIEZAS_LEIDA), "", rs!FLOA_ENT_CANTIDAD_PIEZAS_LEIDA)
                  list_item.SubItems(6) = 0
                  list_item.SubItems(7) = IIf(IsNull(rs!floa_ent_costo), "", rs!floa_ent_costo)
                  list_item.SubItems(8) = IIf(IsNull(rs!floa_ent_precio), "", rs!floa_ent_precio)
                  list_item.SubItems(9) = IIf(IsNull(rs!vcha_Ent_Estatus), "", rs!vcha_Ent_Estatus)
                  rs.MoveNext:
            Wend
         Else
            rsaux.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + Me.txt_planta + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               If var_unidad_factura = "17" Then
                  var_conexion_facturas_ei = "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=sidtextilera;Data Source=sqlquezada2"
               Else
                  var_conexion_facturas_ei = IIf(IsNull(rsaux!vcha_uor_conexion), "", rsaux!vcha_uor_conexion)
               End If
               'MsgBox var_conexion_facturas_ei
               If var_conexion_facturas_ei <> "" Then
                  'MsgBox cnn_facturas_ei
                   If cnn_facturas_ei.State = 1 Then
                      cnn_facturas_ei.Close
                   End If
                   'MsgBox var_conexion_facturas_ei
                   
                   cnn_facturas_ei.Open var_conexion_facturas_ei
                   var_cadena = "SELECT  dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO WHERE (dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = '" + var_serie_FACTURA + "') AND (dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO = " + var_numero_factura + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'VDIP')"
                   'MsgBox cnn_facturas_ei.ConnectionString
                   rsaux1.Open var_cadena, cnn_facturas_ei, adOpenDynamic, adLockOptimistic
                   If Not rsaux1.EOF Then
                      If Me.txt_planta = "28" Then
                         var_unidad_proveedor = "04"
                      End If
                      If Me.txt_planta = "29" Then
                         var_unidad_proveedor = "01"
                      End If
                      rsaux2.Open "select * from tb_Archivo_comparacion where vcha_mov_movimiento_id = 'EP' and vcha_com_proveedor = '" + var_unidad_proveedor + "' and inte_com_numero  = " + rsaux1!vcha_Emo_referencia + " AND VCHA_COM_CAJA <> '' AND VCHA_COM_CAJA IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
                      var_equivalencias_faltantes = ""
                      While Not rsaux2.EOF
                            rsaux3.Open "select * from tb_equivalencias where vcha_equ_Codigo_equivalente = '" + rsaux2!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                            If rsaux3.EOF Then
                               If var_equivalencias_faltantes = "" Then
                                  var_equivalencias_faltantes = rsaux2!vcha_Art_Articulo_id
                               Else
                                  var_equivalencias_faltantes = var_equivalencias_faltantes + ", " + rsaux2!vcha_Art_Articulo_id
                               End If
                            End If
                            rsaux3.Close
                            rsaux2.MoveNext
                      Wend
                      If var_equivalencias_faltantes = "" Then
                         If rsaux2.RecordCount > 0 Then
                            rsaux2.MoveFirst
                         End If
                         While Not rsaux2.EOF
                               rsaux3.Open "SELECT * FROM TB_eQUIVALENCIAS WHERE VCHA_EQU_CODIGO_EQUIVALENTE = '" + rsaux2!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                               VAR_CODIGO_EQUIVALENTE = rsaux3!vcha_Art_Articulo_id
                               rsaux3.Close
                               rsaux3.Open "SELECT ISNULL(MONE_aRT_PRECIO_BASE,0) FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + VAR_CODIGO_EQUIVALENTE + "'", cnn, adOpenDynamic, adLockOptimistic
                               var_precio = IIf(IsNull(rsaux3(0).Value), 0, rsaux3(0).Value)
                               rsaux3.Close
                               var_cadena = "INSERT INTO TB_ENTRADAS_BULTOS_INTERCOMPAÑIAS (VCHA_EMP_EMPRESA_ID, VCHA_USU_USUARIO_ID, VCHA_ENT_MAQUINA, VCHA_ENT_PLANTA_PROVEEDOR, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO,                   VCHA_ENT_CAJA, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD_CAJA, FLOA_ENT_CANTIDAD_CAJA_LEIDA, FLOA_ENT_CANTIDAD_PIEZAS_LEIDA, FLOA_ENT_COSTO, FLOA_ENT_PRECIO) "
                               var_cadena = var_cadena + "                      VALUES      ('" + var_empresa + "', '" + var_clave_usuario_global + "','" + fun_NombrePc + "','" + Me.txt_planta + "', '" + var_serie_FACTURA + "'," + var_numero_factura + ",'" + rsaux2!VCHA_COM_CAJA + "','" + VAR_CODIGO_EQUIVALENTE + "', " + CStr(rsaux2!FLOA_COM_CANTIDAD_ENVIADA) + ",0,0," + CStr(IIf(IsNull(rsaux2!FLOA_COM_COSTO), 0, rsaux2!FLOA_COM_COSTO)) + "," + CStr(var_precio) + ")"
                               'MsgBox var_cadena
                               rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                               rsaux2.MoveNext
                         Wend
                         Me.lv_detalle_cajas.ListItems.Clear
                         rsaux3.Open "SELECT VCHA_ENT_PLANTA_PROVEEDOR, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_ENT_cAJA, A.VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ENT_CANTIDAD_CAJA, FLOA_ENT_CANTIDAD_CAJA_LEIDA, FLOA_ENT_CANTIDAD_PIEZAS_LEIDA, floa_Ent_costo, floa_ent_precio, isnull(vcha_ent_Estatus,'') as vcha_ent_estatus   FROM  TB_ENTRADAS_BULTOS_INTERCOMPAÑIAS A, TB_ARTICULOS B WHERE VCHA_ENT_PLANTA_PROVEEDOR = '" + Me.txt_planta + "' AND VCHA_SER_SERIE_ID = '" + var_serie_FACTURA + "' AND INTE_CAR_NUMERO = " + var_numero_factura + " AND A.VCHA_aRT_ARTICULO_ID = B.VCHA_ART_ARTICULO_ID", cnn, adOpenDynamic, adLockOptimistic
                         If Not rsaux3.EOF Then
                            While Not rsaux3.EOF
                                  Set list_item = Me.lv_detalle_cajas.ListItems.Add(, , rsaux3!vcha_ent_Caja)
                                  list_item.SubItems(1) = IIf(IsNull(rsaux3!vcha_Art_Articulo_id), "", rsaux3!vcha_Art_Articulo_id)
                                  list_item.SubItems(2) = IIf(IsNull(rsaux3!vcha_Art_nombre_español), "", rsaux3!vcha_Art_nombre_español)
                                  list_item.SubItems(3) = IIf(IsNull(rsaux3!FLOA_ENT_CANTIDAD_CAJA), "", rsaux3!FLOA_ENT_CANTIDAD_CAJA)
                                  list_item.SubItems(4) = IIf(IsNull(rsaux3!floa_ent_cantidad_caja_leida), "", rsaux3!floa_ent_cantidad_caja_leida)
                                  list_item.SubItems(5) = IIf(IsNull(rsaux3!FLOA_ENT_CANTIDAD_PIEZAS_LEIDA), "", rsaux3!FLOA_ENT_CANTIDAD_PIEZAS_LEIDA)
                                  list_item.SubItems(6) = 0
                                  list_item.SubItems(7) = IIf(IsNull(rsaux3!floa_ent_costo), "", rsaux3!floa_ent_costo)
                                  list_item.SubItems(8) = IIf(IsNull(rsaux3!floa_ent_precio), "", rsaux3!floa_ent_precio)
                                  list_item.SubItems(9) = IIf(IsNull(rsaux3!vcha_Ent_Estatus), "", rsaux3!vcha_Ent_Estatus)
                                  rsaux3.MoveNext:
                            Wend
                         End If
                         rsaux3.Close
                      Else
                         MsgBox "Los siguientes códigos no tienen equivalencia en CANTIA " + var_equivalencias_faltantes, vbOKOnly, "ATENCION"
                      End If
                      rsaux2.Close
                   Else
                      MsgBox "La factura no existe", vbOKOnly, "ATENCION"
                   End If
                   rsaux1.Close
                Else
                   MsgBox "La planta no cuenta con conexión", vbOKOnly, "ATENCION"
                End If
            Else
               MsgBox "La planta no existe", vbOKOnly, "ATENCION"
            End If
            'MsgBox "La factura no existe", vbOKOnly, "ATENCION"
            rsaux.Close
            'AQUI DEBE DE IR LA BUSQUEDA DE LA NOTA
         End If
         rs.Close
      Else
         MsgBox "Factura incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a indicado ninguna factura", vbOKOnly, "ATENCION"
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
            rs.Open "SELECT * FROM  TB_ENTRADAS_BULTOS_INTERCOMPAÑIAS  WHERE VCHA_ENT_PLANTA_PROVEEDOR = '" + Me.txt_planta + "' AND VCHA_SER_SERIE_ID = '" + var_serie_FACTURA + "' AND INTE_CAR_NUMERO = " + var_numero_factura + " ", cnn, adOpenDynamic, adLockOptimistic
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
                           rsaux.Open "update TB_ENTRADAS_BULTOS_INTERCOMPAÑIAS  set VCHA_ent_estatus = 'I', DTIM_ENT_FECHA = GETDATE()  WHERE VCHA_ENT_PLANTA_PROVEEDOR = '" + Me.txt_planta + "' AND VCHA_SER_SERIE_ID = '" + var_serie_FACTURA + "' AND INTE_CAR_NUMERO = " + var_numero_factura + " ", cnn, adOpenDynamic, adLockOptimistic
                           For var_j = 1 To Me.lv_detalle_cajas.ListItems.Count
                               Me.lv_detalle_cajas.ListItems.item(var_j).Selected = True
                               Me.lv_detalle_cajas.selectedItem.SubItems(9) = "I"
                           Next var_j
                           var_numero_factura = CStr(CDbl(var_numero_factura))
                           Set reporte = appl.OpenReport(App.Path + "\REP_BULTOS_ENTRADAS.rpt")
                           reporte.RecordSelectionFormula = "{VW_BULTOS_ENTRADAS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_BULTOS_ENTRADAS.VCHA_ENT_PLANTA_PROVEEDOR} = '" + Me.txt_planta + "' AND {VW_BULTOS_ENTRADAS.VCHA_SER_SERIE_ID} = '" + var_serie_FACTURA + "' AND {VW_BULTOS_ENTRADAS.INTE_CAR_NUMERO} = '" + var_numero_factura + "'"
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
                  reporte.RecordSelectionFormula = "{VW_BULTOS_ENTRADAS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_BULTOS_ENTRADAS.VCHA_ENT_PLANTA_PROVEEDOR} = '" + Me.txt_planta + "' AND {VW_BULTOS_ENTRADAS.VCHA_SER_SERIE_ID} = '" + var_serie_FACTURA + "' AND {VW_BULTOS_ENTRADAS.INTE_CAR_NUMERO} = '" + var_numero_factura + "'"
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
   If IsNumeric(Me.txt_folio) Then
      If var_estatus_movimiento = "I" Then
         Set reporte = appl.OpenReport(App.Path + "\REP_ENTRADAS_EBI.rpt")
         reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_ENTRADA.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_MOVIMIENTOS_ENTRADA.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' AND {VW_MOVIMIENTOS_ENTRADA.VCHA_MOV_MOVIMIENTO_ID} = 'EBI' AND {VW_MOVIMIENTOS_ENTRADA.INTE_eMO_NUMERO} = " + Me.txt_folio
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
                  var_posible_x = 1
                  cnn.BeginTrans
                  If Not rs.EOF Then
                     var_inserta = False
                     var_posible_x = 1
                     If var_posible_x = 1 Then
                        var_inserta = TB_ENTRADAS_I.Anadir(CStr(var_empresa), CStr(var_unidad_organizacional), CStr(var_almacen_Destino), CStr(var_clave_movimiento), CDbl(var_numero_folio), 1, CStr(var_almacen_origen), 0)
                     End If
                  End If
                  rs.Close
                  var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(CStr(var_empresa), CStr(var_unidad_organizacional), CStr(var_almacen_Destino), CStr(var_clave_movimiento), CDbl(var_numero_folio), "", Now, 1)
                  var_estatus_movimiento = "I"
                  var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(CStr(var_empresa), CStr(var_unidad_organizacional), CStr(var_almacen_Destino), CStr(var_clave_movimiento), CDbl(var_numero_folio), "I", Now, 1)
                  cnn.CommitTrans
                  
                  
                  If var_clave_movimiento = "EBI" Then
                     var_cadena = "select sum(floa_ent_costo * floa_Ent_Cantidad) as costo, sum(floa_ent_precio *  floa_ent_precio) as precio from tb_temporal_entradas where  vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
                     rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rsaux11.Open "select * from tb_generador_polizas where empresa_id = '" + var_empresa + "'", cnnoracle, adOpenDynamic, adLockOptimistic
                     While Not rsaux11.EOF
                           var_tipo_poliza = rsaux11!tipo
                           var_origen_poliza = rsaux11!Origen
                           var_categoria_poliza = rsaux11!categoria
                           var_moneda_poliza = rsaux11!moneda
                           var_segmento1_poliza = rsaux11!segmento1
                           var_segmento2_poliza = rsaux11!segmento2
                           var_segmento3_poliza = rsaux11!segmento3
                           var_segmento4_poliza = rsaux11!segmento4
                           var_segmento5_poliza = rsaux11!segmento5
                           var_segmento6_poliza = rsaux11!segmento6
                           var_segmento7_poliza = rsaux11!segmento7
                           var_juego_libros_poliza = rsaux11!juego_libros
                           var_descripcion_poliza = rsaux11!descripcion
                           var_cargo_poliza = rsaux11!cargo
                           var_abono_poliza = rsaux11!abono
                           var_precio = rsaux11!Precio
                           If var_precio = 1 Then
                              var_importe_precio = rsaux10!Precio
                           Else
                              var_importe_precio = rsaux10!Costo
                           End If
                           var_cadena = "InsERT INTO IN_TB_POLIZAS_INT (STATUS, SET_OF_BOOKS_ID, USER_JE_SOURCE_NAME, USER_JE_CATEGORY_NAME, ACCOUNTING_DATE, CURRENCY_CODE, DATE_CREATED, ACTUAL_FLAG,  SEGMENT1, SEGMENT2, SEGMENT3, SEGMENT4, SEGMENT5, SEGMENT6, SEGMENT7, ENTERED_DR, ENTERED_CR, ACCOUNTED_DR, ACCOUNTED_CR, GROUP_ID, REFERENCE4, REFERENCE5, REFERENCe10, REFERENCE1, REFERENCE2, CREATED_BY)"
                           If var_cargo_poliza = 1 Then
                              var_cadena = var_cadena + " VALUES ('NEW', " + CStr(var_juego_libros_poliza) + ",'" + var_origen_poliza + "','" + var_categoria_poliza + "',TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'" + var_moneda_poliza + "',TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'A','" + var_segmento1_poliza + "','" + var_segmento2_poliza + "','" + var_segmento3_poliza + "','" + var_segmento4_poliza + "','" + var_segmento5_poliza + "','" + var_segmento6_poliza + "','" + var_segmento7_poliza + "'," + CStr(var_importe_precio) + ",0," + CStr(var_importe_precio) + ",0,1,'FACTURAS INTERCOMPAÑIA " + Me.txt_facturas + "','FACTURA NUM: " + Me.txt_facturas + "','" + var_descripcion_poliza + "','POLIZA FACTURAS INTERCOMPAÑIA','POLIZA FACTURAS INTERCOMPAÑIA',1143)"
                           Else
                              var_cadena = var_cadena + " VALUES ('NEW', " + CStr(var_juego_libros_poliza) + ",'" + var_origen_poliza + "','" + var_categoria_poliza + "',TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'" + var_moneda_poliza + "',TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'A','" + var_segmento1_poliza + "','" + var_segmento2_poliza + "','" + var_segmento3_poliza + "','" + var_segmento4_poliza + "','" + var_segmento5_poliza + "','" + var_segmento6_poliza + "','" + var_segmento7_poliza + "',0," + CStr(var_importe_precio) + ",0," + CStr(var_importe_precio) + ",1,'FACTURAS INTERCOMPAÑIA " + Me.txt_facturas + "','FACTURA NUM: " + Me.txt_facturas + "','" + var_descripcion_poliza + "','POLIZA FACTURAS INTERCOMPAÑIA','POLIZA FACTURAS INTERCOMPAÑIA',1143)"
                           End If
                           'MsgBox var_cadena
                           rsaux9.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                           rsaux11.MoveNext
                     Wend
                     rsaux11.Close
                     
                     rsaux11.Open "select sq_id_facturas.nextval from dual", cnnoracle, adOpenDynamic, adLockOptimistic
                     var_consecutivo = rsaux11(0).Value
                     rsaux11.Close
                     If Me.txt_planta = "28" Then
                        var_proveedor_oracle = "04"
                     End If
                     If Me.txt_planta = "29" Then
                        var_proveedor_oracle = "01"
                     End If
                     'MsgBox "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_unidad_id = '" + var_proveedor_oracle + "'"
                     rsaux11.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_unidad_id = '" + var_proveedor_oracle + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_empresa_emite = rsaux11!VCHA_EMP_EMPRESA_ID
                     var_proveedor_oracle_2 = rsaux11!vcha_uor_proveedor_oracle
                     rsaux11.Close
                     'MsgBox "select * from tb_empresas_cruzadas_oracle where vcha_emp_Empresa_emite = '" + var_empresa_emite + "' and vcha_emp_empresa_recibe = '" + var_empresa + "'"
                     rsaux11.Open "select * from tb_empresas_cruzadas_oracle where vcha_emp_Empresa_emite = '" + var_empresa_emite + "' and vcha_emp_empresa_recibe = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_unidad_oracle = rsaux11!vcha_emp_organizacion
                     rsaux11.Close
                     'MsgBox "SELECT vendor_site_id FROM po_vendor_sites_all@perpvia.vianney.com.mx Where vendor_id = '" + var_proveedor_oracle_2 + "'  AND vendor_site_id in (4070,2803,1125,1200,1202,1126,1519,1327,1520,3545,1127,1674,4383,2668,1529,1326,2669,2925,3755,1268,1324,3768,2392,9016,3332,1462, 1473,1737,1992, 1991,4618,5454,6407,4380,10626) AND ORG_ID = '" + var_unidad_oracle + "'"
                     rsaux11.Open "SELECT vendor_site_id FROM po_vendor_sites_all@perpvia.vianney.com.mx Where vendor_id = '" + var_proveedor_oracle_2 + "'  AND vendor_site_id in (4070,2803,1125,1200,1202,1126,1519,1327,1520,3545,1127,1674,4383,2668,1529,1326,2669,2925,3755,1268,1324,3768,2392,9016,3332,1462, 1473,1737,1992, 1991,4618,5454,6407,4380,10626,9899, 7244) AND ORG_ID = '" + var_unidad_oracle + "'", cnnoracle, adOpenDynamic, adLockOptimistic
                     var_clave_proveedor_oracle = rsaux11!vendor_site_id
                     rsaux11.Close
                     
                     var_cadena = "select sum(floa_ent_costo * floa_Ent_Cantidad) as costo, sum(floa_ent_precio *  floa_ent_precio) as precio from tb_temporal_entradas where  vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
                     rsaux11.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     var_importe_total = rsaux11(0).Value
                     rsaux11.Close
                  
                     var_importe_total = var_importe_total * 1.16
                     
                     var_cadena = "insert into IN_TB_FACTURAS_INT (INVOICE_ID,INVOICE_NUM,INVOICE_TYPE_LOOKUP_CODE,VENDOR_ID,VENDOR_SITE_ID,INVOICE_AMOUNT,INVOICE_CURRENCY_CODE,EXCHANGE_RATE_TYPE,EXCHANGE_DATE,EXCHANGE_RATE,Description,Source,GL_DATE,INVOICE_DATE,ORG_ID) values (" + CStr(var_consecutivo) + ",'" + Me.txt_facturas + "','STANDARD'," + CStr(var_proveedor_oracle_2) + "," + CStr(var_clave_proveedor_oracle) + "," + CStr(var_importe_total) + ",'MXP',null,null,null,'FACTURA DE RECEPCION NUM: " + Me.txt_facturas + "','FACTURA INTERCOMPAÑIAS',TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),TO_DATE('" + CStr(Date) + "','DD/MM/YYYY')," + var_unidad_oracle + ")"
                     rsaux11.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                     
                     rsaux11.Open "select sq_id_lineas_factura.nextval from dual", cnnoracle, adOpenDynamic, adLockOptimistic
                     var_consecutivo_linea = rsaux11(0).Value
                     rsaux11.Close
                     var_subimporte = var_importe_total / 1.16
                     var_importe_iva = var_importe_total - var_subimporte
                     rsaux11.Open "select amount_includes_tax_flag, vat_code from po_vendor_sites_all@perpvia.vianney.com.mx Where vendor_id = " + CStr(var_proveedor_oracle_2) + " and vendor_site_id = " + CStr(var_clave_proveedor_oracle) + " and org_id = " + CStr(var_unidad_oracle), cnnoracle, adOpenDynamic, adLockOptimistic
                     amount_includes_tax_flag = rsaux11!amount_includes_tax_flag
                     TAX_CODE = IIf(IsNull(rsaux11!vat_code), 0, rsaux11!vat_code)
                     rsaux11.Close
                     rsaux.Open "select awt_group_id from po_vendors@perpvia.vianney.com.mx Where vendor_id = " + CStr(var_proveedor_oracle), cnnoracle, adOpenDynamic, adLockOptimistic
                     AWT_GROUP_ID = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                     rsaux.Close
                     'MsgBox CStr(AWT_GROUP_ID)
                     If TAX_CODE = 0 Then
                        If AWT_GROUP_ID = 0 Then
                           var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",1,'ITEM'," + CStr(var_subimporte) + ", TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                        Else
                           var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",1,'ITEM'," + CStr(var_subimporte) + ", TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "',NULL,NULL," + CStr(AWT_GROUP_ID) + ",NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                        End If
                     Else
                        If AWT_GROUP_ID = 0 Then
                           var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",1,'ITEM'," + CStr(var_subimporte) + ", TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "','" + CStr(TAX_CODE) + "',NULL,NULL,NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                        Else
                           var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",1,'ITEM'," + CStr(var_subimporte) + ", TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "','" + CStr(TAX_CODE) + "',NULL," + CStr(AWT_GROUP_ID) + ",NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                        End If
                     End If
                     
                     'MsgBox var_cadena
                     rsaux11.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                     rsaux11.Open "select sq_id_lineas_factura.nextval from dual", cnnoracle, adOpenDynamic, adLockOptimistic
                     var_consecutivo_linea = rsaux11(0).Value
                     rsaux11.Close
                     
                     var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description,AMOUNT_INCLUDES_TAX_FLAG,TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",2,'TAX'," + CStr(var_importe_iva) + ", TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'IMPUESTO',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                     rsaux11.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                  End If
                  
                  Set reporte = appl.OpenReport(App.Path + "\REP_ENTRADAS_EBI.rpt")
                  reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_ENTRADA.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_MOVIMIENTOS_ENTRADA.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' AND {VW_MOVIMIENTOS_ENTRADA.VCHA_MOV_MOVIMIENTO_ID} = 'EBI' AND {VW_MOVIMIENTOS_ENTRADA.INTE_eMO_NUMERO} = " + Me.txt_folio
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
         End If
      End If
   End If
End Sub

Private Sub cmd_nuevo_Click()
  var_primera_vez = 1
  Me.txt_almacen = ""
  Me.txt_nombre_almacen = ""
  Me.txt_facturas = ""
  Me.txt_planta = ""
  Me.txt_nombre_planta = ""
  Me.lv_detalle_cajas.ListItems.Clear
  Me.txt_almacen.Enabled = True
  Me.txt_planta.Enabled = True
  Me.txt_facturas.Enabled = True
  Me.txt_almacen.SetFocus
  var_estatus_movimiento = ""
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
   Me.frm_codigo.Visible = False
   var_primera_vez = 1
   Me.frm_busqueda.Visible = False
   Me.frm_lista.Visible = False
   var_estatus_movimiento = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_entradas_sin_comparacion)
End Sub

Private Sub lv_detalle_cajas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Me.lv_detalle_cajas.selectedItem.SubItems(9) = "I" Then
         Me.lbl_caja = Me.lv_detalle_cajas.selectedItem
         Me.frm_codigo.Visible = True
         Me.txt_codigo = ""
         Me.txt_codigo.SetFocus
      Else
         MsgBox "No se a cerrado la lectura de cajas", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If var_tipo_lista = 1 Then
      If Me.lv_lista.ListItems.Count > 0 Then
         Me.txt_almacen = lv_lista.selectedItem
         Me.txt_nombre_almacen = Me.lv_lista.selectedItem.SubItems(1)
      End If
      Me.txt_almacen.SetFocus
   End If
   
   If var_tipo_lista = 2 Then
      If Me.lv_lista.ListItems.Count > 0 Then
         Me.txt_planta = Me.lv_lista.selectedItem
         Me.txt_nombre_planta = Me.lv_lista.selectedItem.SubItems(1)
      End If
      Me.txt_planta.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "SELECT dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_MOVIMIENTOS_ALMACENES.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ALMACENES.VCHA_EMP_EMPRESA_ID FROM dbo.TB_MOVIMIENTOS_ALMACENES INNER JOIN dbo.TB_ALMACENES ON dbo.TB_MOVIMIENTOS_ALMACENES.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID WHERE  (dbo.TB_ALMACENES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_MOVIMIENTOS_ALMACENES.VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "')", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Almacenes"
      var_tipo_lista = 1
      Dim var_n As Integer
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_almacen_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.txt_planta.SetFocus
   End If
End Sub

Private Sub txt_almacen_LostFocus()
   If Trim(Me.txt_almacen) <> "" Then
      rs.Open "SELECT dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_MOVIMIENTOS_ALMACENES.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ALMACENES.VCHA_EMP_EMPRESA_ID FROM dbo.TB_MOVIMIENTOS_ALMACENES INNER JOIN dbo.TB_ALMACENES ON dbo.TB_MOVIMIENTOS_ALMACENES.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID WHERE (dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID = '" + Me.txt_almacen + "') AND (dbo.TB_ALMACENES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_MOVIMIENTOS_ALMACENES.VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "')", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_almacen = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
      Else
         MsgBox "El almacén no existe o no tiene permisos para este movimiento", vbOKOnly, "ATENCION"
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
            If rs!VCHA_PRO_PROVEEDOR_ID = "04" Then
               Me.txt_planta = "28"
            End If
            If rs!VCHA_PRO_PROVEEDOR_ID = "01" Then
               Me.txt_planta = "29"
            End If
            
            rsaux.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + Me.txt_planta + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_nombre_planta = IIf(IsNull(rsaux!VCHA_UOR_NOMBRE), "", rsaux!VCHA_UOR_NOMBRE)
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
            If IsNumeric(var_numero_factura) Then
               Me.lv_detalle_cajas.ListItems.Clear
               rsaux.Open "SELECT VCHA_ENT_PLANTA_PROVEEDOR, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_ENT_cAJA, A.VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ENT_CANTIDAD_CAJA, FLOA_ENT_CANTIDAD_CAJA_LEIDA, FLOA_ENT_CANTIDAD_PIEZAS_LEIDA, floa_Ent_costo, floa_ent_precio, isnull(vcha_ent_Estatus,'') as vcha_ent_estatus  FROM  TB_ENTRADAS_BULTOS_INTERCOMPAÑIAS A, TB_ARTICULOS B WHERE VCHA_ENT_PLANTA_PROVEEDOR = '" + Me.txt_planta + "' AND VCHA_SER_SERIE_ID = '" + var_serie_FACTURA + "' AND INTE_CAR_NUMERO = " + var_numero_factura + " AND A.VCHA_aRT_ARTICULO_ID = B.VCHA_ART_ARTICULO_ID", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  While Not rsaux.EOF
                        Set list_item = Me.lv_detalle_cajas.ListItems.Add(, , rsaux!vcha_ent_Caja)
                        list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_Art_Articulo_id), "", rsaux!vcha_Art_Articulo_id)
                        list_item.SubItems(2) = IIf(IsNull(rsaux!vcha_Art_nombre_español), "", rsaux!vcha_Art_nombre_español)
                        list_item.SubItems(3) = IIf(IsNull(rsaux!FLOA_ENT_CANTIDAD_CAJA), "", rsaux!FLOA_ENT_CANTIDAD_CAJA)
                        list_item.SubItems(4) = IIf(IsNull(rsaux!floa_ent_cantidad_caja_leida), "", rsaux!floa_ent_cantidad_caja_leida)
                        list_item.SubItems(5) = IIf(IsNull(rsaux!FLOA_ENT_CANTIDAD_PIEZAS_LEIDA), "", rsaux!FLOA_ENT_CANTIDAD_PIEZAS_LEIDA)
                        list_item.SubItems(6) = 0
                        list_item.SubItems(7) = IIf(IsNull(rsaux!floa_ent_costo), "", rsaux!floa_ent_costo)
                        list_item.SubItems(8) = IIf(IsNull(rsaux!floa_ent_precio), "", rsaux!floa_ent_precio)
                        list_item.SubItems(9) = IIf(IsNull(rsaux!vcha_Ent_Estatus), "", rsaux!vcha_Ent_Estatus)
                        rsaux.MoveNext:
                  Wend
               End If
               rsaux.Close
               rsaux.Open "select * from tb_temporal_Entradas where  vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + Me.txt_busqueda_folio, cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux.EOF
                     For var_j = 1 To Me.lv_detalle_cajas.ListItems.Count
                         Me.lv_detalle_cajas.ListItems.item(var_j).Selected = True
                         If rsaux!vcha_Art_Articulo_id = Me.lv_detalle_cajas.selectedItem.SubItems(1) And rsaux!vcha_ent_Caja = Me.lv_detalle_cajas.selectedItem Then
                            Me.lv_detalle_cajas.selectedItem.SubItems(6) = Me.lv_detalle_cajas.selectedItem.SubItems(6) + rsaux!floa_ent_Cantidad
                         End If
                     Next var_j
                     rsaux.MoveNext
               Wend
               rsaux.Close
            End If
            
            
            
            Me.frm_busqueda.Visible = False
            Me.txt_almacen.Enabled = False
            Me.txt_nombre_almacen.Enabled = False
            Me.txt_planta.Enabled = False
            Me.txt_nombre_planta.Enabled = False
            Me.txt_facturas.Enabled = False
            var_primera_vez = 0
         Else
            MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
            Me.txt_almacen = ""
            Me.txt_folio = ""
            Me.txt_nombre_almacen = ""
            Me.txt_planta = ""
            Me.txt_nombre_planta = ""
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
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Me.lv_detalle_cajas.ListItems.Count > 0 Then
         If Mid(Me.txt_caja, 1, 2) = "CA" Then
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
               rs.Open "SELECT * FROM TB_ENTRADAS_BULTOS_INTERCOMPAÑIAS WHERE VCHA_ENT_PLANTA_PROVEEDOR = '" + Me.txt_planta + "' AND VCHA_SER_sERIE_ID = '" + var_serie_FACTURA + "' AND INTE_cAR_NUMERO = " + var_numero_factura + " AND VCHA_ENT_CAJA = '" + Me.txt_caja + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_cantidad_caja = IIf(IsNull(rs!floa_ent_cantidad_caja_leida), 0, rs!floa_ent_cantidad_caja_leida)
                  If var_cantidad_caja > 0 Then
                     Me.txt_caja = ""
                     frmmensaje.lbl_mensaje = "La caja ya fue leida"
                     frmmensaje.Show
                  Else
                     For var_j = 1 To Me.lv_detalle_cajas.ListItems.Count
                         Me.lv_detalle_cajas.ListItems.item(var_j).Selected = True
                         If Me.lv_detalle_cajas.selectedItem = Me.txt_caja Then
                            rsaux.Open "update TB_ENTRADAS_BULTOS_INTERCOMPAÑIAS set floa_ent_Cantidad_caja_leida = floa_ent_cantidad_caja where  VCHA_ENT_PLANTA_PROVEEDOR = '" + Me.txt_planta + "' AND VCHA_SER_sERIE_ID = '" + var_serie_FACTURA + "' AND INTE_cAR_NUMERO = " + var_numero_factura + " AND VCHA_ENT_CAJA = '" + Me.txt_caja + "'", cnn, adOpenDynamic, adLockOptimistic
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
      If Me.txt_codigo <> "" Then
         rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            rsaux.Open "SELECT * FROM TB_EQUIVALENCIAS WHERE VCHA_EQU_CODIGO_EQUIVALENTE = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               rsaux1.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  Me.txt_codigo = rsaux1!vcha_Art_Articulo_id
               End If
               rsaux1.Close
            End If
            rsaux.Close
         End If
         rs.Close
         If Me.txt_codigo = Me.lv_detalle_cajas.selectedItem.SubItems(1) Then
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
               If Me.txt_planta = "28" Then
                  var_proveedor = "04"
               End If
               If IsNumeric(var_numero_factura) Then
                  If var_primera_vez = 1 Then
                     var_numero_folio = 0
                     var_folio_enviado = 0
                     var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, CStr(var_almacen_Destino), var_clave_movimiento, Now, CDbl(var_numero_folio), CDbl(var_folio_enviado), "", CStr(var_proveedor), CStr(var_almacen_origen), CStr(var_almacen_Destino), "", var_clave_usuario_global, fun_NombrePc, CStr(var_factura), "", CStr(Me.txt_facturas), "", "B", "", "", 0, 0, 0, CStr(var_clave_moneda), CDbl(var_tipo_Cambio))
                     var_numero_folio = var_numero_folio_regreso
                     Me.txt_folio = var_numero_folio
                     var_primera_vez = 0
                  End If
                  rs.Open "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(Me.txt_folio) + " and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     rsaux.Open "update tb_temporal_Entradas set floa_ent_cantidad = floa_ent_cantidad + 1 where vcha_emp_empresa_id  = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + CStr(Me.txt_folio) + " and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     var_costo = CDbl(Me.lv_detalle_cajas.selectedItem.SubItems(7))
                     var_precio = CDbl(Me.lv_detalle_cajas.selectedItem.SubItems(8))
                     rsaux.Open "insert into tb_temporal_Entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_precio, floa_ent_costo, floa_ent_Cantidad, vcha_ent_Caja) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "'," + CStr(Me.txt_folio) + ",'" + Me.txt_codigo + "', " + CStr(var_precio) + "," + CStr(var_costo) + ", 1,'" + Me.lv_detalle_cajas.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rs.Close
                  rsaux.Open "update TB_ENTRADAS_BULTOS_INTERCOMPAÑIAS set floa_Ent_Cantidad_piezas_leida = isnull(floa_Ent_Cantidad_piezas_leida,0) + 1 where  VCHA_ENT_PLANTA_PROVEEDOR = '" + Me.txt_planta + "' AND VCHA_SER_sERIE_ID = '" + var_serie_FACTURA + "' AND INTE_cAR_NUMERO = " + var_numero_factura + " AND VCHA_ENT_CAJA = '" + Me.lbl_caja + "' AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  Me.lv_detalle_cajas.selectedItem.SubItems(5) = CDbl(Me.lv_detalle_cajas.selectedItem.SubItems(5)) + 1
                  Me.lv_detalle_cajas.selectedItem.SubItems(6) = CDbl(Me.lv_detalle_cajas.selectedItem.SubItems(6)) + 1
                  Me.txt_codigo = ""
               Else
               End If
            Else
               frmmensaje.lbl_mensaje = "La cantidad excede a la cantidad que viene en la caja"
               frmmensaje.Show 1
               Me.txt_caja = ""
            End If
         Else
            rs.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
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

Private Sub txt_codigo_LostFocus()
   Me.frm_codigo.Visible = False
End Sub

Private Sub txt_facturas_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.cmd_buscar.SetFocus
   End If
End Sub

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_planta.SetFocus
   Else
      If KeyAscii = 27 Then
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_nombre_planta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_facturas.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_planta_Change()
   Me.txt_nombre_planta = ""
   Me.lv_detalle_cajas.ListItems.Clear
   Me.txt_folio = ""
   Me.txt_facturas = ""
End Sub

Private Sub txt_planta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE  VCHA_EMP_EMPRESA_ID= '06'", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_UOR_UNIDAD_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_UOR_NOMBRE), "", rs!VCHA_UOR_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Plantas"
      var_tipo_lista = 2
      Dim var_n As Integer
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_planta_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.txt_nombre_planta.SetFocus
   End If
End Sub

Private Sub txt_planta_LostFocus()
   If Trim(Me.txt_planta) <> "" Then
      rs.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + Me.txt_planta + "' AND VCHA_EMP_EMPRESA_ID= '06'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_planta = IIf(IsNull(rs!VCHA_UOR_NOMBRE), "", rs!VCHA_UOR_NOMBRE)
         Me.txt_facturas.Enabled = True
         Me.txt_facturas.SetFocus
      Else
         MsgBox "Clave de planta incorrecta", vbOKOnly, "ATENCION"
         Me.txt_facturas = ""
         Me.txt_nombre_planta = ""
         Me.lv_detalle_cajas.ListItems.Clear
         Me.txt_facturas.Enabled = False
      End If
      rs.Close
   End If
End Sub
