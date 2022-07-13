VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsalidas_apartados_cantia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salida para almacén de apartados pendientes por entregar"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_cargar_movimientos 
      Caption         =   "Command1"
      Height          =   270
      Left            =   2190
      TabIndex        =   10
      Top             =   15
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmsalidas_apartados_cantia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Movimiento Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11130
      Picture         =   "frmsalidas_apartados_cantia.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   75
      Left            =   30
      TabIndex        =   8
      Top             =   330
      Width           =   11535
   End
   Begin VB.Frame Frame2 
      Height          =   3390
      Left            =   105
      TabIndex        =   6
      Top             =   3810
      Width           =   11430
      Begin VB.Frame frm_cantidad_eliminar_apartada_entregar 
         Height          =   840
         Left            =   5565
         TabIndex        =   20
         Top             =   2160
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar_apartada_eliminar 
            Height          =   330
            Left            =   75
            TabIndex        =   21
            Top             =   375
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   "Cantidad a eliminar"
            ForeColor       =   &H8000000E&
            Height          =   225
            Index           =   4
            Left            =   0
            TabIndex        =   22
            Top             =   15
            Width           =   2895
         End
      End
      Begin VB.Frame frm_cantidad_apartada_entregar 
         Height          =   840
         Left            =   5580
         TabIndex        =   17
         Top             =   1260
         Width           =   2910
         Begin VB.TextBox txt_cantidad_apartada_entregar 
            Height          =   330
            Left            =   75
            TabIndex        =   18
            Top             =   375
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   "Cantidad"
            ForeColor       =   &H8000000E&
            Height          =   225
            Index           =   2
            Left            =   0
            TabIndex        =   19
            Top             =   15
            Width           =   2895
         End
      End
      Begin MSComctlLib.ListView lv_almacen_apartado_entragar 
         Height          =   2805
         Left            =   45
         TabIndex        =   1
         Top             =   525
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   4948
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
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Posibles"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Existen"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Pasar"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Costo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "precio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "almacen"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Almacén apartados pendientes por entregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   0
         Left            =   45
         TabIndex        =   7
         Top             =   135
         Width           =   11325
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3390
      Left            =   105
      TabIndex        =   4
      Top             =   390
      Width           =   11430
      Begin VB.Frame frm_cantidad_eliminar_apartada 
         Height          =   840
         Left            =   5250
         TabIndex        =   14
         Top             =   2085
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar_apartada 
            Height          =   330
            Left            =   75
            TabIndex        =   15
            Top             =   375
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   "Cantidad a eliminar"
            ForeColor       =   &H8000000E&
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   16
            Top             =   15
            Width           =   2895
         End
      End
      Begin VB.Frame frm_cantidad_apartada 
         Height          =   840
         Left            =   5265
         TabIndex        =   11
         Top             =   1185
         Width           =   2910
         Begin VB.TextBox txt_cantidad_apartada 
            Height          =   330
            Left            =   75
            TabIndex        =   12
            Top             =   375
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   "Cantidad"
            ForeColor       =   &H8000000E&
            Height          =   225
            Index           =   3
            Left            =   0
            TabIndex        =   13
            Top             =   15
            Width           =   2895
         End
      End
      Begin MSComctlLib.ListView lv_almacen_apartado 
         Height          =   2805
         Left            =   45
         TabIndex        =   0
         Top             =   525
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   4948
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
            Text            =   "Código"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Posibles"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Existen"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Pasar"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Costo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "precio"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   " Almacén apartados pendientes "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   45
         TabIndex        =   5
         Top             =   135
         Width           =   11340
      End
   End
   Begin VB.Label lbl_consecutivo 
      Height          =   345
      Left            =   1635
      TabIndex        =   9
      Top             =   -15
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmsalidas_apartados_cantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_cargar_movimientos_Click()
   var_cadena = "SELECT dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.INTE_TEM_CONSECUTIVO, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_TEM_ALMACEN_APARTADO, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_TEM_ALMACEN_APARTADO_ENTREGAR, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_EMP_EMPRESA_ID, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_UOR_UNIDAD_ID, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_ALM_ALMACEN_ID, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.INTE_ENT_NUMERO, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_ART_ARTICULO_ID, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.FLOA_ENT_CANTIDAD, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.FLOA_ENT_COSTO, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.FLOA_ENT_PRECIO, "
   var_cadena = var_cadena + "dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.FLOA_ENT_CANTIDAD_ALMACEN_APARTADO, "
   var_cadena = var_cadena + " dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR, dbo.TB_Articulos.VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_PASAR, FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR FROM dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_ART_ARTICULO_ID = dbo.TB_Articulos.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_TEM_ALMACEN_APARTADO = 'ALAPP')  AND inte_ent_numero = " + CStr(var_consecutivo_apartados_Cantia) + " AND FLOA_ENT_CANTIDAD_ALMACEN_APARTADO < 0 and vcha_mov_movimiento_id = '" + var_clave_movimiento_apartados_Cantia + "'"
   'MsgBox var_cadena
   rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
         Set list_item = Me.lv_almacen_apartado.ListItems.Add(, , rsaux!vcha_Art_Articulo_id)
         list_item.SubItems(1) = rsaux!vcha_Art_nombre_español
         list_item.SubItems(2) = Format(IIf(IsNull(rsaux!floa_ent_Cantidad), 0, rsaux!floa_ent_Cantidad), "###,###,##0.00")
         list_item.SubItems(3) = Format(IIf(IsNull(rsaux!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO), 0, rsaux!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO) * (0 - 1), "###,###,##0.00")
         list_item.SubItems(4) = Format(IIf(IsNull(rsaux!floa_Ent_Cantidad_almacen_apartado_pasar), 0, rsaux!floa_Ent_Cantidad_almacen_apartado_pasar), "###,###,##0.00")
         list_item.SubItems(5) = Format(IIf(IsNull(rsaux!floa_ent_costo), 0, rsaux!floa_ent_costo), "###,###,##0.00")
         list_item.SubItems(6) = Format(IIf(IsNull(rsaux!floa_ent_precio), 0, rsaux!floa_ent_precio), "###,###,##0.00")
         rsaux.MoveNext
   Wend
   rsaux.Close
   var_cadena = "SELECT dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.INTE_TEM_CONSECUTIVO, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_TEM_ALMACEN_APARTADO, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_TEM_ALMACEN_APARTADO_ENTREGAR, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_EMP_EMPRESA_ID, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_UOR_UNIDAD_ID, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_ALM_ALMACEN_ID, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.INTE_ENT_NUMERO, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_ART_ARTICULO_ID, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.FLOA_ENT_CANTIDAD, "
   var_cadena = var_cadena + " dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.FLOA_ENT_COSTO, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.FLOA_ENT_PRECIO, dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.FLOA_ENT_CANTIDAD_ALMACEN_APARTADO, "
   var_cadena = var_cadena + " dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR, dbo.TB_Articulos.VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_PASAR, FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR  FROM dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_ART_ARTICULO_ID = dbo.TB_Articulos.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA.VCHA_TEM_ALMACEN_APARTADO_ENTREGAR = 'ALAPPE')  AND inte_Ent_numero = " + CStr(var_consecutivo_apartados_Cantia) + " AND FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR < 0 and vcha_mov_movimiento_id = '" + var_clave_movimiento_apartados_Cantia + "'"
   rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
         Set list_item = Me.lv_almacen_apartado_entragar.ListItems.Add(, , rsaux!vcha_Art_Articulo_id)
         list_item.SubItems(1) = rsaux!vcha_Art_nombre_español
         list_item.SubItems(2) = Format(IIf(IsNull(rsaux!floa_ent_Cantidad), 0, rsaux!floa_ent_Cantidad), "###,###,##0.00")
         list_item.SubItems(3) = Format(IIf(IsNull(rsaux!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR), 0, rsaux!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR) * (0 - 1), "###,###,##0.00")
         list_item.SubItems(4) = Format(IIf(IsNull(rsaux!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR), 0, rsaux!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR), "###,###,##0.00")
         list_item.SubItems(5) = Format(IIf(IsNull(rsaux!floa_ent_costo), 0, rsaux!floa_ent_costo), "###,###,##0.00")
         list_item.SubItems(6) = Format(IIf(IsNull(rsaux!floa_ent_precio), 0, rsaux!floa_ent_precio), "###,###,##0.00")
         list_item.SubItems(7) = IIf(IsNull(rsaux!VCHA_ALM_ALMACEN_ID), "", rsaux!VCHA_ALM_ALMACEN_ID)
         rsaux.MoveNext
   Wend
   rsaux.Close
   
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
   Dim var_almacen_Destino_traspaso As String
   Dim var_almacen_origen_traspaso As String
   Dim var_clave_movimiento_traspaso As String
   Dim var_numero_folio_traspaso As Double
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
   
   rsaux.Open "UPDATE  TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA SET VCHA_ENT_MARCA_POSIBLES = '', vcha_ent_marca_traspaso = '' WHERE VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento_apartados_Cantia + "' AND INTE_eNT_NUMERO = " + CStr(var_consecutivo_apartados_Cantia), cnn, adOpenDynamic, adLockOptimistic
   rsaux.Open "UPDATE  TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA SET VCHA_ENT_MARCA_POSIBLES = '*' WHERE (FLOA_ENT_CANTIDAD_ALMACEN_APARTADO*(0-1)) + (FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR * (0 - 1)) = 0 AND VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento_apartados_Cantia + "' AND INTE_eNT_NUMERO = " + CStr(var_consecutivo_apartados_Cantia), cnn, adOpenDynamic, adLockOptimistic
   var_cadena = "select * from TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA where ((floa_ent_cantidad < ((FLOA_ENT_CANTIDAD_ALMACEN_APARTADO*(0-1)) + (FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR * (0 - 1))) and (isnull(floa_Ent_Cantidad_almacen_apartado_pasar,0) + isnull(floa_ent_cantidad_almacen_apartado_entregar_pasar,0)) = floa_ent_Cantidad) or"
   var_cadena = var_cadena + " (floa_ent_cantidad >= ((FLOA_ENT_CANTIDAD_ALMACEN_APARTADO*(0-1)) + (FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR * (0 - 1)))) and (((FLOA_ENT_CANTIDAD_ALMACEN_APARTADO*(0-1)) + (FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR * (0 - 1))) =  (isnull(floa_Ent_Cantidad_almacen_apartado_pasar,0) + isnull(floa_ent_cantidad_almacen_apartado_entregar_pasar,0)))) AND VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento_apartados_Cantia + "' AND INTE_eNT_NUMERO = " + CStr(var_consecutivo_apartados_Cantia)
   rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
         rsaux2.Open "UPDATE TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA SET VCHA_ENT_MARCA_POSIBLES = '*' WHERE VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento_apartados_Cantia + "' AND INTE_eNT_NUMERO = " + CStr(var_consecutivo_apartados_Cantia) + " AND VCHA_aRT_aRTICULO_ID = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
         rsaux.MoveNext
   Wend
   rsaux.Close
   rsaux.Open "SELECT * FROM TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA WHERE VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento_apartados_Cantia + "' AND INTE_eNT_NUMERO = " + CStr(var_consecutivo_apartados_Cantia) + " AND VCHA_ENT_MARCA_POSIBLES = ''", cnn, adOpenDynamic, adLockOptimistic
   If rsaux.EOF Then
      var_si = MsgBox("Desea efectuar el traspaso", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar el traspaso", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rsaux1.Open "SELECT * FROM TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA WHERE VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento_apartados_Cantia + "' AND INTE_eNT_NUMERO = " + CStr(var_consecutivo_apartados_Cantia) + " AND VCHA_ENT_MARCA_POSIBLES = '*' AND floa_Ent_Cantidad_almacen_apartado_pasar > 0", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               var_almacen_Destino_traspaso = IIf(IsNull(rsaux1!vcha_tem_almacen_apartado), "", rsaux1!vcha_tem_almacen_apartado)
               var_almacen_origen_traspaso = IIf(IsNull(rsaux1!VCHA_ALM_ALMACEN_ID), "", rsaux1!VCHA_ALM_ALMACEN_ID)
               var_clave_movimiento_traspaso = "T"
               var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen_traspaso, var_clave_movimiento_traspaso, Now, var_numero_folio_traspaso, 0, "", "", var_almacen_origen_traspaso, var_almacen_Destino_traspaso, "", var_clave_usuario_global, fun_NombrePc, "", "", "TRASPASO POR APARTADO", "", "", "", "", 0, 0, 0, "1", 1)
               var_numero_folio_traspaso = var_numero_folio_regreso
               While Not rsaux1.EOF
                     rsaux2.Open "select * from tb_temporal_entradas where vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_ent_numero = " + CStr(var_numero_folio_traspaso) + " and vcha_Art_articulo_id = '" + rsaux1!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     If rsaux2.EOF Then
                        var_cadena = "insert into tb_temporal_entradas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_Art_Articulo_id, floa_ent_Cantidad, floa_ent_costo, floa_ent_precio, inte_ent_año, vcha_ent_almacen_origen)"
                        var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + var_almacen_Destino_traspaso + "','T', " + CStr(var_numero_folio_traspaso) + ",'" + rsaux1!vcha_Art_Articulo_id + "', " + CStr(rsaux1!floa_Ent_Cantidad_almacen_apartado_pasar) + "," + CStr(rsaux1!floa_ent_costo) + "," + CStr(rsaux1!floa_ent_precio) + ",2005,'" + var_almacen_origen_traspaso + "')"
                        rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        var_cadena = "insert into tb_entradas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_Art_Articulo_id, floa_ent_Cantidad, floa_ent_costo, floa_ent_precio, inte_ent_año, vcha_ent_almacen_origen)"
                        var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + var_almacen_Destino_traspaso + "','T', " + CStr(var_numero_folio_traspaso) + ",'" + rsaux1!vcha_Art_Articulo_id + "', " + CStr(rsaux1!floa_Ent_Cantidad_almacen_apartado_pasar) + "," + CStr(rsaux1!floa_ent_costo) + "," + CStr(rsaux1!floa_ent_precio) + ",2005,'" + var_almacen_origen_traspaso + "')"
                        rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  
                        var_cadena = "insert into tb_temporal_salidas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_Sal_numero, vcha_Art_Articulo_id, floa_sal_Cantidad, floa_sal_costo, floa_sal_precio, inte_Sal_año)"
                        var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + var_almacen_origen_traspaso + "','T', " + CStr(var_numero_folio_traspaso) + ",'" + rsaux1!vcha_Art_Articulo_id + "', " + CStr(rsaux1!floa_Ent_Cantidad_almacen_apartado_pasar) + "," + CStr(rsaux1!floa_ent_costo) + "," + CStr(rsaux1!floa_ent_precio) + ",2005)"
                        rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        var_cadena = "insert into tb_salidas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_Sal_numero, vcha_Art_Articulo_id, floa_sal_Cantidad, floa_sal_costo, floa_sal_precio, inte_Sal_año)"
                        var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + var_almacen_origen_traspaso + "','T', " + CStr(var_numero_folio_traspaso) + ",'" + rsaux1!vcha_Art_Articulo_id + "', " + CStr(rsaux1!floa_Ent_Cantidad_almacen_apartado_pasar) + "," + CStr(rsaux1!floa_ent_costo) + "," + CStr(rsaux1!floa_ent_precio) + ",2005)"
                        rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     Else
                        rsaux10.Open "update tb_Temporal_entradas set floa_ent_Cantidad =  floa_ent_cantidad + " + CStr(rsaux1!floa_Ent_Cantidad_almacen_apartado_pasar) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_ent_numero = " + CStr(var_numero_folio_traspaso) + " and vcha_Art_articulo_id = '" + rsaux1!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        rsaux10.Open "update tb_entradas set floa_ent_Cantidad =  floa_ent_cantidad + " + CStr(rsaux1!floa_Ent_Cantidad_almacen_apartado_pasar) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_ent_numero = " + CStr(var_numero_folio_traspaso) + " and vcha_Art_articulo_id = '" + rsaux1!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                
                        rsaux10.Open "update tb_Temporal_salidas set floa_Sal_Cantidad =  floa_sal_cantidad + " + CStr(rsaux1!floa_Ent_Cantidad_almacen_apartado_pasar) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_Sal_numero = " + CStr(var_numero_folio_traspaso) + " and vcha_Art_articulo_id = '" + rsaux1!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        rsaux10.Open "update tb_salidas set floa_Sal_Cantidad =  floa_sal_cantidad + " + CStr(rsaux1!floa_Ent_Cantidad_almacen_apartado_pasar) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_Sal_numero = " + CStr(var_numero_folio_traspaso) + " and vcha_Art_articulo_id = '" + rsaux1!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux2.Close
                     rsaux1.MoveNext
                     
                     
                     
                     
               Wend
               rsaux10.Open "update tb_Encabezado_movimientos set char_emo_estatus = 'I', inte_Emo_bloqueado = 0 where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_emo_numero = " + CStr(var_numero_folio_traspaso), cnn, adOpenDynamic, adLockOptimistic
               Set reporte = appl.OpenReport(App.Path + "\rep_salidas_traspasos.rpt")
               reporte.RecordSelectionFormula = "{VW_SALIDAS_TRASPASOS.VCHA_EMO_ALMACEN_ORIGEN} = '" + var_almacen_origen_traspaso + "' and {VW_SALIDAS_TRASPASOS.INTE_EMO_NUMERO} = " + Str(var_numero_folio_traspaso) + " and {VW_SALIDAS_TRASPASOS.VCHA_MOV_MOVIMIENTO_ID} = 'T' and {VW_SALIDAS_TRASPASOS.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_SALIDAS_TRASPASOS.vcha_uor_unidad_id} = '" + var_unidad_organizacional + "'"
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte de Movimientos"
               frmvistasprevias.Show 1
               Set reporte = Nothing
            End If
            rsaux1.Close
            rsaux1.Open "SELECT * FROM TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA WHERE VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento_apartados_Cantia + "' AND INTE_eNT_NUMERO = " + CStr(var_consecutivo_apartados_Cantia) + " AND VCHA_ENT_MARCA_POSIBLES = '*' AND FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR > 0", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               'var_almacen_Destino_traspaso = IIf(IsNull(rsaux1!vcha_tem_almacen_apartado_entregar), "", rsaux1!vcha_tem_almacen_apartado_entregar)
               var_almacen_Destino_traspaso = "CC_1"
               var_almacen_origen_traspaso = IIf(IsNull(rsaux1!VCHA_ALM_ALMACEN_ID), "", rsaux1!VCHA_ALM_ALMACEN_ID)
               var_clave_movimiento_traspaso = "T"
               var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen_traspaso, var_clave_movimiento_traspaso, Now, var_numero_folio_traspaso, 0, "", "", var_almacen_origen_traspaso, var_almacen_Destino_traspaso, "", var_clave_usuario_global, fun_NombrePc, "", "", "TRASPASO PARA ENTREGAS A DOMICILIO", "", "", "", "", 0, 0, 0, "1", 1)
               var_numero_folio_traspaso = var_numero_folio_regreso
               While Not rsaux1.EOF
                     rsaux2.Open "select * from tb_temporal_entradas where vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_ent_numero = " + CStr(var_numero_folio_traspaso) + " and vcha_Art_articulo_id = '" + rsaux1!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     If rsaux2.EOF Then
                        var_cadena = "insert into tb_temporal_entradas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_Art_Articulo_id, floa_ent_Cantidad, floa_ent_costo, floa_ent_precio, inte_ent_año, vcha_ent_almacen_origen)"
                        var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + var_almacen_Destino_traspaso + "','T', " + CStr(var_numero_folio_traspaso) + ",'" + rsaux1!vcha_Art_Articulo_id + "', " + CStr(rsaux1!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR) + "," + CStr(rsaux1!floa_ent_costo) + "," + CStr(rsaux1!floa_ent_precio) + ",2005,'" + var_almacen_origen_traspaso + "')"
                        rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        var_cadena = "insert into tb_entradas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_Art_Articulo_id, floa_ent_Cantidad, floa_ent_costo, floa_ent_precio, inte_ent_año, vcha_ent_almacen_origen)"
                        var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + var_almacen_Destino_traspaso + "','T', " + CStr(var_numero_folio_traspaso) + ",'" + rsaux1!vcha_Art_Articulo_id + "', " + CStr(rsaux1!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR) + "," + CStr(rsaux1!floa_ent_costo) + "," + CStr(rsaux1!floa_ent_precio) + ",2005,'" + var_almacen_origen_traspaso + "')"
                        rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  
                        var_cadena = "insert into tb_temporal_salidas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_Sal_numero, vcha_Art_Articulo_id, floa_sal_Cantidad, floa_sal_costo, floa_sal_precio, inte_Sal_año)"
                        var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + var_almacen_origen_traspaso + "','T', " + CStr(var_numero_folio_traspaso) + ",'" + rsaux1!vcha_Art_Articulo_id + "', " + CStr(rsaux1!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR) + "," + CStr(rsaux1!floa_ent_costo) + "," + CStr(rsaux1!floa_ent_precio) + ",2005)"
                        rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        var_cadena = "insert into tb_salidas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_Sal_numero, vcha_Art_Articulo_id, floa_sal_Cantidad, floa_sal_costo, floa_sal_precio, inte_Sal_año)"
                        var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + var_almacen_origen_traspaso + "','T', " + CStr(var_numero_folio_traspaso) + ",'" + rsaux1!vcha_Art_Articulo_id + "', " + CStr(rsaux1!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR) + "," + CStr(rsaux1!floa_ent_costo) + "," + CStr(rsaux1!floa_ent_precio) + ",2005)"
                        rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     Else
                        rsaux10.Open "update tb_Temporal_entradas set floa_ent_Cantidad =  floa_ent_cantidad + " + CStr(rsaux1!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_ent_numero = " + CStr(var_numero_folio_traspaso) + " and vcha_Art_articulo_id = '" + rsaux1!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        rsaux10.Open "update tb_entradas set floa_ent_Cantidad =  floa_ent_cantidad + " + CStr(rsaux1!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_ent_numero = " + CStr(var_numero_folio_traspaso) + " and vcha_Art_articulo_id = '" + rsaux1!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                
                        rsaux10.Open "update tb_Temporal_salidas set floa_Sal_Cantidad =  floa_sal_cantidad + " + CStr(rsaux1!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_Sal_numero = " + CStr(var_numero_folio_traspaso) + " and vcha_Art_articulo_id = '" + rsaux1!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        rsaux10.Open "update tb_salidas set floa_Sal_Cantidad =  floa_sal_cantidad + " + CStr(rsaux1!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_Sal_numero = " + CStr(var_numero_folio_traspaso) + " and vcha_Art_articulo_id = '" + rsaux1!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux2.Close
                     rsaux1.MoveNext
               Wend
               rsaux10.Open "update tb_Encabezado_movimientos set char_emo_estatus = 'I', inte_Emo_bloqueado = 0 where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_emo_numero = " + CStr(var_numero_folio_traspaso), cnn, adOpenDynamic, adLockOptimistic
               Set reporte = appl.OpenReport(App.Path + "\rep_salidas_traspasos.rpt")
               reporte.RecordSelectionFormula = "{VW_SALIDAS_TRASPASOS.VCHA_EMO_ALMACEN_ORIGEN} = '" + var_almacen_origen_traspaso + "' and {VW_SALIDAS_TRASPASOS.INTE_EMO_NUMERO} = " + Str(var_numero_folio_traspaso) + " and {VW_SALIDAS_TRASPASOS.VCHA_MOV_MOVIMIENTO_ID} = 'T' and {VW_SALIDAS_TRASPASOS.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_SALIDAS_TRASPASOS.vcha_uor_unidad_id} = '" + var_unidad_organizacional + "'"
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte de Movimientos"
               frmvistasprevias.Show 1
               Set reporte = Nothing
            
            
            End If
            rsaux1.Close
            rsaux1.Open "UPDATE  TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA SET VCHA_ENT_MARCA_traspaso = 'T' WHERE VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento_apartados_Cantia + "' AND INTE_eNT_NUMERO = " + CStr(var_consecutivo_apartados_Cantia), cnn, adOpenDynamic, adLockOptimistic
            Unload Me
         End If
      End If
   Else
      MsgBox "No se puede cerrar el movimiento", vbOKOnly, "ATENCION"
   End If
   rsaux.Close
   
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
   cmd_cargar_movimientos_Click
   Me.frm_cantidad_apartada.Visible = False
   Me.frm_cantidad_apartada_entregar.Visible = False
   Me.frm_cantidad_eliminar_apartada.Visible = False
   Me.frm_cantidad_eliminar_apartada_entregar.Visible = False
End Sub

Private Sub lv_almacen_apartado_entragar_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Me.lv_almacen_apartado_entragar.selectedItem.SubItems(7) <> "CC_1" Then
         Me.frm_cantidad_apartada_entregar.Visible = True
         Me.txt_cantidad_apartada_entregar = ""
         Me.txt_cantidad_apartada_entregar.SetFocus
      Else
         MsgBox "No se puede hacer el traspaso", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyCode = 114 Then
      If Me.lv_almacen_apartado_entragar.selectedItem <> "CC_1" Then
         Me.frm_cantidad_eliminar_apartada_entregar.Visible = True
         Me.txt_cantidad_eliminar_apartada_eliminar = ""
         Me.txt_cantidad_eliminar_apartada_eliminar.SetFocus
      Else
         MsgBox "No se puede hacer el traspaso", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub lv_almacen_apartado_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.frm_cantidad_apartada.Visible = True
      Me.txt_cantidad_apartada = ""
      Me.txt_cantidad_apartada.SetFocus
   End If
   If KeyCode = 114 Then
      Me.frm_cantidad_eliminar_apartada.Visible = True
      Me.txt_cantidad_eliminar_apartada = ""
      Me.txt_cantidad_eliminar_apartada.SetFocus
   End If
End Sub

Private Sub txt_cantidad_apartada_entregar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad_apartada_entregar) Then
         rsaux.Open "select * from TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento_apartados_Cantia + "' and inte_ent_numero = " + CStr(var_consecutivo_apartados_Cantia) + " and vcha_Art_Articulo_id = '" + Me.lv_almacen_apartado_entragar.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_cantidad_pasada = IIf(IsNull(rsaux!floa_Ent_Cantidad_almacen_apartado_pasar), 0, rsaux!floa_Ent_Cantidad_almacen_apartado_pasar) + IIf(IsNull(rsaux!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR), 0, rsaux!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR)
            'MsgBox (Me.lv_almacen_apartado_entragar.selectedItem.SubItems(2))
            If var_cantidad_pasada + CDbl(Me.txt_cantidad_apartada_entregar) <= CDbl(Me.lv_almacen_apartado_entragar.selectedItem.SubItems(2)) Then
               If CDbl(Me.lv_almacen_apartado_entragar.selectedItem.SubItems(4)) + CDbl(Me.txt_cantidad_apartada_entregar) <= CDbl(Me.lv_almacen_apartado_entragar.selectedItem.SubItems(3)) Then
                  rsaux1.Open "UPDATE TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA SET FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR = ISNULL(FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR,0) + " + Me.txt_cantidad_apartada_entregar + "  where  vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento_apartados_Cantia + "' and inte_ent_numero = " + CStr(var_consecutivo_apartados_Cantia) + " and vcha_Art_Articulo_id = '" + Me.lv_almacen_apartado_entragar.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                  Me.lv_almacen_apartado_entragar.selectedItem.SubItems(4) = Format(CDbl(Me.lv_almacen_apartado_entragar.selectedItem.SubItems(4)) + CDbl(Me.txt_cantidad_apartada_entregar), "###,###,##0.00")
                  Me.lv_almacen_apartado_entragar.SetFocus
                  Me.frm_cantidad_apartada_entregar.Visible = False
               Else
                  MsgBox "La existencia es menor a la cantidad a pasar", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No se puede pasar mas de lo disponible", vbOKOnly, "ATENCION"
            End If
         End If
         rsaux.Close
      Else
         MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      Me.lv_almacen_apartado_entragar.SetFocus
      Me.frm_cantidad_apartada.Visible = False
   End If
End Sub

Private Sub txt_cantidad_apartada_entregar_LostFocus()
   Me.frm_cantidad_apartada_entregar.Visible = False
End Sub

Private Sub txt_cantidad_apartada_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad_apartada) Then
         rsaux.Open "select * from TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento_apartados_Cantia + "' and inte_ent_numero = " + CStr(var_consecutivo_apartados_Cantia) + " and vcha_Art_Articulo_id = '" + Me.lv_almacen_apartado.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_cantidad_pasada = IIf(IsNull(rsaux!floa_Ent_Cantidad_almacen_apartado_pasar), 0, rsaux!floa_Ent_Cantidad_almacen_apartado_pasar) + IIf(IsNull(rsaux!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR), 0, rsaux!FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR)
            If CDbl(var_cantidad_pasada) + CDbl(Me.txt_cantidad_apartada) <= Me.lv_almacen_apartado.selectedItem.SubItems(2) Then
               If CDbl(Me.lv_almacen_apartado.selectedItem.SubItems(4)) + CDbl(Me.txt_cantidad_apartada) <= CDbl(Me.lv_almacen_apartado.selectedItem.SubItems(3)) Then
                  rsaux1.Open "UPDATE TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA SET FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_PASAR = ISNULL(FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_PASAR,0) + " + Me.txt_cantidad_apartada + "  where  vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento_apartados_Cantia + "' and inte_ent_numero = " + CStr(var_consecutivo_apartados_Cantia) + " and vcha_Art_Articulo_id = '" + Me.lv_almacen_apartado.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                  Me.lv_almacen_apartado.selectedItem.SubItems(4) = Format(CDbl(lv_almacen_apartado.selectedItem.SubItems(4)) + CDbl(Me.txt_cantidad_apartada), "###,###,##0.00")
                  Me.lv_almacen_apartado.SetFocus
                  Me.frm_cantidad_apartada.Visible = False
               Else
                  MsgBox "La existencia es menor a la cantidad a pasar", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No se puede pasar mas de lo disponible", vbOKOnly, "ATENCION"
            End If
         End If
         rsaux.Close
      Else
         MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   
   If KeyAscii = 27 Then
      Me.lv_almacen_apartado.SetFocus
      Me.frm_cantidad_apartada.Visible = False
   End If
End Sub


Private Sub txt_cantidad_apartada_LostFocus()
   Me.frm_cantidad_apartada.Visible = False
End Sub

Private Sub txt_cantidad_eliminar_apartada_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad_eliminar_apartada_eliminar) Then
        If CDbl(Me.txt_cantidad_eliminar_apartada_eliminar) <= Me.lv_almacen_apartado_entragar.selectedItem.SubItems(4) Then
           rsaux1.Open "UPDATE TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA SET FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_eNTREGAR_PASAR = ISNULL(FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_ENTREGAR_PASAR,0) - " + Me.txt_cantidad_eliminar_apartada_eliminar + "  where  vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento_apartados_Cantia + "' and inte_ent_numero = " + CStr(var_consecutivo_apartados_Cantia) + " and vcha_Art_Articulo_id = '" + Me.lv_almacen_apartado_entragar.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
           Me.lv_almacen_apartado_entragar.selectedItem.SubItems(4) = Format(CDbl(Me.lv_almacen_apartado_entragar.selectedItem.SubItems(4)) - CDbl(Me.txt_cantidad_eliminar_apartada_eliminar), "###,###,##0.00")
           Me.lv_almacen_apartado_entragar.SetFocus
           Me.frm_cantidad_eliminar_apartada_entregar.Visible = False
        Else
           MsgBox "La cantidad es mayor a la cantidad pasada", vbOKOnly, "ATENCION"
        End If
      Else
         MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      Me.lv_almacen_apartado_entragar.SetFocus
      Me.frm_cantidad_eliminar_apartada_entregar.Visible = False
   End If
End Sub

Private Sub txt_cantidad_eliminar_apartada_eliminar_LostFocus()
   Me.frm_cantidad_eliminar_apartada_entregar.Visible = False
End Sub

Private Sub txt_cantidad_eliminar_apartada_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad_eliminar_apartada) Then
        If CDbl(Me.txt_cantidad_eliminar_apartada) <= Me.lv_almacen_apartado.selectedItem.SubItems(4) Then
           rsaux1.Open "UPDATE TB_TEMPORAL_EXISTENCIAS_ALMACENES_APARTADOS_CANTIA SET FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_PASAR = ISNULL(FLOA_ENT_CANTIDAD_ALMACEN_APARTADO_PASAR,0) - " + Me.txt_cantidad_eliminar_apartada + "  where  vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento_apartados_Cantia + "' and inte_ent_numero = " + CStr(var_consecutivo_apartados_Cantia) + " and vcha_Art_Articulo_id = '" + Me.lv_almacen_apartado.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
           Me.lv_almacen_apartado.selectedItem.SubItems(4) = Format(CDbl(lv_almacen_apartado.selectedItem.SubItems(4)) - CDbl(Me.txt_cantidad_eliminar_apartada), "###,###,##0.00")
           Me.frm_cantidad_eliminar_apartada.Visible = False
        Else
           MsgBox "La cantidad es mayor a la cantidad pasada", vbOKOnly, "ATENCION"
        End If
      Else
         MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_cantidad_eliminar_apartada.Visible = False
   End If
End Sub

Private Sub txt_cantidad_eliminar_apartada_LostFocus()
   Me.frm_cantidad_eliminar_apartada.Visible = False
End Sub
