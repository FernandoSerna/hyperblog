VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsalidas_almacen_calidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salidas del almacen de calidad"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   3015
      TabIndex        =   23
      Top             =   2310
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   24
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
         TabIndex        =   25
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_reproceso 
      Caption         =   " Movimiento "
      Height          =   1215
      Left            =   2025
      TabIndex        =   30
      Top             =   450
      Width           =   7665
      Begin VB.TextBox txt_numero_reproceso 
         Height          =   330
         Left            =   4905
         TabIndex        =   43
         Top             =   450
         Width           =   1350
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame6"
         Height          =   1185
         Left            =   2280
         TabIndex        =   42
         Top             =   15
         Width           =   30
      End
      Begin VB.OptionButton opt_reproceso_interno 
         Caption         =   "Reproceso interno"
         Height          =   270
         Left            =   180
         TabIndex        =   37
         Top             =   735
         Width           =   1980
      End
      Begin VB.OptionButton opt_devolucion_cliente 
         Caption         =   "Devolución de cliente"
         Height          =   465
         Left            =   180
         TabIndex        =   36
         Top             =   300
         Width           =   1890
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Reimpresión número:"
         Height          =   195
         Left            =   3255
         TabIndex        =   44
         Top             =   525
         Width           =   1485
      End
   End
   Begin VB.Frame frm_planta 
      Caption         =   " Movimiento "
      Height          =   1215
      Left            =   2025
      TabIndex        =   29
      Top             =   450
      Width           =   7665
      Begin VB.TextBox txt_numero_traspaso 
         Height          =   330
         Left            =   1800
         TabIndex        =   41
         Top             =   780
         Width           =   1350
      End
      Begin VB.TextBox txt_planta 
         Height          =   330
         Left            =   1800
         TabIndex        =   10
         Top             =   270
         Width           =   1350
      End
      Begin VB.TextBox txt_nombre_planta 
         Height          =   330
         Left            =   3165
         TabIndex        =   11
         Top             =   270
         Width           =   4290
      End
      Begin VB.Frame Frame4 
         Height          =   120
         Left            =   15
         TabIndex        =   39
         Top             =   585
         Width           =   7620
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Reimpresión número:"
         Height          =   195
         Left            =   150
         TabIndex        =   40
         Top             =   855
         Width           =   1485
      End
      Begin VB.Label Label3 
         Caption         =   "Planta:"
         Height          =   270
         Left            =   180
         TabIndex        =   38
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmsalidas_almacen_calidad.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmsalidas_almacen_calidad.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9375
      Picture         =   "frmsalidas_almacen_calidad.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_cliente 
      Caption         =   " Cliente "
      Height          =   1215
      Left            =   2025
      TabIndex        =   22
      Top             =   450
      Width           =   7665
      Begin VB.TextBox txt_nombre_Establecimiento 
         Height          =   315
         Left            =   2640
         TabIndex        =   9
         Top             =   735
         Width           =   4575
      End
      Begin VB.TextBox txt_establecimiento 
         Height          =   315
         Left            =   1500
         TabIndex        =   8
         Top             =   735
         Width           =   1125
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   330
         Left            =   2640
         TabIndex        =   7
         Top             =   345
         Width           =   4575
      End
      Begin VB.TextBox txt_cliente 
         Height          =   330
         Left            =   1500
         TabIndex        =   6
         Top             =   345
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   28
         Top             =   765
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   345
         Width           =   525
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Tipo de salida "
      Height          =   1215
      Left            =   60
      TabIndex        =   21
      Top             =   450
      Width           =   1905
      Begin VB.OptionButton opt_traspaso 
         Caption         =   "Traspaso"
         Height          =   270
         Left            =   240
         TabIndex        =   5
         Top             =   780
         Width           =   1530
      End
      Begin VB.OptionButton opt_venta 
         Caption         =   "Venta"
         Height          =   270
         Left            =   1425
         TabIndex        =   3
         Top             =   150
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.OptionButton opt_reproceso 
         Caption         =   "Reproceso"
         Height          =   270
         Left            =   240
         TabIndex        =   4
         Top             =   435
         Width           =   1530
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5550
      Left            =   60
      TabIndex        =   12
      Top             =   1665
      Width           =   9660
      Begin VB.TextBox txt_busqueda 
         Height          =   345
         Left            =   3120
         TabIndex        =   34
         Top             =   127
         Width           =   2115
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   6330
         TabIndex        =   31
         Top             =   2370
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   60
            TabIndex        =   32
            Top             =   375
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   "Cantidad "
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   33
            Top             =   15
            Width           =   2895
         End
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1695
         Picture         =   "frmsalidas_almacen_calidad.frx":083E
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Pasar todos F4"
         Top             =   135
         Width           =   330
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   30
         TabIndex        =   19
         Top             =   480
         Width           =   9630
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1365
         Picture         =   "frmsalidas_almacen_calidad.frx":0940
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   705
         Picture         =   "frmsalidas_almacen_calidad.frx":0B56
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Marcar (Enter)"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1035
         Picture         =   "frmsalidas_almacen_calidad.frx":0DA0
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton Command9 
         Height          =   315
         Left            =   45
         Picture         =   "frmsalidas_almacen_calidad.frx":0E72
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmsalidas_almacen_calidad.frx":0F74
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   135
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_existencias 
         Height          =   4920
         Left            =   45
         TabIndex        =   13
         Top             =   570
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   8678
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6085
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Costo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Precio"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Pasar"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Precio 0"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Equivalencia"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   2535
         TabIndex        =   35
         Top             =   195
         Width           =   540
      End
   End
   Begin VB.Frame Frame5 
      Height          =   120
      Left            =   30
      TabIndex        =   26
      Top             =   270
      Width           =   9705
   End
End
Attribute VB_Name = "frmsalidas_almacen_calidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_tipo_lista As Integer
Dim list_item As ListItem
Dim cnn_cantia As ADODB.Connection


Private Sub cmd_imprimir_Click()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_ENCABEZADO_MOVIMIENTOS_I = New TB_ENCABEZADO_MOVIMIENTOS_I
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Dim var_conexion_intercompañia As String
   Dim var_inserta As Boolean
   Dim var_posible_articulos As Boolean
   Dim var_primera_vez  As Boolean
   Dim var_cantidad_leida As Double
   Dim var_costo As Double
   Dim var_precio As Double
   Dim var_almacen_Destino As String
   Dim var_clave_movimiento As String
   Dim var_clave_moneda As String
   Dim txt_proveedor As String
   Dim txt_nombre_proveedor As String
   Dim var_numero_folio  As Double
   
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
   
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
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
   
   
   var_primera_vez = True
   If Me.opt_traspaso.Value = True Then
   
      var_posible_articulos = False
      If Me.txt_planta <> "" Then
         If Me.lv_existencias.ListItems.Count > 0 Then
            For var_j = 1 To Me.lv_existencias.ListItems.Count
                Me.lv_existencias.ListItems.Item(var_j).Selected = True
                If CDbl(Me.lv_existencias.selectedItem.SubItems(5)) > 0 Then
                   var_posible_articulos = True
                End If
            Next var_j
            If var_posible_articulos = True Then
               var_si = MsgBox("¿Desea cerrar el movimiento?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  var_si = MsgBox("Confirmar el cerrado del movimiento", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     var_primera_vez = True
                     var_almacen_Destino = "CAEE"
                     var_clave_movimiento = "DPL"
                     var_clave_moneda = "1"
                     txt_proveedor = Me.txt_planta
                     txt_nombre_proveedor = Me.txt_nombre_planta
                     For var_j = 1 To Me.lv_existencias.ListItems.Count
                         lv_existencias.ListItems.Item(var_j).Selected = True
                         If CDbl(lv_existencias.selectedItem.SubItems(5)) > 0 Then
                            txt_codigo = lv_existencias.selectedItem
                            var_cantidad_leida = CDbl(Me.lv_existencias.selectedItem.SubItems(5))
                            'MsgBox cnn_cantia.ConnectionString
                            rsaux10.Open "select * from tb_producto where vcha_pro_producto_id = '" + txt_codigo + "'", cnn_cantia, adOpenDynamic, adLockOptimistic
                            var_costo = IIf(IsNull(rsaux10!mon_pro_costorea), 0, rsaux10!mon_pro_costorea)
                            rsaux10.Close
                            var_precio = CDbl(Me.lv_existencias.selectedItem.SubItems(4))
                            If Trim(txt_codigo) <> "" Then
                               If var_primera_vez = True Then
                                  var_inserta = False
                                  var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", txt_proveedor, "", var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", txt_nombre_proveedor, "", "B", "", "", 0, 0, 0, var_clave_moneda, 0)
                                  var_numero_folio = var_numero_folio_regreso
                                  var_primera_vez = False
                               End If 'aqui voy
                               var_posible_leido = 1
                               If var_posible_leido = 1 Then
                                  Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + CStr(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
                                  rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                  If Not rs.EOF Then
                                     var_inserta = False
                                     var_inserta = TB_TEMPORAL_SALIDAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida)
                                     rs.Close
                                  Else
                                     var_inserta = False
                                     rsaux.Open "INSERT INTO TB_TEMPORAL_SALIDAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ")", cnn, adOpenDynamic, adLockOptimistic
                                     rs.Close
                                  End If
                               End If
                            End If
                         End If
                     Next var_j
                     
                  
            
                     Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
                     rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     While Not rs.EOF
                           var_inserta = False
                           var_suma_cantidad = 0
                           var_cantidad_llegar = IIf(IsNull(rs!FLOA_sAL_cANTIDAD), 0, rs!FLOA_sAL_cANTIDAD)
                           var_cantidad = 0
                           While var_suma_cantidad < var_cantidad_llegar
                                 rsaux2.Open "select * from tb_existencias where vcha_art_articulo_id =  '" + rs!vcha_Art_articulo_id + "' and vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux2.EOF Then
                                    If rsaux2!floa_exi_cantidad_2004 >= var_cantidad_llegar Then
                                       var_año = 2004
                                       var_suma_cantidad = var_cantidad_llegar
                                       var_cantidad = var_cantidad_llegar
                                       var_costo = IIf(IsNull(rsaux2!FLOA_EXI_COSTO_2004), 0, rsaux2!FLOA_EXI_COSTO_2004)
                                    Else
                                       var_cantidad_disponible = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                       If var_cantidad_disponible > 0 Then
                                          var_año = 2004
                                          var_suma_cantidad = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                          var_cantidad = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                          var_costo = rsaux2!FLOA_EXI_COSTO_2004
                                       Else
                                          var_año = 2005
                                          var_cantidad = rs!FLOA_sAL_cANTIDAD - var_suma_cantidad
                                          var_suma_cantidad = var_cantidad_llegar
                                          var_costo = IIf(IsNull(rsaux2!floa_exi_costo_2005), 0, rsaux2!floa_exi_costo_2005)
                                       End If
                                    End If
                                 Else
                                    var_año = 2005
                                    var_suma_cantidad = var_cantidad_llegar
                                    var_cantidad = var_cantidad_llegar
                                    rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id =  '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux4.EOF Then
                                       var_costo = IIf(IsNull(rsaux4!mone_Art_costo_estandar), 0, rsaux4!mone_Art_costo_estandar)
                                    Else
                                       var_costo = 0
                                    End If
                                    rsaux4.Close
                                 End If
                                 rsaux2.Close
                                 If var_costo = 0 Then
                                    rsaux4.Open "select floa_exi_costo_2005 from tb_existencias where vcha_alm_almacen_id = '8' and vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux4.EOF Then
                                       var_costo = IIf(IsNull(rsaux4!floa_exi_costo_2005), 0, rsaux4!floa_exi_costo_2005)
                                    Else
                                       var_costo = 0
                                    End If
                                    rsaux4.Close
                                 End If
                                 rsaux.Open "insert into tb_salidas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_sal_numero, vcha_art_articulo_id, floa_sal_cantidad, floa_sal_costo, floa_sal_precio, inte_sal_año) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_articulo_id + "', " + CStr(var_cantidad) + ", " + CStr(var_costo) + " , " + CStr(rs!floa_Sal_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
                           Wend
                           rs.MoveNext
                     Wend
                     rs.Close
                     var_estatus_movimiento = "I"
                     var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
                     var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
                  
                     Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_SALIDA.rpt")
                     reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_SALIDA.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_Destino + "' AND {VW_MOVIMIENTOS_SALIDA.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_MOVIMIENTOS_SALIDA.INTE_EMO_NUMERO} = " + Str(var_numero_folio)
                     frmvistasprevias.cr.ReportSource = reporte
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     frmvistasprevias.cr.ViewReport
                     frmvistasprevias.Caption = "Reporte de Movimientos"
                     frmvistasprevias.Show 1
                     Set reporte = Nothing
                     rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1, inte_emo_bloqueado = 0 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  
               
                  End If
               End If
            Else
               MsgBox "No se a seleccionado ningún artículo para el traspaso a plantas", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El movimiento no contiene artículos", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a seleccionado una planta", vbOKOnly, "ATENCION"
      End If
   End If 'fin del traspaso
   
   If opt_venta.Value = True Then
      'rs.Open "SELECT * FROM TB_ALMACENES WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
      'If Not rs.EOF Then
      '   var_almacen_Destino = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
      '   var_almacen_origen = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
      'End If
      'rs.Close
   
      'inserta en temporal salidas
      Dim var_posible_disponible As Boolean
      Dim var_posible_movimiento As Boolean
      Dim var_numero_movimiento_leido As Double
      Dim var_numero_folio_entrada As Double
      var_año = 2005
      'rs.Open "select * from tb_encabezado_movimientos where vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_emo_referencia = '" + Me.txt_referencia + "'", cnn, adOpenDynamic, adLockOptimistic
      'If rs.EOF Then
      '   var_posible_movimiento = True
      'Else
      '   var_numero_movimiento_leido = rs!inte_emo_numero
      '   var_posible_movimiento = False
      'End If
      'rs.Close
      var_posible_movimiento = True
      If var_posible_movimiento = True Then
         rs.Open "select vcha_lis_lista_id from tb_clientes where vcha_cli_clave_id = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
         var_lista_precios = IIf(IsNull(rs(0).Value), "", rs(0).Value)
         rs.Close
         var_cadena_precios = ""
         For var_j = 1 To Me.lv_existencias.ListItems.Count
             Me.lv_existencias.ListItems.Item(var_j).Selected = True
             If CDbl(Me.lv_existencias.selectedItem.SubItems(5)) > 0 Then
                rs.Open "select * from tb_Detalle_lista_precios where vcha_lis_lista_precios_id = '" + var_lista_precios + "' and vcha_Art_Articulo_id = '" + Me.lv_existencias.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                If rs.EOF Then
                   If var_cadena_precios = "" Then
                      var_cadena_precios = Me.lv_existencias.selectedItem + " " + Me.lv_existencias.selectedItem.SubItems(1)
                   Else
                      var_cadena_precios = var_cadena_precios + ", " + Me.lv_existencias.selectedItem + " " + Me.lv_existencias.selectedItem.SubItems(1)
                   End If
                End If
                rs.Close
             End If
         Next var_j
         If var_cadena_precios = "" Then
            If var_estatus_movimiento <> "I" Then
               var_si = MsgBox("¿Desea cerrar el movimiento?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  var_si = MsgBox("Confirmar el cerrado del movimiento", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     For var_zzz = 1 To Me.lv_existencias.ListItems.Count
                         var_almacen_origen = "CAEE"
                         var_clave_movimiento = "VDI"
                         Me.lv_existencias.ListItems.Item(var_zzz).Selected = True
                         If CDbl(Me.lv_existencias.selectedItem.SubItems(5)) > 0 Then
                            txt_codigo = Me.lv_existencias.selectedItem
                            var_cantidad_leida = CDbl(Me.lv_existencias.selectedItem.SubItems(5))
                         
                            rsaux10.Open "select * from tb_producto where vcha_pro_producto_id = '" + txt_codigo + "'", cnn_cantia, adOpenDynamic, adLockOptimistic
                            If Not rsaux10.EOF Then
                               var_costo = IIf(IsNull(rsaux10!mon_pro_costorea), 0, rsaux10!mon_pro_costorea)
                            Else
                               rsaux11.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_aRTICULO_ID = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                               If Not rsaux11.EOF Then
                                  var_costo = IIf(IsNull(rsaux11!mone_Art_costo_estandar), 0, rsaux11!mone_Art_costo_estandar)
                               Else
                                  var_costo = 0
                               End If
                               rsaux11.Close
                            End If
                            rsaux10.Close
                            
                            If Trim(txt_codigo) <> "" Then
                               If rsaux5.State = 1 Then
                                  rsaux5.Close
                               End If
                               If var_empresa = "18" Then
                                  rsaux5.Open "select isnull(floa_Exi_cantidad_disponible,0) - isnull(floa_Exi_Temporal_cantidad_salida,0) from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_Art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                  If Not rsaux5.EOF Then
                                     var_posible_disponible = True
                                  Else
                                     var_posible_disponible = False
                                  End If
                               Else
                                  var_posible_disponible = True
                               End If
                               If var_posible_disponible = True Then
                                  'Me.txt_referencia.Enabled = False
                                  bandera_suma = False
                                  If var_primera_vez = True Then
                                     var_primera_vez = False
                                     rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                                     If Not rs.EOF Then
                                        var_lista_precios = IIf(IsNull(rs!vcha_LIS_LISTA_iD), "", rs!vcha_LIS_LISTA_iD)
                                        var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                                        var_clave_titular = rs!vcha_tit_titular_id
                                        var_canal_venta = IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id)
                                        var_descuento_1 = IIf(IsNull(rs!floa_gac_Descuento_1), 0, rs!floa_gac_Descuento_1)
                                        var_descuento_2 = IIf(IsNull(rs!FLOA_GAC_DESCUENTO_2), 0, rs!FLOA_GAC_DESCUENTO_2)
                                        txt_agente = rs!VCHA_AGE_AGENTE_ID
                                     End If
                                     rs.Close
                                     'rs.Open "select * from tb_Detalle_Establecimientos where vcha_cli_clave_id = '" + txt_almacen_destino + "'", cnn, adOpenDynamic, adLockOptimistic
                                     'var_clave_establecimiento = ""
                                     'If Not rs.EOF Then
                                     '   var_clave_establecimiento = rs!vcha_esb_establecimiento_id
                                     'End If
                                     'rs.Close
                                     var_clave_establecimiento = Me.txt_establecimiento
                                     var_numero_folio = 0
                                     var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, CStr(var_almacen_origen), CStr(var_clave_movimiento), Now, CDbl(var_numero_folio), 0, CStr(Me.txt_cliente), "", CStr(var_almacen_origen), "", "", var_clave_usuario_global, fun_NombrePc, 0, "", "", CStr(var_clave_establecimiento), "", CStr(var_clave_titular), CStr(txt_agente), CDbl(var_descuento_1), CDbl(var_descuento_2), 0, CStr(var_clave_moneda), 0)
                                     var_numero_folio = var_numero_folio_regreso
                                     txt_folio = var_numero_folio
                                   
                                     'var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, "EPVD", Now, CDbl(var_numero_folio), 0, CStr(txt_almacen_destino), "", var_almacen_origen, var_almacen_origen, "", var_clave_usuario_global, fun_NombrePc, 0, "", "", var_clave_establecimiento, "", var_clave_titular, CStr(txt_agente), var_descuento_1, var_descuento_2, 0, var_clave_moneda, 0)
                                     'var_numero_folio_entrada = var_numero_folio_regreso
                              
                            
                                     rsaux.Open "update tb_encabezado_movimientos set VCHA_CAN_CANAL_VENTA_ID = '" + var_canal_venta + "', vcha_emo_referencia = 'SALIDA DEL ALMACEN DE CALIDAD' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                                     'rsaux.Open "update tb_encabezado_movimientos set VCHA_CAN_CANAL_VENTA_ID = '" + var_canal_venta + "', vcha_emo_referencia = '" + Me.txt_referencia + "' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio_entrada) + " and vcha_mov_movimiento_id = 'EPVD'", cnn, adOpenDynamic, adLockOptimistic
                                  End If
               
              
                                  If var_posible_kanban = 1 Then
                                     Set TB_RESERVAR_FUERA_DE_KANBAN = New TB_RESERVAR_FUERA_DE_KANBAN
                                     Set TB_RESERVAR_KANBAN = New TB_RESERVAR_KANBAN
                                     If var_kanban_es_un_kanban = "S" Then
                                        var_inserta = TB_RESERVAR_KANBAN.Anadir(var_kanban, var_clave_movimiento, var_numero_folio, var_almacen_origen, txt_codigo, "", "")
                                        If var_kanban_exito = "S" Then
                                           var_posible_leido = 1
                                        Else
                                           var_posible_leido = 0
                                        End If
                                     Else
                                        var_inserta = TB_RESERVAR_FUERA_DE_KANBAN.Anadir(var_numero_folio, var_clave_movimiento, var_almacen_origen, txt_codigo, "", "")
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
                                     rsaux4.Open "select floa_dli_precio from tb_detalle_lista_precios where vcha_Art_Articulo_id = '" + txt_codigo + "' and vcha_lis_lista_precios_id = '" + var_lista_precios + "'", cnn, adOpenDynamic, adLockOptimistic
                                     If Not rsaux4.EOF Then
                                        If Not rsaux4.EOF Then
                                           var_precio = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                        End If
                                        'If var_clasificacion <> "PRIMERA" Then
                                        '   var_precio = 0
                                        'End If
                                        rsaux4.Close
                                        If var_empresa = "18" Then
                                           rs.Open "update tb_existencias set floa_Exi_temporal_cantidad_salida = isnull(floa_Exi_temporal_cantidad_salida,0) + " + CStr(var_cantidad_leida) + " where vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_Art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                              
                                        Cadena = "select * from tb_temporal_salidas with (nolock) where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
                               
                                        rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                        If Not rs.EOF Then
                                           var_inserta = False
                                           var_inserta = TB_TEMPORAL_SALIDAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida)
                                        Else
                                           var_inserta = False
                                           var_inserta = TB_TEMPORAL_SALIDAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, CStr(var_almacen_origen), CStr(var_clave_movimiento), CDbl(var_numero_folio), CStr(txt_codigo), CDbl(var_cantidad_leida), var_costo, var_precio, "0", 0, 0)
                                        End If
                                        rs.Close
                                        'Cadena = "select * from TB_TEMPORAL_ENTRADAS where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = 'EPVD' and inte_ent_numero = " + Str(var_numero_folio_entrada) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
                                        'rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                        'If Not rs.EOF Then
                                        '   var_inserta = False
                                        '   var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, "EPVD", var_numero_folio_entrada, txt_codigo, var_cantidad_leida, var_año)
                                        'Else
                                        '   var_inserta = False
                                        '   var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, "EPVD", var_numero_folio_entrada, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", var_almacen_origen, var_año)
                                        'End If
                                        'rs.Close
                                     Else
                                        rsaux4.Close
                                        rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                        If Not rsaux4.EOF Then
                                           frmmensaje.lbl_articulo = IIf(IsNull(rsaux4!vcha_art_nombre_español), "", rsaux4!vcha_art_nombre_español)
                                        End If
                                        txt_codigo = ""
                                        frmmensaje.lbl_mensaje = "El artículo no se encuentra en la lista de precios del cliente"
                                        frmmensaje.Show 1
                                        rsaux4.Close
                                    End If
                                  Else
                                     frmmensaje.lbl_mensaje = var_kanban_mensaje
                                     frmmensaje.Show 1
                                     txt_codigo = ""
                                  End If
                               Else
                                  rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                  If Not rsaux4.EOF Then
                                     frmmensaje.lbl_articulo = IIf(IsNull(rsaux4!vcha_art_nombre_español), "", rsaux4!vcha_art_nombre_español)
                                  End If
                                  txt_codigo = ""
                                  frmmensaje.lbl_mensaje = "El artículo no se encuentra en el inventario del almacén"
                                  frmmensaje.Show 1
                                  rsaux4.Close
                               End If
                               If rsaux5.State = 1 Then
                                  rsaux5.Close
                               End If
                            End If
                         End If
                     Next var_zzz
                  End If
               End If
            End If
' fin de inserta en temporal salidas
            C = 1
            If C = 1 Then
               Dim var_correo_electronico As String
               If var_numero_folio > 0 Then
                  If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_ventas_directas.rpt")
                     reporte.RecordSelectionFormula = "{VW_VENTAS_DIRECTAS.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_origen + "' and {VW_VENTAS_DIRECTAS.INTE_sal_NUMERO} = " + Str(var_numero_folio) + " and {VW_VENTAS_DIRECTAS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_VENTAS_DIRECTAS.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_VENTAS_DIRECTAS.vcha_uor_unidad_id} = '" + var_unidad_organizacional + "'"
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
                     'var_si = MsgBox("¿Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
                     var_si = 1
                     If var_si = 1 Then
                        cnn.BeginTrans
                        var_posible_cerrar_KANBAN = True
                        If var_posible_kanban = 1 Then
                           Set TB_PROC_KANBANS_EN_MOVIMIENTO = New TB_PROC_KANBANS_EN_MOVIMIENTO
                           var_inserta = TB_PROC_KANBANS_EN_MOVIMIENTO.Anadir(var_almacen_origen, var_clave_movimiento, CDbl(txt_folio), "", "")
                           If var_kanban_exito = "N" Then
                              var_posible_cerrar_KANBAN = False
                           End If
                        Else
                           var_posible_cerrar_KANBAN = True
                        End If
                        If var_posible_cerrar_KANBAN = True Then
                           rs.Open "select * from VW_clientes where vcha_cli_clave_id = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
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
                           ok = TB_ENC_PEDIDOS_I.Anadir(var_empresa, var_unidad_organizacional, CStr(var_almacen_origen), "M", maximo_pedido, 0, Date, Date, CStr(txt_agente), CStr(txt_titular), Me.txt_cliente, "", 0, 0, "", var_descuento_1, var_descuento_2, 0, CDbl(var_dias_condiciones), 0, var_clave_usuario_global, fun_NombrePc, Date, var_clave_moneda, 0)
                           rsaux.Open "update tb_encabezado_pedidos set vcha_ped_pedido_externo = '' where inte_ped_numero = " + CStr(maximo_pedido), cnn, adOpenDynamic, adLockOptimistic
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
                                 rsaux4.Open "update tb_existencias set floa_exi_temporal_cantidad_salida = isnull(floa_exi_temporal_cantidad_salida,0) - " + CStr(rsaux5!FLOA_sAL_cANTIDAD) + " where vcha_alm_almacen_id = '" + rsaux5!VCHA_ALM_ALMACEN_ID + "' and vcha_Art_Articulo_id = '" + rsaux5!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                 rsaux5.MoveNext
                           Wend
                           ok = TB_ENC_ORDEN_SURTIDO.Anadir(var_empresa, var_unidad_organizacional, "M", maximo_pedido, CStr(var_almacen_origen), CDbl(var_maximo_orden), Date, Date + 0, "", CStr(txt_titular), Me.txt_cliente, "", var_descuento_1, var_descuento_2, 0, "", "", Date, 0, var_clave_moneda, Date)
                           rs.Open "update tb_encabezado_movimientos set inte_emo_numero_origen = " + CStr(var_maximo_orden) + " where inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                           cnn.CommandTimeout = 360
                           If var_empresa = "15" Then
                              rsaux4.Open "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES_PLANTAS '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", 1,'','','" + var_clave_titular + "','" + Me.txt_cliente + "',0,0,0", cnn, adOpenDynamic, adLockOptimistic
                           Else
                              rsaux4.Open "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", 1,'','','" + var_clave_titular + "','" + Me.txt_cliente + "',0,0,0", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           
                           'Cadena = "select * from TB_TEMPORAL_ENTRADAS where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND  vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = 'EPVD' and inte_ent_numero = " + Str(var_numero_folio_entrada)
                           'rsaux6.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                           'While Not rsaux6.EOF
                           '      var_cadena = "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_Articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, inte_ent_año) "
                           '      var_cadena = var_cadena + "    values ('" + rsaux6!vcha_emp_empresa_id + "', '" + rsaux6!VCHA_UOR_UNIDAD_ID + "', '" + rsaux6!VCHA_ALM_ALMACEN_ID + "', '" + rsaux6!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rsaux6!inte_ent_numero) + ", '" + rsaux6!VCHA_aRT_ARTICULO_ID + "', " + CStr(rsaux6!FLOA_ent_CANTIDAD) + ", " + CStr(rsaux6!floa_ent_costo) + ", " + CStr(rsaux6!floa_ent_precio) + ", " + CStr(rsaux6!inte_ent_año) + ")"
                           '      rsaux7.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                           '      rsaux6.MoveNext
                           'Wend
                           'rsaux6.Close
                 
                           'rsaux6.Open "UPDATE TB_ENCABEZADO_MOVIMIENTOS SET CHAR_EMO_ESTATUS = 'I' WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = 'EPVD' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio_entrada), cnn, adOpenDynamic, adLockOptimistic
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
                           var_inserta = TB_DET_EMBARQUE_I.Anadir(var_empresa, var_unidad_organizacional, CStr(var_almacen_origen), var_numero_embarque, var_clave_movimiento, var_numero_folio, "")
                           txt_numero_embarque = CStr(var_numero_embarque)
                           var_estatus_embarque = "I"
                           cnn.CommitTrans
               
                           Set reporte = appl.OpenReport(App.Path + "\rep_ventas_directas.rpt")
                           reporte.RecordSelectionFormula = "{VW_VENTAS_DIRECTAS.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_origen + "' and {VW_VENTAS_DIRECTAS.INTE_SAL_NUMERO} = " + Str(var_numero_folio) + " and {VW_VENTAS_DIRECTAS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_VENTAS_DIRECTAS.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_VENTAS_DIRECTAS.vcha_uor_unidad_id} = '" + var_unidad_organizacional + "'"
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
                           'If var_empresa = "06" Or var_empresa = "15" Then
                           '   var_numero_embarque_global = var_numero_embarque
                           '   frmfactura_embarques.Show
                           'End If
                        Else
                           cnn.RollbackTrans
                           MsgBox "No se pudo cerrar el movimiento kanban", vbOKOnly, "ATENCION"
                        End If
                     End If
                  End If
               Else
                  MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
               End If
            End If
         Else
            MsgBox "Los siguientes artículos no tienen precio " + var_cadena_precios, vbOKOnly, "ATENCION"
         End If
      Else
         If CStr(var_numero_movimiento_leido) = txt_folio Then
            Set reporte = appl.OpenReport(App.Path + "\rep_ventas_directas.rpt")
            reporte.RecordSelectionFormula = "{VW_VENTAS_DIRECTAS.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_origen + "' and {VW_VENTAS_DIRECTAS.INTE_sal_NUMERO} = " + Str(var_numero_folio) + " and {VW_VENTAS_DIRECTAS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_VENTAS_DIRECTAS.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_VENTAS_DIRECTAS.vcha_uor_unidad_id} = '" + var_unidad_organizacional + "'"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Movimientos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La nota ya fue leida en el movimiento " + CStr(var_numero_movimiento_leido), vbOKOnly, "ATENCION"
         End If
      End If
   End If 'fin de la venta
   
   If Me.opt_reproceso = True Then
      var_si = MsgBox("¿Desea enviar la mercancia a reproceso?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar el envio de la mercancia a reproceso", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_primera_vez = True
            var_contador = 0
            For var_j = 1 To Me.lv_existencias.ListItems.Count
                Me.lv_existencias.ListItems.Item(var_j).Selected = True
                If CDbl(Me.lv_existencias.selectedItem.SubItems(5)) > 0 Then
                   var_contador = var_contador + 1
                End If
            Next var_j
            If var_contador > 0 Then
               For var_j = 1 To Me.lv_existencias.ListItems.Count
                   lv_existencias.ListItems.Item(var_j).Selected = True
                   If CDbl(Me.lv_existencias.selectedItem.SubItems(5)) > 0 Then
                      txt_codigo = Me.lv_existencias.selectedItem
                      var_costo = CDbl(Me.lv_existencias.selectedItem.SubItems(3))
                      var_precio = CDbl(Me.lv_existencias.selectedItem.SubItems(4))
                      var_descripcion_articulo = Me.lv_existencias.selectedItem.SubItems(1)
                      var_cantidad_leida = CDbl(Me.lv_existencias.selectedItem.SubItems(5))
                      var_equivalencia = ""
                      If Me.lv_existencias.selectedItem.SubItems(8) = "*" Then
                         rsaux.Open "select * from tb_equivalencias_blackout where vcha_Art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                         If Not rsaux.EOF Then
                            var_equivalencia = IIf(IsNull(rsaux!vcha_equ_codigo_equivalente), "", rsaux!vcha_equ_codigo_equivalente)
                         End If
                         rsaux.Close
                      End If
                   
                      If Trim(txt_codigo) <> "" Then
                         bandera_suma = False
                         If var_primera_vez = True Then
                            var_inserta = False
                            txt_referencia = "ENVIO A REPROCESO"
                            var_clave_moneda = "1"
                            var_almacen_Destino = "CAEE"
                            var_clave_movimiento = "SRE"
                            var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, CStr(var_almacen_Destino), CStr(var_clave_movimiento), Now, CDbl(var_numero_folio), 0, "", "", "", CStr(var_almacen_Destino), "", var_clave_usuario_global, fun_NombrePc, 0, "", CStr(txt_referencia), "", "B", "", "", 0, 0, 0, var_clave_moneda, 0)
                            var_numero_folio = var_numero_folio_regreso
                            txt_folio = var_numero_folio
                            var_primera_vez = False
                         End If
                         Cadena = "select * from tb_temporal_salidas with (nolock) where VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
                         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                         If Not rs.EOF Then
                            var_inserta = False
                            var_inserta = TB_TEMPORAL_SALIDAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida)
                            rs.Close
                         Else
                            var_inserta = False
                            rsaux.Open "INSERT INTO TB_TEMPORAL_SALIDAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, VCHA_SAL_CODIGO_EQUIVALENTE) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ", '" + var_equivalencia + "')", cnn, adOpenDynamic, adLockOptimistic
                            rs.Close
                         End If
                         'rsaux.Open "update tb_temporal_salidas set VCHA_SAL_CODIGO_EQUIVALENTE = '" + var_equivalencia + "' where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id= '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + CStr(var_numero_folio) + " and vcha_Art_Articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic

                      End If
                   End If
               Next var_j
               rs.Open "select * from tb_temporal_salidas with (nolock) where VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_SAL_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     rsaux.Open "select max(cast(substring(vcha_lot_lote_id,3,50)as bigint)) from tb_lotes", cnn_cantia, adOpenDynamic, adLockOptimistic
                     var_consecutivo = rsaux(0).Value
                     rsaux.Close
                     var_equivalencia = IIf(IsNull(rs!VCHA_SAL_CODIGO_EQUIVALENTE), "", rs!VCHA_SAL_CODIGO_EQUIVALENTE)
                     If var_equivalencia = "" Then
                        var_equivalencia = rs!vcha_Art_articulo_id
                     End If
                     var_lote = "0_" + CStr(var_consecutivo + 1)
                     rsaux.Open "select * from tb_producto where vcha_pro_producto_id = '" + rs!vcha_Art_articulo_id + "'", cnn_cantia, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        var_talla = IIf(IsNull(rsaux!vcha_pro_talla), "", rsaux!vcha_pro_talla)
                     Else
                        var_talla = ""
                     End If
                     rsaux.Close
                     
                     If Me.opt_devolucion_cliente = True Then
                        var_cadena = "insert into tb_lotes (vcha_lot_lote_id, vcha_pro_producto_id, vcha_lot_cantidadpro, vcha_lot_cantidadter, vcha_lot_talla, vcha_lot_notaenvio, vcha_lot_fechalta, bint_mod_modulo_id, vcha_det_status, dtim_aud_fecha, vcha_Aud_maquina, bint_pla_planta_id, vcha_lot_auditado, bint_lot_primeras, bint_lot_segundas, vcha_lot_clasificacion, bint_lot_parcialidad1, bint_lot_parcialidad2, dtim_lot_fechacorte, bint_lot_requisicion, vcha_lot_cantidadcor, vcha_lot_acupado, vcha_lot_ajuste, floa_cto_unit )"
                        var_cadena = var_cadena + " values ('" + var_lote + "','" + var_equivalencia + "', '" + CStr(rs!FLOA_sAL_cANTIDAD) + "','0','" + var_talla + "','0',getdate(),9,'S',getdate(),'" + fun_NombrePc + "',0,'A',                                                                                                                                                                                                          0,0,'',0,0,getdate(),0," + CStr(rs!FLOA_sAL_cANTIDAD) + ",'','2'," + CStr(rs!floa_Sal_costo) + " )"
                     End If
                     If Me.opt_reproceso_interno = True Then
                        var_cadena = "insert into tb_lotes (vcha_lot_lote_id, vcha_pro_producto_id, vcha_lot_cantidadpro, vcha_lot_cantidadter, vcha_lot_talla, vcha_lot_notaenvio, vcha_lot_fechalta, bint_mod_modulo_id, vcha_det_status, dtim_aud_fecha, vcha_Aud_maquina, bint_pla_planta_id, vcha_lot_auditado, bint_lot_primeras, bint_lot_segundas, vcha_lot_clasificacion, bint_lot_parcialidad1, bint_lot_parcialidad2, dtim_lot_fechacorte, bint_lot_requisicion, vcha_lot_cantidadcor, vcha_lot_acupado, vcha_lot_ajuste, floa_cto_unit )"
                        var_cadena = var_cadena + " values ('" + var_lote + "','" + var_equivalencia + "', '" + CStr(rs!FLOA_sAL_cANTIDAD) + "','0','" + var_talla + "','0',getdate(),9,'S',getdate(),'" + fun_NombrePc + "',0,'A',                                                                                                                                                                                                          0,0,'',0,0,getdate(),0," + CStr(rs!FLOA_sAL_cANTIDAD) + ",'','3'," + CStr(rs!floa_Sal_costo) + " )"
                     End If
                     rsaux.Open "update tb_lfolios set bint_lot_folios = bint_lot_folios + 1", cnn_cantia, adOpenDynamic, adLockOptimistic
                     rsaux.Open var_cadena, cnn_cantia, adOpenDynamic, adLockOptimistic
                     rsaux.Open "insert into tb_salidas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_sal_numero, vcha_art_articulo_id, floa_sal_cantidad, floa_sal_costo, floa_sal_precio, inte_sal_año, vcha_sal_lote, VCHA_SAL_CODIGO_EQUIVALENTE) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_articulo_id + "', " + CStr(rs!FLOA_sAL_cANTIDAD) + ", " + CStr(rs!floa_Sal_costo) + " , " + CStr(rs!floa_Sal_precio) + ", 2005,'" + var_lote + "','" + IIf(IsNull(rs!VCHA_SAL_CODIGO_EQUIVALENTE), "", rs!VCHA_SAL_CODIGO_EQUIVALENTE) + "')", cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rs.Close
               rs.Open "UPDATE TB_ENCABEZADO_MOVIMIENTOS SET CHAR_EMO_ESTATUS = 'I', INTE_EMO_BLOQUEADO = 0 WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
               
               
               Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_SALIDA_REPROCESO_2.rpt")
               reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_SALIDA.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_MOVIMIENTOS_SALIDA.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_Destino + "' AND {VW_MOVIMIENTOS_SALIDA.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_MOVIMIENTOS_SALIDA.INTE_EMO_NUMERO} = " + Str(var_numero_folio)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte de Movimientos"
               frmvistasprevias.Show 1
               Set reporte = Nothing
               
               Me.frm_cliente.Visible = False
               Me.frm_planta.Visible = False
               Me.frm_reproceso.Visible = True
      
               Me.txt_cliente = ""
               Me.txt_nombre_cliente = ""
               Me.lv_existencias.ListItems.Clear
               
               rs.Open "SELECT dbo.TB_EXISTENCIAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_EXISTENCIAS.FLOA_EXI_CANTIDAD, dbo.TB_EXISTENCIAS.FLOA_EXI_COSTO , dbo.TB_Articulos.MONE_ART_PRECIO_BASE, dbo.TB_EXISTENCIAS.VCHA_ALM_ALMACEN_ID FROM dbo.TB_EXISTENCIAS INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_EXISTENCIAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_EXISTENCIAS.VCHA_ALM_ALMACEN_ID = 'CAEE') and floa_exi_cantidad > 0", cnn, adOpenDynamic, adLockOptimistic
               lv_existencias.ListItems.Clear
               While Not rs.EOF
                     Set list_item = lv_existencias.ListItems.Add(, , rs!vcha_Art_articulo_id)
                     list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
                     list_item.SubItems(2) = IIf(IsNull(rs!floa_Exi_Cantidad), 0, rs!floa_Exi_Cantidad)
                     If IIf(IsNull(rs!FLOA_eXI_COSTO), 0, rs!FLOA_eXI_COSTO) = 0 Then
                        rsaux.Open "select * from tb_articulos where vcha_Art_Articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           list_item.SubItems(3) = IIf(IsNull(rsaux!mone_Art_costo_estandar), "", rsaux!mone_Art_costo_estandar)
                        End If
                        rsaux.Close
                     Else
                        list_item.SubItems(3) = IIf(IsNull(rs!FLOA_eXI_COSTO), "", rs!FLOA_eXI_COSTO)
                     End If
                     list_item.SubItems(4) = IIf(IsNull(rs!mone_art_precio_base), "", rs!mone_art_precio_base)
                     list_item.SubItems(5) = 0
                     rs.MoveNext
               Wend
               rs.Close
               If lv_existencias.ListItems.Count < 21 Then
                  lv_existencias.ColumnHeaders(2).Width = 3449.76
               Else
                  lv_existencias.ColumnHeaders(2).Width = 3200.76
               End If
               
               
               
            Else
               MsgBox "No se a seleccionado articulos para reproceso", vbOKOnly, "ATENCION"
            End If
         End If
      End If
   End If
   
End Sub

Private Sub cmd_nuevo_Click()
   var_estatus_movimiento = "I"
   Me.opt_reproceso = False
   Me.opt_venta = False
   Me.txt_cliente = ""
   Me.txt_nombre_cliente = ""
   Me.txt_establecimiento = ""
   Me.txt_nombre_establecimiento = ""
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   For var_j = 1 To lv_existencias.ListItems.Count
       Me.lv_existencias.ListItems.Item(var_j).Selected = True
       If Me.lv_existencias.selectedItem.SubItems(6) = "*" Then
          Me.lv_existencias.selectedItem.SubItems(5) = Me.lv_existencias.selectedItem.SubItems(2)
          lv_existencias.selectedItem.SubItems(6) = ""
          lv_existencias.ListItems.Item(var_j).Bold = False
          lv_existencias.ListItems.Item(var_j).ForeColor = &H80000012
          lv_existencias.ListItems.Item(var_j).ListSubItems(1).Bold = False
          lv_existencias.ListItems.Item(var_j).ListSubItems(2).Bold = False
          lv_existencias.ListItems.Item(var_j).ListSubItems(3).Bold = False
          lv_existencias.ListItems.Item(var_j).ListSubItems(4).Bold = False
          lv_existencias.ListItems.Item(var_j).ListSubItems(5).Bold = False
          lv_existencias.ListItems.Item(var_j).ListSubItems(1).ForeColor = &H80000012
          lv_existencias.ListItems.Item(var_j).ListSubItems(2).ForeColor = &H80000012
          lv_existencias.ListItems.Item(var_j).ListSubItems(3).ForeColor = &H80000012
          lv_existencias.ListItems.Item(var_j).ListSubItems(4).ForeColor = &H80000012
          lv_existencias.ListItems.Item(var_j).ListSubItems(5).ForeColor = &H80000012
       End If
   Next var_j
End Sub

Private Sub Command10_Click()
   var_todos_lineas = 1
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_existencias.ListItems.Count
   For i = 1 To n
       If lv_existencias.ListItems.Item(i).SubItems(7) <> "0" Then
          lv_existencias.ListItems.Item(i).SubItems(6) = "*"
          lv_existencias.ListItems.Item(i).Bold = True
          lv_existencias.ListItems.Item(i).ForeColor = &HFF0000
          lv_existencias.ListItems.Item(i).ListSubItems(1).Bold = True
          lv_existencias.ListItems.Item(i).ListSubItems(2).Bold = True
          lv_existencias.ListItems.Item(i).ListSubItems(3).Bold = True
          lv_existencias.ListItems.Item(i).ListSubItems(4).Bold = True
          lv_existencias.ListItems.Item(i).ListSubItems(5).Bold = True
          lv_existencias.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
          lv_existencias.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
          lv_existencias.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
          lv_existencias.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
          lv_existencias.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
       End If
   Next
   lv_existencias.Refresh
End Sub

Private Sub Command6_Click()
   If var_todos_lineas = 1 Then
   Else
         var_todos_lineas = 0
   End If
   n = lv_existencias.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_existencias.ListItems.Item(i).Selected = True
      If lv_existencias.selectedItem.SubItems(7) <> "0" Then
         If var_encontro = True And lv_existencias.selectedItem.SubItems(6) = "" And var_rellena = True Then
            lv_existencias.selectedItem.SubItems(6) = "*"
            lv_existencias.ListItems.Item(i).Bold = True
            lv_existencias.ListItems.Item(i).ForeColor = &HFF0000
            lv_existencias.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_existencias.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_existencias.ListItems.Item(i).ListSubItems(3).Bold = True
            lv_existencias.ListItems.Item(i).ListSubItems(4).Bold = True
            lv_existencias.ListItems.Item(i).ListSubItems(5).Bold = True
            lv_existencias.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
            lv_existencias.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
            lv_existencias.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
            lv_existencias.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
            lv_existencias.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         Else
            If var_encontro = True And lv_existencias.selectedItem.SubItems(6) = "*" Then
               var_rellena = False
            End If
         End If
         If lv_existencias.selectedItem.SubItems(6) = "*" And var_encontro = False Then
            var_encontro = True
         End If
      End If
   Next i
End Sub

Private Sub Command7_Click()
   var_todos_lineas = 0
   i = lv_existencias.selectedItem.Index
   If lv_existencias.selectedItem.SubItems(7) <> "0" Then
      If lv_existencias.selectedItem.SubItems(6) = "*" Then
         lv_existencias.selectedItem.SubItems(6) = ""
         lv_existencias.ListItems.Item(i).Bold = False
         lv_existencias.ListItems.Item(i).ForeColor = &H80000012
         lv_existencias.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_existencias.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_existencias.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_existencias.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_existencias.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_existencias.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_existencias.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_existencias.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_existencias.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_existencias.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_existencias.Refresh
      Else
         lv_existencias.selectedItem.SubItems(6) = "*"
         lv_existencias.ListItems.Item(i).Bold = True
         lv_existencias.ListItems.Item(i).ForeColor = &HFF0000
         lv_existencias.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_existencias.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_existencias.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_existencias.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_existencias.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_existencias.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_existencias.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_existencias.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_existencias.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_existencias.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_existencias.Refresh
      End If
   End If
End Sub

Private Sub Command8_Click()
   If var_todos_lineas = 1 Then
   Else
        var_todos_lineas = 0
   End If
   n = lv_existencias.ListItems.Count
   For i = 1 To n
      lv_existencias.ListItems.Item(i).Selected = True
      If lv_existencias.selectedItem.SubItems(7) <> "0" Then
         If lv_existencias.selectedItem.SubItems(6) = "*" Then
            lv_existencias.selectedItem.SubItems(6) = ""
            lv_existencias.ListItems.Item(i).Bold = False
            lv_existencias.ListItems.Item(i).ForeColor = &H80000012
            lv_existencias.ListItems.Item(i).ListSubItems(1).Bold = False
            lv_existencias.ListItems.Item(i).ListSubItems(2).Bold = False
            lv_existencias.ListItems.Item(i).ListSubItems(3).Bold = False
            lv_existencias.ListItems.Item(i).ListSubItems(4).Bold = False
            lv_existencias.ListItems.Item(i).ListSubItems(5).Bold = False
            lv_existencias.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
            lv_existencias.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
            lv_existencias.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
            lv_existencias.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
            lv_existencias.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         Else
            lv_existencias.selectedItem.SubItems(6) = "*"
            lv_existencias.ListItems.Item(i).Bold = True
            lv_existencias.ListItems.Item(i).ForeColor = &HFF0000
            lv_existencias.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_existencias.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_existencias.ListItems.Item(i).ListSubItems(3).Bold = True
            lv_existencias.ListItems.Item(i).ListSubItems(4).Bold = True
            lv_existencias.ListItems.Item(i).ListSubItems(5).Bold = True
            lv_existencias.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
            lv_existencias.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
            lv_existencias.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
            lv_existencias.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
            lv_existencias.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         End If
      End If
   Next i
End Sub

Private Sub Command9_Click()
   var_todos_lineas = 0
   n = lv_existencias.ListItems.Count
   For i = 1 To n
      lv_existencias.ListItems.Item(i).Selected = True
      If Me.lv_existencias.selectedItem.SubItems(7) <> "0" Then
         lv_existencias.selectedItem.SubItems(6) = ""
         lv_existencias.ListItems.Item(i).Bold = False
         lv_existencias.ListItems.Item(i).ForeColor = &H80000012
         lv_existencias.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_existencias.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_existencias.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_existencias.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_existencias.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_existencias.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_existencias.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_existencias.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_existencias.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_existencias.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      End If
   Next i
   lv_existencias.Refresh
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 1000
   Me.frm_lista.Visible = False
   Me.opt_venta = False
   Me.opt_reproceso = False
   
   Me.frm_cliente.Visible = False
   Me.frm_planta.Visible = False
   Me.frm_reproceso.Visible = True
   Set cnn_cantia = CreateObject("ADODB.connection")
   rs.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_conexion_cantia = IIf(IsNull(rs!vcha_uor_conexion), "", rs!vcha_uor_conexion)
   Else
      var_conexion_cantia = ""
   End If
   rs.Close
   If var_conexion_cantia <> "" Then
      cnn_cantia.Open var_conexion_cantia
   End If
   frm_eliminar.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_entradas_sin_comparacion)
End Sub

Private Sub lv_existencias_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_existencias, ColumnHeader)
End Sub

Private Sub lv_existencias_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      Me.txt_cantidad_eliminar = ""
      frm_eliminar.Visible = True
      txt_cantidad_eliminar.SetFocus
   End If
   If KeyCode = 116 Then
      If Me.lv_existencias.ListItems.Count > 0 Then
         If Me.lv_existencias.selectedItem.SubItems(7) = "0" Then
            MsgBox "No se puede seleccionar este artículo ya que su precio es 0", vbOKOnly, "ATENCION"
         Else
            var_todos_lineas = 0
            i = lv_existencias.selectedItem.Index
            If lv_existencias.selectedItem.SubItems(8) = "*" Then
               lv_existencias.selectedItem.SubItems(8) = ""
               lv_existencias.ListItems.Item(i).Bold = False
               lv_existencias.ListItems.Item(i).ForeColor = &H80000012
               lv_existencias.ListItems.Item(i).ListSubItems(1).Bold = False
               lv_existencias.ListItems.Item(i).ListSubItems(2).Bold = False
               lv_existencias.ListItems.Item(i).ListSubItems(3).Bold = False
               lv_existencias.ListItems.Item(i).ListSubItems(4).Bold = False
               lv_existencias.ListItems.Item(i).ListSubItems(5).Bold = False
               lv_existencias.ListItems.Item(i).ListSubItems(6).Bold = False
               lv_existencias.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
               lv_existencias.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
               lv_existencias.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
               lv_existencias.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
               lv_existencias.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
               lv_existencias.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
               lv_existencias.Refresh
            Else
               lv_existencias.selectedItem.SubItems(8) = "*"
               lv_existencias.ListItems.Item(i).Bold = True
               lv_existencias.ListItems.Item(i).ForeColor = &HFF00FF
               lv_existencias.ListItems.Item(i).ListSubItems(1).Bold = True
               lv_existencias.ListItems.Item(i).ListSubItems(2).Bold = True
               lv_existencias.ListItems.Item(i).ListSubItems(3).Bold = True
               lv_existencias.ListItems.Item(i).ListSubItems(4).Bold = True
               lv_existencias.ListItems.Item(i).ListSubItems(5).Bold = True
               lv_existencias.ListItems.Item(i).ListSubItems(6).Bold = True
               lv_existencias.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF00FF
               lv_existencias.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF00FF
               lv_existencias.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF00FF
               lv_existencias.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF00FF
               lv_existencias.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF00FF
               lv_existencias.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF00FF
               lv_existencias.Refresh
            End If
         End If
      End If
   End If
   
End Sub

Private Sub lv_existencias_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_existencias.ListItems.Count > 0 Then
         If Me.lv_existencias.selectedItem.SubItems(7) = "0" Then
            MsgBox "No se puede seleccionar este artículo ya que su precio es 0", vbOKOnly, "ATENCION"
         Else
            var_todos_lineas = 0
            i = lv_existencias.selectedItem.Index
            If lv_existencias.selectedItem.SubItems(6) = "*" Then
               lv_existencias.selectedItem.SubItems(6) = ""
               lv_existencias.ListItems.Item(i).Bold = False
               lv_existencias.ListItems.Item(i).ForeColor = &H80000012
               lv_existencias.ListItems.Item(i).ListSubItems(1).Bold = False
               lv_existencias.ListItems.Item(i).ListSubItems(2).Bold = False
               lv_existencias.ListItems.Item(i).ListSubItems(3).Bold = False
               lv_existencias.ListItems.Item(i).ListSubItems(4).Bold = False
               lv_existencias.ListItems.Item(i).ListSubItems(5).Bold = False
               lv_existencias.ListItems.Item(i).ListSubItems(6).Bold = False
               lv_existencias.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
               lv_existencias.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
               lv_existencias.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
               lv_existencias.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
               lv_existencias.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
               lv_existencias.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
               lv_existencias.Refresh
            Else
               lv_existencias.selectedItem.SubItems(6) = "*"
               lv_existencias.ListItems.Item(i).Bold = True
               lv_existencias.ListItems.Item(i).ForeColor = &HFF0000
               lv_existencias.ListItems.Item(i).ListSubItems(1).Bold = True
               lv_existencias.ListItems.Item(i).ListSubItems(2).Bold = True
               lv_existencias.ListItems.Item(i).ListSubItems(3).Bold = True
               lv_existencias.ListItems.Item(i).ListSubItems(4).Bold = True
               lv_existencias.ListItems.Item(i).ListSubItems(5).Bold = True
               lv_existencias.ListItems.Item(i).ListSubItems(6).Bold = True
               lv_existencias.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
               lv_existencias.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
               lv_existencias.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
               lv_existencias.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
               lv_existencias.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
               lv_existencias.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
               lv_existencias.Refresh
            End If
         End If
      End If
   End If
End Sub

Private Sub Option2_Click()

End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If Me.lv_lista.ListItems.Count > 0 Then
      If var_tipo_lista = 1 Then
         Me.txt_cliente = lv_lista.selectedItem
         Me.txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
         Me.txt_cliente.SetFocus
      End If
      If var_tipo_lista = 2 Then
         Me.txt_establecimiento = lv_lista.selectedItem
         Me.txt_nombre_establecimiento = lv_lista.selectedItem.SubItems(1)
         Me.txt_establecimiento.SetFocus
      End If
      If var_tipo_lista = 3 Then
         Me.txt_planta = lv_lista.selectedItem
         Me.txt_nombre_planta = lv_lista.selectedItem.SubItems(1)
         Me.txt_planta.SetFocus
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub opt_reproceso_Click()
   If Me.opt_reproceso.Value = True Then
      Me.opt_devolucion_cliente = True
      Me.frm_cliente.Visible = False
      Me.frm_planta.Visible = False
      Me.frm_reproceso.Visible = True
      
      Me.txt_cliente = ""
      Me.txt_nombre_cliente = ""
      Me.lv_existencias.ListItems.Clear
      rs.Open "SELECT dbo.TB_EXISTENCIAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_EXISTENCIAS.FLOA_EXI_CANTIDAD, dbo.TB_EXISTENCIAS.FLOA_EXI_COSTO , dbo.TB_Articulos.MONE_ART_PRECIO_BASE, dbo.TB_EXISTENCIAS.VCHA_ALM_ALMACEN_ID FROM dbo.TB_EXISTENCIAS INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_EXISTENCIAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_EXISTENCIAS.VCHA_ALM_ALMACEN_ID = 'CAEE') and floa_exi_cantidad > 0", cnn, adOpenDynamic, adLockOptimistic
      lv_existencias.ListItems.Clear
      While Not rs.EOF
            Set list_item = lv_existencias.ListItems.Add(, , rs!vcha_Art_articulo_id)
            list_item.SubItems(1) = UCase(IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español))
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_Exi_Cantidad), 0, rs!floa_Exi_Cantidad), "###,###,##0.00")
            If IIf(IsNull(rs!FLOA_eXI_COSTO), 0, rs!FLOA_eXI_COSTO) = 0 Then
               rsaux.Open "select * from tb_articulos where vcha_Art_Articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  list_item.SubItems(3) = Format(IIf(IsNull(rsaux!mone_Art_costo_estandar), "", rsaux!mone_Art_costo_estandar), "###,###,##0.00")
               End If
               rsaux.Close
            Else
               list_item.SubItems(3) = Format(IIf(IsNull(rs!FLOA_eXI_COSTO), "", rs!FLOA_eXI_COSTO), "###,###,##0.00")
            End If
            list_item.SubItems(4) = Format(IIf(IsNull(rs!mone_art_precio_base), "", rs!mone_art_precio_base), "###,###,##0.00")
            list_item.SubItems(5) = 0
            rs.MoveNext
      Wend
      rs.Close
      If lv_existencias.ListItems.Count < 21 Then
         lv_existencias.ColumnHeaders(2).Width = 3449.76
      Else
         lv_existencias.ColumnHeaders(2).Width = 3200.76
      End If

   End If
End Sub

Private Sub opt_traspaso_Click()
   If opt_traspaso = True Then
      Me.frm_cliente.Visible = False
      Me.frm_planta.Visible = True
      Me.frm_reproceso.Visible = False
      Me.lv_existencias.ListItems.Clear
      rs.Open "SELECT dbo.TB_EXISTENCIAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_EXISTENCIAS.FLOA_EXI_CANTIDAD, dbo.TB_EXISTENCIAS.FLOA_EXI_COSTO , dbo.TB_Articulos.MONE_ART_PRECIO_BASE, dbo.TB_EXISTENCIAS.VCHA_ALM_ALMACEN_ID FROM dbo.TB_EXISTENCIAS INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_EXISTENCIAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_EXISTENCIAS.VCHA_ALM_ALMACEN_ID = 'CAEE') and floa_exi_cantidad > 0", cnn, adOpenDynamic, adLockOptimistic
      lv_existencias.ListItems.Clear
      While Not rs.EOF
            Set list_item = lv_existencias.ListItems.Add(, , rs!vcha_Art_articulo_id)
            list_item.SubItems(1) = UCase(IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español))
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_Exi_Cantidad), 0, rs!floa_Exi_Cantidad), "###,###,##0.00")
            If IIf(IsNull(rs!FLOA_eXI_COSTO), 0, rs!FLOA_eXI_COSTO) = 0 Then
               rsaux.Open "select * from tb_articulos where vcha_Art_Articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  list_item.SubItems(3) = Format(IIf(IsNull(rsaux!mone_Art_costo_estandar), "", rsaux!mone_Art_costo_estandar), "###,###,##0.00")
               End If
               rsaux.Close
            Else
               list_item.SubItems(3) = Format(IIf(IsNull(rs!FLOA_eXI_COSTO), 0, rs!FLOA_eXI_COSTO), "###,###,##0.00")
            End If
            list_item.SubItems(4) = Format(IIf(IsNull(rs!mone_art_precio_base), 0, rs!mone_art_precio_base), "###,###,##0.00")
            list_item.SubItems(5) = 0
            rs.MoveNext
      Wend
      rs.Close
      
      If lv_existencias.ListItems.Count < 21 Then
         lv_existencias.ColumnHeaders(2).Width = 3449.76
      Else
         lv_existencias.ColumnHeaders(2).Width = 3200.76
      End If
      
   End If
End Sub

Private Sub opt_venta_Click()
   If Me.opt_venta.Value = True Then
      Me.lv_existencias.ListItems.Clear
      Me.txt_cliente.Enabled = True
      Me.txt_nombre_cliente.Enabled = True
      Me.txt_cliente = ""
      Me.txt_nombre_cliente = ""
      Me.txt_establecimiento = ""
      Me.txt_nombre_establecimiento = ""
      'Me.txt_cliente.SetFocus
      Me.frm_cliente.Visible = True
      Me.frm_planta.Visible = False
      Me.frm_reproceso.Visible = False
   End If
End Sub

Private Sub Text4_Change()

End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub txt_busqueda_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_existencias, Me.txt_busqueda, False)
   End If
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad_eliminar) Then
         If CDbl(Me.txt_cantidad_eliminar) <= Me.lv_existencias.selectedItem.SubItems(2) Then
            Me.lv_existencias.selectedItem.SubItems(5) = Me.txt_cantidad_eliminar
         Else
            MsgBox "La cantidad a pasar no debe exceder a la cantidad disponible", vbOKOnly, "ATENCION"
         End If
         Me.lv_existencias.SetFocus
      Else
         MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
         Me.lv_existencias.SetFocus
      End If
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   Me.frm_eliminar.Visible = False
End Sub

Private Sub txt_cliente_Change()
   Me.lv_existencias.ListItems.Clear
   Me.txt_nombre_cliente = ""
   Me.txt_nombre_establecimiento = ""
   Me.txt_establecimiento = ""
End Sub

Private Sub txt_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select vcha_cli_clave_id, vcha_cli_nombre from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
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

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_cliente_LostFocus()
   If Me.txt_cliente <> "" Then
      rs.Open "select * from vw_clientes where vcha_Cli_clave_id = '" + Me.txt_cliente + "' and vcha_Emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         Me.lv_existencias.ListItems.Clear
         rsaux2.Open "SELECT dbo.TB_EXISTENCIAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_EXISTENCIAS.FLOA_EXI_CANTIDAD, dbo.TB_EXISTENCIAS.FLOA_EXI_COSTO , dbo.TB_Articulos.MONE_ART_PRECIO_BASE, dbo.TB_EXISTENCIAS.VCHA_ALM_ALMACEN_ID FROM dbo.TB_EXISTENCIAS INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_EXISTENCIAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_EXISTENCIAS.VCHA_ALM_ALMACEN_ID = 'CAEE') and floa_exi_cantidad > 0", cnn, adOpenDynamic, adLockOptimistic
         lv_existencias.ListItems.Clear
         While Not rsaux2.EOF
               Set list_item = lv_existencias.ListItems.Add(, , rsaux2!vcha_Art_articulo_id)
               list_item.SubItems(1) = UCase(IIf(IsNull(rsaux2!vcha_art_nombre_español), "", rsaux2!vcha_art_nombre_español))
               list_item.SubItems(2) = IIf(IsNull(rsaux2!floa_Exi_Cantidad), 0, rsaux2!floa_Exi_Cantidad)
               If IIf(IsNull(rsaux2!FLOA_eXI_COSTO), 0, rsaux2!FLOA_eXI_COSTO) = 0 Then
                  rsaux.Open "select * from tb_articulos where vcha_Art_Articulo_id = '" + rsaux2!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     list_item.SubItems(3) = IIf(IsNull(rsaux!mone_Art_costo_estandar), "", rsaux!mone_Art_costo_estandar)
                  End If
                  rsaux.Close
               Else
                  list_item.SubItems(3) = IIf(IsNull(rsaux2!FLOA_eXI_COSTO), "", rsaux2!FLOA_eXI_COSTO)
               End If
               rsaux.Open "select * from tb_Detalle_lista_precios where vcha_art_articulo_id = '" + rsaux2!vcha_Art_articulo_id + "' and vcha_lis_lista_precios_id = '" + rs!vcha_LIS_LISTA_iD + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  list_item.SubItems(4) = IIf(IsNull(rsaux!floa_dli_precio), 0, rsaux!floa_dli_precio)
               Else
                  list_item.SubItems(4) = 0
               End If
               rsaux.Close
               list_item.SubItems(5) = 0
               rsaux2.MoveNext
         Wend
         rsaux2.Close
         For i = 1 To lv_existencias.ListItems.Count
             Me.lv_existencias.ListItems(i).Selected = True
             If Me.lv_existencias.selectedItem.SubItems(4) = "0" Then
                lv_existencias.ListItems.Item(i).Bold = True
                lv_existencias.ListItems.Item(i).ForeColor = &HFF&
                lv_existencias.ListItems.Item(i).ListSubItems(1).Bold = True
                lv_existencias.ListItems.Item(i).ListSubItems(2).Bold = True
                lv_existencias.ListItems.Item(i).ListSubItems(3).Bold = True
                lv_existencias.ListItems.Item(i).ListSubItems(4).Bold = True
                lv_existencias.ListItems.Item(i).ListSubItems(5).Bold = True
                lv_existencias.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF&
                lv_existencias.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF&
                lv_existencias.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF&
                lv_existencias.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF&
                lv_existencias.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF&
                lv_existencias.selectedItem.SubItems(7) = "0"
                lv_existencias.Refresh
             End If
         
         
         
         
         Next i
         
         
         Me.txt_cliente.Enabled = False
         Me.txt_nombre_cliente.Enabled = False
         
         If lv_existencias.ListItems.Count < 21 Then
            lv_existencias.ColumnHeaders(2).Width = 3449.76
         Else
            lv_existencias.ColumnHeaders(2).Width = 3200.76
         End If
         
         
         
      Else
         Me.txt_nombre_cliente = ""
         Me.lv_existencias.ListItems.Clear
      End If
      rs.Close
   Else
      Me.lv_existencias.ListItems.Clear
   End If
End Sub

Private Sub txt_establecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select vcha_Esb_establecimiento_id, vcha_esb_nombre from vw_Establecimientos where vcha_cli_clave_id = '" + Me.txt_cliente + "' order by vcha_esb_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ESB_ESTABLECIMIENTO_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Establecimientos"
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

Private Sub txt_establecimiento_LostFocus()
   If Me.txt_establecimiento <> "" Then
      rs.Open "SELECT * FROM VW_eSTABLECIMIENTOS WHERE VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "' AND VCHA_ESB_eSTABLECIMIENTO_ID = '" + Me.txt_establecimiento + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_establecimiento = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
      Else
         MsgBox "Clave de establecimiento incorrecta", vbOKOnly, "ATENCION"
         Me.txt_establecimiento = ""
         Me.txt_nombre_establecimiento = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_establecimiento = ""
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_planta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_numero_reproceso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_numero_reproceso) Then
         var_almacen_Destino = "CAEE"
         var_clave_movimiento = "SRE"
         Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_SALIDA_REPROCESO_2.rpt")
         reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_SALIDA.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_MOVIMIENTOS_SALIDA.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_Destino + "' AND {VW_MOVIMIENTOS_SALIDA.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_MOVIMIENTOS_SALIDA.INTE_EMO_NUMERO} = " + Me.txt_numero_reproceso
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
            reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Movimientos"
         frmvistasprevias.Show 1
         Set reporte = Nothing
      Else
         MsgBox "Número de movimiento incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_numero_traspaso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_numero_traspaso) Then
         var_almacen_Destino = "CAEE"
         var_clave_movimiento = "DPL"
         Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_SALIDA.rpt")
         reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_SALIDA.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_Destino + "' AND {VW_MOVIMIENTOS_SALIDA.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_MOVIMIENTOS_SALIDA.INTE_EMO_NUMERO} = " + Me.txt_numero_traspaso
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Movimientos"
         frmvistasprevias.Show 1
         Set reporte = Nothing
      Else
         MsgBox "Número de traspaso incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_planta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select vcha_uor_unidad_id, vcha_uor_nombre from tb_unidadesorganizacionales where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_uor_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_uor_unidad_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_UOR_NOMBRE), "", rs!VCHA_UOR_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Plantas"
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
End Sub

Private Sub txt_planta_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_planta_LostFocus()
   If Me.txt_planta <> "" Then
      rs.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + Me.txt_planta + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_planta = IIf(IsNull(rs!VCHA_UOR_NOMBRE), "", rs!VCHA_UOR_NOMBRE)
      Else
         MsgBox "Clave de planta incorrecta", vbOKOnly, "ATENCION"
         Me.txt_planta = ""
         Me.txt_nombre_planta = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_planta = ""
   End If
End Sub
