VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsalidas_crossdocking 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salidas crossdocking"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Height          =   315
      Left            =   120
      Picture         =   "frmsalidas_crossdocking.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Desmarcar Todos Alt + D"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   2430
      TabIndex        =   16
      Top             =   900
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   17
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
         TabIndex        =   18
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   450
      Picture         =   "frmsalidas_crossdocking.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8910
      Picture         =   "frmsalidas_crossdocking.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   120
      Left            =   75
      TabIndex        =   13
      Top             =   270
      Width           =   9195
   End
   Begin VB.Frame Frame1 
      Height          =   6885
      Left            =   75
      TabIndex        =   11
      Top             =   345
      Width           =   9225
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1710
         Picture         =   "frmsalidas_crossdocking.frx":0886
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Pasar todos F4"
         Top             =   150
         Width           =   330
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   5040
         TabIndex        =   19
         Top             =   3375
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   60
            TabIndex        =   20
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
            TabIndex        =   21
            Top             =   15
            Width           =   2895
         End
      End
      Begin VB.TextBox txt_nombre_planta 
         Height          =   315
         Left            =   4560
         TabIndex        =   9
         Top             =   210
         Width           =   4590
      End
      Begin VB.TextBox txt_planta 
         Height          =   315
         Left            =   3465
         TabIndex        =   8
         Top             =   210
         Width           =   1065
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmsalidas_crossdocking.frx":0988
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   150
         Width           =   330
      End
      Begin VB.CommandButton Command9 
         Height          =   315
         Left            =   60
         Picture         =   "frmsalidas_crossdocking.frx":0B9E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   150
         Width           =   330
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         Picture         =   "frmsalidas_crossdocking.frx":0CA0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   150
         Width           =   330
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         Picture         =   "frmsalidas_crossdocking.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Marcar (Enter)"
         Top             =   150
         Width           =   330
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1380
         Picture         =   "frmsalidas_crossdocking.frx":0FBC
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   150
         Width           =   330
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   30
         TabIndex        =   12
         Top             =   570
         Width           =   9180
      End
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   6135
         Left            =   45
         TabIndex        =   10
         Top             =   645
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   10821
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
            Object.Width           =   5997
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
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Planta:"
         Height          =   195
         Left            =   2865
         TabIndex        =   14
         Top             =   255
         Width           =   495
      End
   End
   Begin VB.Label lbl_numero_entrada 
      Height          =   285
      Left            =   1005
      TabIndex        =   15
      Top             =   105
      Visible         =   0   'False
      Width           =   1065
   End
End
Attribute VB_Name = "frmsalidas_crossdocking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_aceptar_pedidos_Click()
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
   Dim var_almacen_destino_cross As String
   Dim var_clave_movimiento_cross As String
   Dim var_clave_moneda As String
   Dim txt_proveedor As String
   Dim txt_nombre_proveedor As String
   Dim var_numero_folio_cross  As Double
   'Dim var_almacen_Destino_cross As String
   'Dim var_clave_movimiento_cross As String
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
   
   var_posible_articulos = False
   If Me.txt_planta <> "" Then
      If Me.lv_articulos.ListItems.Count > 0 Then
         For var_j = 1 To Me.lv_articulos.ListItems.Count
             Me.lv_articulos.ListItems.Item(var_j).Selected = True
             If Me.lv_articulos.selectedItem.SubItems(5) <> "" Then
                If CDbl(Me.lv_articulos.selectedItem.SubItems(5)) > 0 Then
                   var_posible_articulos = True
                End If
             End If
         Next var_j
         If var_posible_articulos = True Then
            var_si = MsgBox("¿Desea cerrar el movimiento?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_si = MsgBox("Confirmar el cerrado del movimiento", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  var_primera_vez = True
                  If var_empresa = "06" Then
                     var_almacen_destino_cross = "Q0Z"
                  End If
                  If var_empresa = "17" Then
                     var_almacen_destino_cross = "BORPT"
                  End If
                  var_clave_movimiento_cross = "DPL"
                  var_clave_moneda = "1"
                  txt_proveedor = Me.txt_planta
                  txt_nombre_proveedor = Me.txt_nombre_planta
                  For var_j = 1 To Me.lv_articulos.ListItems.Count
                      Me.lv_articulos.ListItems.Item(var_j).Selected = True
                      If Trim(Me.lv_articulos.selectedItem.SubItems(5)) = "" Then
                         Me.lv_articulos.selectedItem.SubItems(5) = "0"
                      End If
                      If CDbl(Me.lv_articulos.selectedItem.SubItems(5)) > 0 Then
                         txt_codigo = Me.lv_articulos.selectedItem
                         var_cantidad_leida = CDbl(Me.lv_articulos.selectedItem.SubItems(5))
                         var_costo = CDbl(Me.lv_articulos.selectedItem.SubItems(3))
                         var_precio = CDbl(Me.lv_articulos.selectedItem.SubItems(4))
                         If Trim(txt_codigo) <> "" Then
                            If var_primera_vez = True Then
                               var_inserta = False
                               var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino_cross, var_clave_movimiento_cross, Now, var_numero_folio_cross, 0, "", txt_proveedor, "", var_almacen_destino_cross, "", var_clave_usuario_global, fun_NombrePc, 0, "", txt_nombre_proveedor, "", "B", "", "", 0, 0, 0, var_clave_moneda, 0)
                               var_numero_folio_cross = var_numero_folio_regreso
                               var_primera_vez = False
                            End If
                            var_posible_leido = 1
                            If var_posible_leido = 1 Then
                               Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_destino_cross + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento_cross + "' and inte_sal_numero = " + CStr(var_numero_folio_cross) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
                               rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                               If Not rs.EOF Then
                                  var_inserta = False
                                  var_inserta = TB_TEMPORAL_SALIDAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino_cross, var_clave_movimiento_cross, CDbl(var_numero_folio_cross), CStr(txt_codigo), CDbl(var_cantidad_leida))
                                  rs.Close
                               Else
                                  var_inserta = False
                                  rsaux.Open "INSERT INTO TB_TEMPORAL_SALIDAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_destino_cross + "', '" + var_clave_movimiento_cross + "', " + CStr(var_numero_folio_cross) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ")", cnn, adOpenDynamic, adLockOptimistic
                                  rs.Close
                               End If
                            End If
                         End If
                      End If
                  Next var_j
                  
                  
                  
                  
                  Dim var_posible_Cantidad As Integer
                  var_posible_Cantidad = 1
                  var_cadena_articulos = ""
                  If var_empresa = "06" Or var_empresa = "17" Or var_empresa = "18" Then
                     rsaux10.Open "select * from tb_plantas where vcha_uor_unidad_id = '" + txt_planta + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux10.EOF Then
                        var_clave_planta_destino = IIf(IsNull(rsaux10!vcha_pla_planta_id), "", rsaux10!vcha_pla_planta_id)
                        VAR_NOMBRE_PLANTA_DESTINO = IIf(IsNull(rsaux10!vcha_pla_descripc), "", rsaux10!vcha_pla_descripc)
                     End If
                     rsaux10.Close
                     If rsaux10.State = 1 Then
                        rsaux10.Close
                     End If
                     rsaux10.Open "select * from tb_plantas where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_clave_planta_origen = IIf(IsNull(rsaux10!vcha_pla_planta_id), "", rsaux10!vcha_pla_planta_id)
                     rsaux10.Close
                  End If
                  
            
                  
                  
                  
                  Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_destino_cross + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento_cross + "' and inte_sal_numero = " + Str(var_numero_folio_cross) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
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
                              
                              If var_empresa = "06" Or var_empresa = "17" Or var_empresa = "18" Then
                                 If rsaux9.State = 1 Then
                                    rsaux9.Close
                                 End If
                                 rsaux9.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                 var_descripcion_articulo = IIf(IsNull(rsaux9!vcha_art_nombre_español), "", rsaux9!vcha_art_nombre_español)
                                 rsaux9.Close
                                 var_cadena = "insert into tb_transito (vcha_tra_nota_envio, vcha_Art_Articulo_id,                                                              vcha_Art_descripcion,           floa_Tra_cantidad_Enviada,                            floa_tra_costo, vcha_tra_planta_origen, vcha_tra_planta_destino, floa_tra_Cantidad_recibida, vcha_tra_Calidad, VCHA_TRA_STATUS, VCHA_MOV_MOVIMIENTO_ID, VCHA_EMP_EMPRESA_ID) "
                                 var_cadena = var_cadena + "   values  ('" + var_clave_planta_origen + "_" + CStr(var_numero_folio_cross) + "', '" + rs!vcha_Art_articulo_id + "','" + var_descripcion_articulo + "', " + CStr(rs!FLOA_sAL_cANTIDAD) + ", " + CStr(var_costo) + ",'" + var_clave_planta_origen + "','" + var_clave_planta_destino + "',0,'1','A','SALTRA', '" + var_empresa + "')"
                                 rsaux9.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              
                              
                              
                        Wend
                        rs.MoveNext
                  Wend
                  rs.Close
                  
                  
                  If var_empresa = "06" Or var_empresa = "17" Then
                     rsaux10.Open "select sum(floa_sal_Cantidad * floa_sal_costo) as costo from tb_salidas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_destino_cross + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento_cross + "' and inte_sal_numero = " + Str(var_numero_folio_cross) + " and floa_Sal_cantidad > 0", cnn, adOpenDynamic, adLockOptimistic
                     txt_folio = Str(var_numero_folio_cross)
                     rsaux11.Open "select * from tb_generador_polizas where poliza_id = '8'", cnnoracle, adOpenDynamic, adLockOptimistic
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
                           'var_precio = rsaux11!Precio
                           If var_precio = 1 Then
                              var_importe_precio = rsaux10!Costo
                           Else
                              var_importe_precio = rsaux10!Costo
                           End If
                           var_cadena = "InsERT INTO IN_TB_POLIZAS_INT (STATUS, SET_OF_BOOKS_ID, USER_JE_SOURCE_NAME, USER_JE_CATEGORY_NAME, ACCOUNTING_DATE, CURRENCY_CODE, DATE_CREATED, ACTUAL_FLAG,  SEGMENT1, SEGMENT2, SEGMENT3, SEGMENT4, SEGMENT5, SEGMENT6, SEGMENT7, ENTERED_DR, ENTERED_CR, ACCOUNTED_DR, ACCOUNTED_CR, GROUP_ID, REFERENCE4, REFERENCE5, REFERENCe10, REFERENCE1, REFERENCE2, CREATED_BY)"
                           If var_cargo_poliza = 1 Then
                              var_cadena = var_cadena + " VALUES ('NEW', " + CStr(var_juego_libros_poliza) + ",'" + var_origen_poliza + "','" + var_categoria_poliza + "',TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'" + var_moneda_poliza + "',TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'A','" + var_segmento1_poliza + "','" + var_segmento2_poliza + "','" + var_segmento3_poliza + "','" + var_segmento4_poliza + "','" + var_segmento5_poliza + "','" + var_segmento6_poliza + "','" + var_segmento7_poliza + "'," + CStr(var_importe_precio) + ",0," + CStr(var_importe_precio) + ",0,1,'SALIDA POR TRASPASO A PLANTAS " + txt_folio + "','DESTINO: " + VAR_NOMBRE_PLANTA_DESTINO + "','" + var_descripcion_poliza + "','POLIZA TRASPASO ENTRE ALMACENES','POLIZA TRASPASO ENTRE ALMACENES',1143)"
                           Else
                              var_cadena = var_cadena + " VALUES ('NEW', " + CStr(var_juego_libros_poliza) + ",'" + var_origen_poliza + "','" + var_categoria_poliza + "',TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'" + var_moneda_poliza + "',TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'A','" + var_segmento1_poliza + "','" + var_segmento2_poliza + "','" + var_segmento3_poliza + "','" + var_segmento4_poliza + "','" + var_segmento5_poliza + "','" + var_segmento6_poliza + "','" + var_segmento7_poliza + "',0," + CStr(var_importe_precio) + ",0," + CStr(var_importe_precio) + ",1,'SALIDA POR TRASPASO A PLANTAS " + txt_folio + "','DESTINO: " + VAR_NOMBRE_PLANTA_DESTINO + "','" + var_descripcion_poliza + "','POLIZA TRASPASO ENTRE ALMACENES','POLIZA TRASPASO ENTRE ALMACENES',1143)"
                           End If
                           'MsgBox var_cadena
                           rsaux9.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                           rsaux11.MoveNext
                     Wend
                     rsaux11.Close
                     rsaux10.Close
                  
                  End If
                  
                  
                  
                  
                  var_estatus_movimiento = "I"
                  var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino_cross, var_clave_movimiento_cross, var_numero_folio_cross, "", Now, 1)
                  var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino_cross, var_clave_movimiento_cross, var_numero_folio_cross, "I", Now, 1)
               
                  Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_SALIDA.rpt")
                  reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_SALIDA.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_destino_cross + "' AND {VW_MOVIMIENTOS_SALIDA.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento_cross + "' AND {VW_MOVIMIENTOS_SALIDA.INTE_EMO_NUMERO} = " + Str(var_numero_folio_cross)
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Movimientos"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1, inte_emo_bloqueado = 0 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_destino_cross + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento_cross + "' and inte_emo_numero = " + CStr(var_numero_folio_cross), cnn, adOpenDynamic, adLockOptimistic
               
               
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
End Sub


Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   Me.txt_planta = ""
   Me.txt_nombre_planta = ""
   var_todos_lineas = 0
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      lv_articulos.selectedItem.SubItems(6) = ""
      lv_articulos.selectedItem.SubItems(5) = ""
      lv_articulos.ListItems.Item(i).Bold = False
      lv_articulos.ListItems.Item(i).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
   Next i
   lv_articulos.Refresh
   Me.txt_planta.SetFocus
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
   n = lv_articulos.ListItems.Count
   For i = 1 To n
       lv_articulos.ListItems.Item(i).SubItems(6) = "*"
       lv_articulos.ListItems.Item(i).Bold = True
       lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
       lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = True
       lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = True
       lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = True
       lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = True
       lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
       lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
       lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
       lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
       lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
       lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
   Next
   lv_articulos.Refresh
End Sub

Private Sub Command2_Click()
   For var_j = 1 To lv_articulos.ListItems.Count
       Me.lv_articulos.ListItems.Item(var_j).Selected = True
       If Me.lv_articulos.selectedItem.SubItems(6) = "*" Then
          Me.lv_articulos.selectedItem.SubItems(5) = Format(Me.lv_articulos.selectedItem.SubItems(2), "###,###,##0.00")
          lv_articulos.selectedItem.SubItems(6) = ""
          lv_articulos.ListItems.Item(var_j).Bold = False
          Me.lv_articulos.selectedItem.SubItems(6) = ""
          lv_articulos.ListItems.Item(var_j).ForeColor = &H80000012
          lv_articulos.ListItems.Item(var_j).ListSubItems(1).Bold = False
          lv_articulos.ListItems.Item(var_j).ListSubItems(2).Bold = False
          lv_articulos.ListItems.Item(var_j).ListSubItems(3).Bold = False
          lv_articulos.ListItems.Item(var_j).ListSubItems(4).Bold = False
          lv_articulos.ListItems.Item(var_j).ListSubItems(5).Bold = False
          lv_articulos.ListItems.Item(var_j).ListSubItems(6).Bold = False
          lv_articulos.ListItems.Item(var_j).ListSubItems(1).ForeColor = &H80000012
          lv_articulos.ListItems.Item(var_j).ListSubItems(2).ForeColor = &H80000012
          lv_articulos.ListItems.Item(var_j).ListSubItems(3).ForeColor = &H80000012
          lv_articulos.ListItems.Item(var_j).ListSubItems(4).ForeColor = &H80000012
          lv_articulos.ListItems.Item(var_j).ListSubItems(5).ForeColor = &H80000012
          lv_articulos.ListItems.Item(var_j).ListSubItems(6).ForeColor = &H80000012
          lv_articulos.Refresh
       End If
   Next var_j
End Sub

Private Sub Command6_Click()
   If var_todos_lineas = 1 Then
   Else
         var_todos_lineas = 0
   End If
   n = lv_articulos.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_articulos.selectedItem.SubItems(6) = "" And var_rellena = True Then
         lv_articulos.selectedItem.SubItems(6) = "*"
         lv_articulos.ListItems.Item(i).Bold = True
         lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_articulos.selectedItem.SubItems(6) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_articulos.selectedItem.SubItems(6) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub Command7_Click()
   var_todos_lineas = 0
   i = lv_articulos.selectedItem.Index
   If lv_articulos.selectedItem.SubItems(6) = "*" Then
      lv_articulos.selectedItem.SubItems(6) = ""
      lv_articulos.ListItems.Item(i).Bold = False
      lv_articulos.ListItems.Item(i).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
      lv_articulos.Refresh
   Else
      lv_articulos.selectedItem.SubItems(6) = "*"
      lv_articulos.ListItems.Item(i).Bold = True
      lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
      lv_articulos.Refresh
   End If
End Sub

Private Sub Command8_Click()
   If var_todos_lineas = 1 Then
   Else
        var_todos_lineas = 0
   End If
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      If lv_articulos.selectedItem.SubItems(6) = "*" Then
         lv_articulos.selectedItem.SubItems(6) = ""
         lv_articulos.ListItems.Item(i).Bold = False
         lv_articulos.ListItems.Item(i).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
      Else
         lv_articulos.selectedItem.SubItems(6) = "*"
         lv_articulos.ListItems.Item(i).Bold = True
         lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub Command9_Click()
   var_todos_lineas = 0
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      lv_articulos.selectedItem.SubItems(5) = ""
      lv_articulos.ListItems.Item(i).Bold = False
      lv_articulos.ListItems.Item(i).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
   Next i
   lv_articulos.Refresh
End Sub

Private Sub Form_Load()
   Me.frm_lista.Visible = False
   Me.frm_eliminar.Visible = False
End Sub

Private Sub lv_articulos_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      Me.txt_cantidad_eliminar = ""
      frm_eliminar.Visible = True
      txt_cantidad_eliminar.SetFocus
   End If
End Sub

Private Sub lv_articulos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_todos_lineas = 0
      i = lv_articulos.selectedItem.Index
      If lv_articulos.selectedItem.SubItems(6) = "*" Then
         lv_articulos.selectedItem.SubItems(6) = ""
         lv_articulos.ListItems.Item(i).Bold = False
         lv_articulos.ListItems.Item(i).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
         lv_articulos.Refresh
      Else
         lv_articulos.selectedItem.SubItems(6) = "*"
         lv_articulos.ListItems.Item(i).Bold = True
         lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
         lv_articulos.Refresh
      End If
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_lista.ListItems.Count > 0 Then
         Me.txt_planta = Me.lv_lista.selectedItem
         Me.txt_nombre_planta = Me.lv_lista.selectedItem.SubItems(1)
         Me.txt_planta.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.txt_planta.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad_eliminar) Then
         If CDbl(Me.lv_articulos.selectedItem.SubItems(2)) >= CDbl(Me.txt_cantidad_eliminar) Then
            Me.lv_articulos.selectedItem.SubItems(5) = Format(CDbl(Me.txt_cantidad_eliminar), "###,###,##0.00")
            Me.lv_articulos.SetFocus
         Else
            MsgBox "La cantidad debe ser menor o igual que a la cantidad del movimiento", vbOKOnly, "ATENCION"
            Me.lv_articulos.SetFocus
         End If
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_eliminar.Visible = False
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   Me.frm_eliminar.Visible = False
End Sub

Private Sub txt_planta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      rs.Open "SELECT * FROM tb_unidadesorganizacionales where vcha_emp_empresa_id = '" + var_empresa + "' ORDER BY VCHA_UOR_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
      lv_lista.ListItems.Clear
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_uor_unidad_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_UOR_NOMBRE), "", rs!VCHA_UOR_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PLANTAS"
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

Private Sub txt_planta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_planta <> "" Then
         rs.Open "select * from TB_UNIDADESORGANIZACIONALES where vcha_uor_unidad_id = '" + Me.txt_planta + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_nombre_planta = IIf(IsNull(rs!VCHA_UOR_NOMBRE), "", rs!VCHA_UOR_NOMBRE)
            If Me.lv_articulos.ListItems.Count > 0 Then
               Me.lv_articulos.SetFocus
            Else
               Me.txt_nombre_planta.SetFocus
            End If
         Else
            MsgBox "Clave de planta incorrecto", vbOKOnly, "ATENCION"
            Me.txt_planta = ""
            Me.txt_nombre_planta = ""
         End If
         rs.Close
      End If
   End If
End Sub
