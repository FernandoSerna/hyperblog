VERSION 5.00
Begin VB.Form frmcarga_pedido_coppel_excel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cargar COPPEL excel"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " Nombre del Archivo "
      Height          =   1140
      Left            =   120
      TabIndex        =   2
      Top             =   105
      Width           =   3600
      Begin VB.TextBox txt_archivo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   375
         TabIndex        =   0
         Top             =   435
         Width           =   2280
      End
      Begin VB.CommandButton cmd_generar_pedido 
         Height          =   450
         Left            =   2820
         Picture         =   "frmcarga_pedido_coppel_excel.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Subir pedido de COPPEL al S.I.D."
         Top             =   465
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmcarga_pedido_coppel_excel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tabla As ADODB.Connection
Dim txt_titular As String
Dim txt_establecimiento As String
Dim txt_clave_cliente As String
Dim txt_agente As String
Dim txt_numero As String
Dim var_primera_vez As Boolean
Dim var_cantidad_pedida As Variant
Dim var_precio_pedido As Variant
Dim var_nombre_articulo As String
Dim var_tipo_cliente As String
Dim var_suma_cantidad As Variant
Dim var_suma_importe As Variant
Dim var_descuento_1 As Variant
Dim var_descuento_2 As Variant
Dim var_descuento_3 As Variant
Dim var_dias_condiciones As Integer
Dim var_dias_caducidad As Integer
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_origen_codigo As Integer
Dim var_almacen As String
Dim var_lista_precios As String
Dim var_canal_venta As String
Dim var_clave_moneda As String
Dim var_resurtible As Integer
Dim var_tipo_lista As Integer
Dim var_renglon As Double
Dim var_estatus As String
Dim canal_venta As String


Private Sub cmd_generar_pedido_Click()
   Set TB_ENC_PEDIDOS_AUTOSERVICIOS_I = New TB_ENC_PEDIDOS_AUTOSERVICIOS_I
   Set TB_DETALLE_PEDIDOS_I = New TB_DETALLE_PEDIDOS_I
   Set TB_DETALLE_PEDIDOS_M = New TB_DETALLE_PEDIDOS_M
   Dim var_i As Integer
   Dim var_n As Integer
   Dim var_precio_anterior As Variant
   Dim list_item As ListItem
   Dim var_catalogo As String
   Dim var_numero_dias As Double
   Dim var_otorga_oferta As Boolean
   Dim var_posible As Boolean
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   Dim agrupador_catalogo As String
   Dim var_precio_externo As Double
   Dim var_catalogo_EFASA As String
   Dim var_posible_equivalencia As Boolean
   var_posible_equivalencia = True
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open "SELECT distinct NUMCODIGOCOPPEL FROM TB_ARCHIVO_PEDIDO_COPPEL_EXCEL WHERE VCHA_aRC_PEDIDO = '" + Me.txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_codigo = rs!NUMCODIGOCOPPEL
         If rsaux5.State = 1 Then
            rsaux5.Close
         End If
         rsaux5.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If rsaux5.EOF Then
            var_posible_equivalencia = False
         End If
         rsaux5.Close
         rs.MoveNext
   Wend
   rs.Close
   If var_posible_equivalencia = True Then
      rs.Open "select * from vw_clientes where vcha_cli_clave_id = 'C000002947'", cnn, adOpenDynamic, adLockOptimistic
      txt_agente = rs!vcha_age_agente_id
      txt_titular = rs!VCHA_TIT_TITULAR_ID
      txt_clave_cliente = rs!vcha_cli_clave_id
      var_descuento_1 = IIf(IsNull(rs!FLOA_GAC_DESCUENTO_1), 0, rs!FLOA_GAC_DESCUENTO_1)
      var_descuento_2 = IIf(IsNull(rs!FLOA_GAC_DESCUENTO_2), 0, rs!FLOA_GAC_DESCUENTO_2)
      var_descuento_3 = 0
      var_dias_condiciones = IIf(IsNull(rs!inte_pla_dias), 0, rs!inte_pla_dias)
      var_dias_caducidad = 6
      var_clave_moneda = rs!vcha_mon_moneda_id
      var_lista_precios = rs!vcha_lis_lista_id
      rs.Close
      If rsaux5.State = 1 Then
         rsaux5.Close
      End If
      rsaux5.Open "select distinct NUMBODEGADESTINO from TB_ARCHIVO_PEDIDO_COPPEL_EXCEL where  VCHA_ARC_PEDIDO = '" + Me.txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
      While Not rsaux5.EOF
            var_primera_vez = True
            If rsaux4.State = 1 Then
               rsaux4.Close
            End If
            
            rsaux4.Open "select * from TB_establecimientos where vcha_tit_titular_id = '" + txt_titular + "' and vcha_esb_establecimiento_anterior_id = '" + rsaux5!NUMBODEGADESTINO + "'", cnn, adOpenDynamic, adLockOptimistic
            txt_establecimiento = rsaux4!vcha_esb_establecimiento_id
            rsaux4.Close
            rs.Open "select * from tb_encabezado_pedidos where VCHA_PED_PEDIDO_EXTERNO = '" + Me.txt_archivo + "' and vcha_Esb_establecimiento_id = '" + txt_establecimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            If rs.EOF Then
               rs.Close
               rsaux4.Open "SELECT * FROM TB_ARCHIVO_PEDIDO_COPPEL_EXCEL WHERE VCHA_aRC_PEDIDO = '" + Me.txt_archivo + "' and NUMBODEGADESTINO = '" + rsaux5!NUMBODEGADESTINO + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                  var_origen_codigo = 0
                  If var_lista_precios <> "" Then
                     If Trim(var_clave_moneda) <> "" Then
                        While Not rsaux4.EOF
                              var_almacen = "8"
                              txt_codigo = rsaux4!NUMCODIGOCOPPEL
                              rsaux3.Open "select vcha_Art_articulo_id from tb_equivalencias where vcha_equ_codigo_equivalente = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                              txt_codigo = rsaux3!vcha_Art_articulo_id
                              rsaux3.Close
                              var_cantidad_pedida = rsaux4!CANTIDADPEDIDA
                              var_descuento_1 = 0
                              var_descuento_2 = 0
                              var_promocion_1 = 0
                              var_promocion_2 = 0
                              rsaux3.Open "SELECT * FROM TB_DETALLE_LISTA_PRECIOS WHERE VCHA_LIS_LISTA_PRECIOS_ID = '" + var_lista_precios + "' AND VCHA_ART_ARTICULO_ID = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 var_precio_pedido = Round(IIf(IsNull(rsaux3!floa_dli_precio), 0, rsaux3!floa_dli_precio), 2)
                              End If
                              rsaux3.Close
                              If Trim(txt_codigo) <> "" Then
                                 If var_primera_vez = True Then
                                    var_primera_vez = False
                                    ok = TB_ENC_PEDIDOS_AUTOSERVICIOS_I.Anadir(var_empresa, var_unidad_organizacional, "8", "M", maximo_pedido, 0, Date, Date, txt_agente, txt_titular, txt_clave_cliente, txt_establecimiento, 1, 0, "", var_descuento_1, var_descuento_2, var_descuento_3, var_dias_condiciones, var_dias_caducidad, var_clave_usuario_global, fun_NombrePc, Date, var_clave_moneda, 0, CStr(Me.txt_archivo))
                                    txt_numero = maximo_pedido
                                 End If
                                 rsaux.Open "select * from tb_detalle_pedidos where INTE_PED_NUMERO = " + txt_numero + " and VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND CHAR_PED_TIPO = 'P'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux.EOF Then
                                    rsaux.Close
                                    rs.Open "update tb_detalle_pedidos set floa_ped_cantidad = floa_ped_cantidad + " + CStr(var_cantidad_pedida) + " where inte_ped_numero = " + txt_numero + " and vcha_art_articulo_id = '" + txt_codigo + "' AND CHAR_PED_TIPO = 'P'", cnn, adOpenDynamic, adLockOptimistic
                                 Else
                                    rsaux.Close
                                    ok = TB_DETALLE_PEDIDOS_I.Anadir(CStr(var_empresa), CStr(var_unidad_organizacional), CStr(var_almacen), CVar(txt_numero), CVar(txt_codigo), CVar(var_precio_pedido), CVar(var_cantidad_pedida), 0, CDbl(var_promocion_1), CDbl(var_promocion_2), "P")
                                 End If
                              Else
                                 MsgBox "Código Incorrecto", vbOKOnly, "ATENCION"
                              End If
                              rsaux4.MoveNext
                         Wend
                     Else
                        MsgBox "El cliente no tiene una moneda asociada", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "El cliente no tiene una lista de precios asociada", vbOKOnly, "ATENCION"
                  End If
               Else
          
               End If
            Else
               rs.Close
               MsgBox "El pedido ya fue cargado con anterioridad", vbOKOnly, "ATENCION"
            End If
            rsaux5.MoveNext
      Wend
   Else
      MsgBox "Existen articulos sin equivalencias", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
   Set var_tabla = CreateObject("ADODB.connection")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_archivo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_generar_pedido.SetFocus
   End If
End Sub

Private Sub txt_archivo_LostFocus()
'On Error GoTo salir:
   If Trim(Me.txt_archivo) <> "" Then
      rs.Open "DELETE FROM TB_ARCHIVO_PEDIDO_COPPEL_EXCEL WHERE VCHA_ARC_PEDIDO = '" + Me.txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
      strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls); DBQ=c:\coppel\" + Me.txt_archivo + ".xls"
      If rsaux2.State = 1 Then
         rsaux2.Close
      End If
      rsaux2.Open "SELECT * FROM [" + Trim(Me.txt_archivo) + "$]", strConnectionString, adOpenDynamic, adLockOptimistic
      While Not rsaux2.EOF
            var_cadena = "INSERT INTO [TB_ARCHIVO_PEDIDO_COPPEL_EXCEL] ( [NUMBODEGARECIBE],[NUMBODEGADESTINO], [MODELOPROVEEDOR], [NUMCODIGOCOPPEL], [NUMTALLACOPPEL], [CANTIDADPEDIDA], [CANTIDADSURTIDA], [PRECIOCOSTO], [PRECIOVENTA], [NUMLOTE], [TOTALLOTES], [NUMFACTURA], [NUMPEDIDO], [TIPOPEDIDO], [IMPORTEFACTURA], [IVAFACTURA], [UNIDADESFACTURADAS], [NUMPROVEEDOR], [NETO], [FECHAFACTURA], VCHA_aRC_PEDIDO) "
            var_cadena = var_cadena + " Values ( '" + rsaux2!NUMBODEGARECIBE + "', '" + rsaux2!NUMBODEGADESTINO + "', '" + rsaux2!MODELOPROVEEDOR + "', '" + rsaux2!NUMCODIGOCOPPEL + "', '" + rsaux2!NUMTALLACOPPEL + "', '" + rsaux2!CANTIDADPEDIDA + "', '" + rsaux2!CANTIDADSURTIDA + "', '" + rsaux2!PRECIOCOSTO + "', '" + rsaux2!PRECIOVENTA + "', '" + rsaux2!NUMLOTE + "', '" + rsaux2!TOTALLOTES + "', '" + rsaux2!NUMFACTURA + "', '" + rsaux2!NUMPEDIDO + "', '" + rsaux2!TIPOPEDIDO + "', '" + rsaux2!IMPORTEFACTURA + "', '" + rsaux2!IVAFACTURA + "', '" + rsaux2!UNIDADESFACTURADAS + "', '" + rsaux2!NUMPROVEEDOR + "', '" + rsaux2!NETO + "', '" + rsaux2!FECHAFACTURA + "', '" + Me.txt_archivo + "')"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            rsaux2.MoveNext
      Wend
      rsaux2.Close
   End If
Exit Sub
salir:
  MsgBox "El archivo c:\coppel\" + Me.txt_archivo + " no existe o esta siendo usado por otro usuario", vbOKOnly, "ATENCION"
End Sub
