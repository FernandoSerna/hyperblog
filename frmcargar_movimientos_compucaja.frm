VERSION 5.00
Begin VB.Form frmcargar_movimientos_compucaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cargar movimientos del compucaja"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " Ultima carga "
      Height          =   750
      Left            =   210
      TabIndex        =   6
      Top             =   1800
      Width           =   4275
      Begin VB.TextBox txt_Fecha 
         Height          =   360
         Left            =   210
         TabIndex        =   7
         Top             =   255
         Width           =   3930
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50000
      Left            =   0
      Top             =   1170
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   225
      TabIndex        =   1
      Top             =   885
      Width           =   4245
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   315
         Width           =   1140
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   375
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2325
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CommandButton cmd_cargar 
      Caption         =   "Cargar movimientos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   180
      TabIndex        =   0
      Top             =   30
      Width           =   4335
   End
End
Attribute VB_Name = "frmcargar_movimientos_compucaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_i As Integer
Dim var_pedimento As String
Dim var_cantidad_multibondeados As Double
Dim var_kanban As String
Dim var_descripcion_etiqueta As String
Dim var_numero_serie As Integer
Dim var_txt_archivo As String
Dim var_clave_almacen_seleccionado As String
Dim var_peso_correcto As Boolean
Dim var_cajas As Boolean
Dim var_codigo_caja As String
Dim var_peso_caja As Double
Dim var_cantidad_caja_peso As Double
Dim var_tolerancia_peso_caja As Double
Dim var_año As Integer
Dim var_origen As String
Dim var_lote As Double
Dim var_consecutivo As Integer
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
Dim var_renglon As Double

Private Sub cmd_cargar_Click()
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
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         rs.Open "delete from TB_CARGAR_MOVIMIENTOS_COMPUCAJA", cnn, adOpenDynamic, adLockOptimistic
         var_dia = CStr(Day(Me.txt_inicio))
         var_mes = CStr(Month(Me.txt_inicio))
         var_año = CStr(Year(Me.txt_inicio))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         var_fecha_inicio = "{d '" + CStr(var_año) + "-" + CStr(var_mes) + "-" + CStr(var_dia) + "'}"
         
         var_dia = CStr(Day(Me.txt_fin))
         var_mes = CStr(Month(Me.txt_fin))
         var_año = CStr(Year(Me.txt_fin))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         var_fecha_fin = "{d '" + CStr(var_año) + "-" + CStr(var_mes) + "-" + CStr(var_dia) + "'}"
    
    
         'rs.Open "select DISTINCT 'CC_'+CAST(ALM_CODIGO AS VARCHAR(50)) AS ALMACEN, ALM_DESCRIPCION  from kardex_para_sid where alm_descripcion <> 'A RESTAURANTE' and (tipo_documento = 'SALIDA' or tma_descripcion = 'NOTA DE CRÉDITO') and fecha >= " + var_fecha_inicio + " and fecha < " + var_fecha_fin + "+1 and tma_codigo not in (104, 13, 7, 8, 105, 14, 17, 102, 101, 6, 1, 103) and folio not in (select vcha_fol_folio_id from TB_FOLIOS_MOVIMIENTOS_SID)", cnn_compucaja, adOpenDynamic, adLockOptimistic
         rs.Open "select DISTINCT 'CC_'+CAST(ALM_CODIGO AS VARCHAR(50)) AS ALMACEN, ALM_DESCRIPCION  from kardex_para_sid where (tipo_documento = 'SALIDA' or tma_descripcion = 'NOTA DE CRÉDITO') and fecha >= " + var_fecha_inicio + " and fecha < " + var_fecha_fin + "+1 and tma_codigo not in (104, 13, 7, 8, 105, 14, 17, 102, 101, 6, 1, 103) and folio not in (select vcha_fol_folio_id from TB_FOLIOS_MOVIMIENTOS_SID)", cnn_compucaja, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux.Open "SELECT * FROM TB_ALMACENES WHERE VCHA_ALM_ALMACEN_ID = '" + rs!ALMACEN + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux.EOF Then
                  rsaux1.Open "INSERT INTO TB_ALMACENES (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_ALM_NOMBRE, VCHA_PAI_PAIS_ID, VCHA_EST_eSTADO_ID, VCHA_CIU_CIUDAD_ID, VCHA_ALM_DIRECCION, VCHA_ALM_CP, CHAR_ALM_TIPO, VCHA_COL_COLONIA_ID, VCHA_MUN_MUNICIPIO_ID) VALUES  ('31','26','" + rs!ALMACEN + "', '" + rs!ALM_DESCRIPCION + " ', '001', '00001', '000002', 'CARRETERA CALVILLO KM 1.5 PARQUE IND EL VERGEL','20219','A', '0000000192', '000001')", cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux.Close
               rs.MoveNext
         Wend
         rs.Close
    
         'rs.Open "select distinct 'CC_'+cast(tma_codigo as varchar(50)) AS VCHA_MOV_MOVIMIENTO_ID, TMA_DESCRIPCION AS VCHA_MOV_NOMBRE, TIPO_DOCUMENTO from kardex_para_sid where alm_descripcion <> 'A RESTAURANTE' AND FECHA >= " + var_fecha_inicio + " AND FECHA < " + var_fecha_fin + "+1  and tma_codigo not in (104, 13, 7, 8, 105, 14, 17, 102, 101, 6, 1, 103) and folio not in (select vcha_fol_folio_id from TB_FOLIOS_MOVIMIENTOS_SID)", cnn_compucaja, adOpenDynamic, adLockOptimistic
         rs.Open "select distinct 'CC_'+cast(tma_codigo as varchar(50)) AS VCHA_MOV_MOVIMIENTO_ID, TMA_DESCRIPCION AS VCHA_MOV_NOMBRE, TIPO_DOCUMENTO from kardex_para_sid where FECHA >= " + var_fecha_inicio + " AND FECHA < " + var_fecha_fin + "+1  and tma_codigo not in (104, 13, 7, 8, 105, 14, 17, 102, 101, 6, 1, 103) and folio not in (select vcha_fol_folio_id from TB_FOLIOS_MOVIMIENTOS_SID)", cnn_compucaja, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux.Open "select * from tb_movimientos where vcha_mov_movimiento_id = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux.EOF Then
                  If rs!TIPO_DOCUMENTO = "ENTRADA" Then
                     afectacion = "+"
                  End If
                  If rs!TIPO_DOCUMENTO = "SALIDA" Then
                     afectacion = "-"
                  End If
                  var_cadena = "insert into tb_movimientos (vcha_mov_movimiento_id, vcha_mov_nombre, char_mov_afectacion, inte_mov_refereancia, char_mov_tipo_proveedor, char_mov_tipo_cliente, inte_mov_ajuste) values "
                  var_cadena = var_cadena + "('" + rs!VCHA_MOV_MOVIMIENTO_ID + "', '" + rs!vcha_mov_nombre + "', '" + afectacion + "',0,'A','A',1)"
                  rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux.Close
               rs.MoveNext
         Wend
         rs.Close
    
         'rs.Open "select 'CC_'+CAST(ALM_CODIGO AS VARCHAR(50)) AS VCHA_ALM_ALMACEN_ID, ALM_DESCRIPCION AS VCHA_ALM_NOMBRE, 'CC_'+cast(tma_codigo as varchar(50)) AS VCHA_MOV_MOVIMIENTO_ID, TMA_DESCRIPCION AS VCHA_MOV_NOMBRE, FOLIO, FECHA, ART_CODIGO AS VCHA_ART_ARTICULO_ID, ART_DESCRIPCION AS VCHA_ART_NOMBRE_ESPAÑOL, TIPO_DOCUMENTO, CANTIDAD, isnull(REFERENCIA1,'')+' '+isnull(REFERENCIA2,'') AS REFERENCIA, PRECIOUNITARIO, COSTO, FP_CODIGO  from kardex_para_sid where alm_descripcion <> 'A RESTAURANTE' AND FECHA >= " + var_fecha_inicio + " AND FECHA < " + var_fecha_fin + "+1  and tma_codigo not in (104, 13, 7, 8, 105, 14, 17, 102, 101, 6, 1, 103) and folio not in (select vcha_fol_folio_id from TB_FOLIOS_MOVIMIENTOS_SID) order by folio", cnn_compucaja, adOpenDynamic, adLockOptimistic
         rs.Open "select 'CC_'+CAST(ALM_CODIGO AS VARCHAR(50)) AS VCHA_ALM_ALMACEN_ID, ALM_DESCRIPCION AS VCHA_ALM_NOMBRE, 'CC_'+cast(tma_codigo as varchar(50)) AS VCHA_MOV_MOVIMIENTO_ID, TMA_DESCRIPCION AS VCHA_MOV_NOMBRE, FOLIO, FECHA, ART_CODIGO AS VCHA_ART_ARTICULO_ID, ART_DESCRIPCION AS VCHA_ART_NOMBRE_ESPAÑOL, TIPO_DOCUMENTO, CANTIDAD, isnull(REFERENCIA1,'')+' '+isnull(REFERENCIA2,'') AS REFERENCIA, PRECIOUNITARIO, COSTO, FP_CODIGO  from kardex_para_sid where FECHA >= " + var_fecha_inicio + " AND FECHA < " + var_fecha_fin + "+1  and tma_codigo not in (104, 13, 7, 8, 105, 14, 17, 102, 101, 6, 1, 103) and folio not in (select vcha_fol_folio_id from TB_FOLIOS_MOVIMIENTOS_SID) order by folio", cnn_compucaja, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               'var_fecha = CStr(rs!fecha)
               'MsgBox var_fecha
               var_dia = CStr(Day(rs!fecha))
               var_mes = CStr(Month(rs!fecha))
               var_año = CStr(Year(rs!fecha))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha = "{d '" + CStr(var_año) + "-" + CStr(var_mes) + "-" + CStr(var_dia) + "'}"
               var_cadena = "INSERT INTO TB_CARGAR_MOVIMIENTOS_COMPUCAJA (VCHA_ALM_ALMACEN_ID, VCHA_ALM_NOMBRE, VCHA_MOV_MOVIMIENTO_ID, VCHA_MOV_NOMBRE, VCHA_EMO_REFERENCIA_FOLIO, DTIM_EMO_FECHA, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_ESPAÑOL, VCHA_EMO_AFECTACION, FLOA_EMO_CANTIDAD, VCHA_EMO_REFERENCIA, VCHA_EMO_FOLIO_CODIGO, FLOA_EMO_COSTO, FLOA_EMO_PRECIO, VCHA_EMO_FP_CODIGO ) "
               var_cadena = var_cadena + " VALUES ('" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_ALM_NOMBRE + "', '" + CStr(rs!VCHA_MOV_MOVIMIENTO_ID) + "', '" + rs!vcha_mov_nombre + "', '" + rs!FOLIO + "'," + var_fecha + ",'" + rs!vcha_Art_articulo_id + "', '" + Mid(rs!vcha_art_nombre_español, 1, 50) + "', '" + rs!TIPO_DOCUMENTO + "', " + CStr(rs!Cantidad) + ", '" + Mid(rs!Referencia, 1, 50) + "','" + rs!FOLIO + rs!vcha_Art_articulo_id + "'," + CStr(rs!Costo) + "," + CStr(rs!PRECIOUNITARIO) + ", '" + CStr(IIf(IsNull(rs!FP_CODIGO), "", rs!FP_CODIGO)) + "')"
               rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               rs.MoveNext
         Wend
         rs.Close
         
         rs.Open "select * from TB_CARGAR_MOVIMIENTOS_COMPUCAJA where vcha_Art_articulo_id not in (select vcha_Art_articulo_id from tb_Articulos)", cnn, adOpenDynamic, adLockOptimistic
         var_codigos_faltantes = ""
         While Not rs.EOF
               rsaux1.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux1.EOF Then
                  var_costo = 0
                  rsaux9.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + rsaux1!VCHA_ALM_ALMACEN_ID + "' and vcha_Art_Articulo_id = '" + rs!vcha_Art_articulo_id + "' ", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux9.EOF Then
                     var_costo = IIf(IsNull(rsaux9!FLOA_eXI_COSTO), 0, rsaux9!FLOA_eXI_COSTO)
                  Else
                     rsaux8.Open "select * from avl_precios where art_codigo = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux8.EOF Then
                        var_costo = rsaux!art_ultimocosto
                     End If
                     rsaux8.Close
                  End If
                  rsaux9.Close
                  var_cadena = "insert into tb_articulos (vcha_Art_articulo_id, vcha_Art_nombre_español, mone_Art_costo_Estandar, mone_Art_precio_base, dtim_Art_fecha_alta, vcha_lic_licencia_id, vcha_art_numero_lic, inte_art_detenido, vcha_equ_equivalencia_id, vcha_art_codigo_externo)"
                  var_cadena = var_cadena + "        values ('" + rs!vcha_Art_articulo_id + "','" + rs!vcha_art_nombre_español + "'," + CStr(var_costo) + ", " + CStr(rs!floa_Emo_precio / 1.16) + ",getdate(),'SIN LICENCIA','SIN LICENCIA',0,'" + rs!vcha_Art_articulo_id + "','" + rs!vcha_Art_articulo_id + "')"
                  rsaux9.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux1.Close
               rs.MoveNext
         Wend
         rs.Close
         rs.Open "select distinct vcha_mov_movimiento_id, vcha_alm_almacen_id from TB_CARGAR_MOVIMIENTOS_COMPUCAJA", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux.Open "select * from tb_movimientos_almacenes where vcha_mov_movimiento_id = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' and vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux.EOF Then
                  rsaux1.Open "insert into tb_movimientos_almacenes (vcha_alm_almacen_id, vcha_mov_movimiento_id) values ('" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "')", cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux.Close
               rs.MoveNext
         Wend
         rs.Close
         rs.Open "select distinct vcha_Emo_referencia_folio, vcha_alm_almacen_id, vcha_mov_movimiento_id, vcha_emo_referencia_folio, dtim_Emo_fecha, vcha_emo_afectacion from tb_Cargar_movimientos_compucaja", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' and vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "' and vcha_emo_referencia = '" + rs!vcha_emo_referencia_folio + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux.EOF Then
                  var_almacen_Destino = rs!VCHA_ALM_ALMACEN_ID
                  var_clave_movimiento = rs!VCHA_MOV_MOVIMIENTO_ID
                  var_almacen_origen = rs!VCHA_ALM_ALMACEN_ID
                  'MsgBox rs!vcha_Emo_referencia
                  var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, "", "", rs!vcha_emo_referencia_folio, "", "", "", "", 0, 0, 0, "1", 1)
                  var_numero_folio = var_numero_folio_regreso
                  rsaux1.Open "update tb_Cargar_movimientos_compucaja set inte_sid_folio = " + CStr(var_numero_folio) + " where vcha_mov_movimiento_id = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' and vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "' and vcha_emo_referencia_folio = '" + rs!vcha_emo_referencia_folio + "'", cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux.Close
               rs.MoveNext
         Wend
         rs.Close
         
         rs.Open "select * from tb_Cargar_movimientos_compucaja where vcha_emo_afectacion = 'SALIDA' and inte_sid_folio is not null", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux1.Open "select * from tb_temporal_salidas with (nolock) where vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' and inte_Sal_numero = " + CStr(rs!inte_Sid_folio) + " and vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux1.EOF Then
                  var_año_iva = Year(rs!dtim_emo_fecha)
                  If var_año_iva < 2010 Then
                     var_precio = rs!floa_Emo_precio / 1.15
                  Else
                     var_precio = rs!floa_Emo_precio / 1.16
                  End If
                  
                  
                  rsaux9.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "' and vcha_Art_Articulo_id = '" + rs!vcha_Art_articulo_id + "' ", cnn, adOpenDynamic, adLockOptimistic
                  var_costo = 0
                  If Not rsaux9.EOF Then
                     var_costo = IIf(IsNull(rsaux9!FLOA_eXI_COSTO), 0, rsaux9!FLOA_eXI_COSTO)
                  Else
                     rsaux8.Open "select * from avl_precios where art_codigo = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux8.EOF Then
                        var_costo = rsaux!art_ultimocosto
                     End If
                     rsaux8.Close
                  End If
                  rsaux9.Close
                  
                  
                  
                  var_cadena = "insert into tb_temporal_salidas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_Sal_numero, vcha_Art_Articulo_id, floa_sal_Cantidad, floa_sal_costo, floa_sal_precio, inte_Sal_año)"
                  var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + rs!VCHA_ALM_ALMACEN_ID + "','" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_Sid_folio) + ",'" + rs!vcha_Art_articulo_id + "', " + CStr(rs!floa_emo_Cantidad) + "," + CStr(var_costo) + "," + CStr(var_precio) + ",2005)"
                  rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  var_cadena = "insert into tb_salidas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_Sal_numero, vcha_Art_Articulo_id, floa_sal_Cantidad, floa_sal_costo, floa_sal_precio, inte_Sal_año)"
                  var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + rs!VCHA_ALM_ALMACEN_ID + "','" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_Sid_folio) + ",'" + rs!vcha_Art_articulo_id + "', " + CStr(rs!floa_emo_Cantidad) + "," + CStr(var_costo) + "," + CStr(var_precio) + ",2005)"
                  rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               Else
                  rsaux.Open "update tb_Temporal_salidas set floa_Sal_Cantidad =  floa_sal_cantidad + " + CStr(rs!floa_emo_Cantidad) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' and inte_Sal_numero = " + CStr(rs!inte_Sid_folio) + " and vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux.Open "update tb_salidas set floa_Sal_Cantidad =  floa_sal_cantidad + " + CStr(rs!floa_emo_Cantidad) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' and inte_Sal_numero = " + CStr(rs!inte_Sid_folio) + " and vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux1.Close
               rs.MoveNext
         Wend
         rs.Close
         
         rs.Open "select * from tb_Cargar_movimientos_compucaja where vcha_emo_afectacion = 'ENTRADA' and inte_sid_folio is not null", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux1.Open "select * from tb_temporal_entradas where vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' and inte_ent_numero = " + CStr(rs!inte_Sid_folio) + " and vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux1.EOF Then
                  var_año_iva = Year(rs!dtim_emo_fecha)
                  If var_año_iva < 2010 Then
                     var_precio = rs!floa_Emo_precio / 1.15
                  Else
                     var_precio = rs!floa_Emo_precio / 1.16
                  End If
                  
                  var_costo = 0
                  rsaux9.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "' and vcha_Art_Articulo_id = '" + rs!vcha_Art_articulo_id + "' ", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux9.EOF Then
                     var_costo = IIf(IsNull(rsaux9!FLOA_eXI_COSTO), 0, rsaux9!FLOA_eXI_COSTO)
                  Else
                     rsaux8.Open "select * from avl_precios where art_codigo = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux8.EOF Then
                        var_costo = rsaux!art_ultimocosto
                     End If
                     rsaux8.Close
                  End If
                  rsaux9.Close
                  
                  
                  var_cadena = "insert into tb_temporal_entradas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_Art_Articulo_id, floa_ent_Cantidad, floa_ent_costo, floa_ent_precio, inte_ent_año)"
                  var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + rs!VCHA_ALM_ALMACEN_ID + "','" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_Sid_folio) + ",'" + rs!vcha_Art_articulo_id + "', " + CStr(rs!floa_emo_Cantidad) + "," + CStr(rs!floa_emo_costo) + "," + CStr(var_precio) + ",2005)"
                  rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  var_cadena = "insert into tb_entradas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_Art_Articulo_id, floa_ent_Cantidad, floa_ent_costo, floa_ent_precio, inte_ent_año)"
                  var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + rs!VCHA_ALM_ALMACEN_ID + "','" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_Sid_folio) + ",'" + rs!vcha_Art_articulo_id + "', " + CStr(rs!floa_emo_Cantidad) + "," + CStr(rs!floa_emo_costo) + "," + CStr(Round(var_precio, 2)) + ",2005)"
                  'MsgBox var_cadena
                  rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               Else
                  rsaux.Open "update tb_Temporal_entradas set floa_ent_Cantidad =  floa_ent_cantidad + " + CStr(rs!floa_emo_Cantidad) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' and inte_ent_numero = " + CStr(rs!inte_Sid_folio) + " and vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux.Open "update tb_entradas set floa_ent_Cantidad =  floa_ent_cantidad + " + CStr(rs!floa_emo_Cantidad) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' and inte_ent_numero = " + CStr(rs!inte_Sid_folio) + " and vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux1.Close
               rs.MoveNext
         Wend
         rs.Close
         
         rs.Open "select distinct vcha_Emo_referencia_folio, vcha_alm_almacen_id, vcha_mov_movimiento_id, vcha_emo_referencia_folio, dtim_Emo_fecha, vcha_emo_afectacion, inte_sid_folio, dtim_Emo_fecha from tb_Cargar_movimientos_compucaja where inte_sid_folio is not null", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               var_dia = CStr(Day(rs!dtim_emo_fecha))
               var_mes = CStr(Month(rs!dtim_emo_fecha))
               var_año = CStr(Year(rs!dtim_emo_fecha))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_fin = "{d '" + CStr(var_año) + "-" + CStr(var_mes) + "-" + CStr(var_dia) + "'}"
               
               rsaux.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' and vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "' and vcha_emo_referencia = '" + rs!vcha_emo_referencia_folio + "' and inte_emo_numero = " + CStr(rs!inte_Sid_folio), cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  rsaux1.Open "update tb_Encabezado_movimientos set char_emo_Estatus = 'I' where  vcha_mov_movimiento_id = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' and vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "' and vcha_emo_referencia = '" + rs!vcha_emo_referencia_folio + "' and inte_emo_numero = " + CStr(rs!inte_Sid_folio), cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "update tb_Encabezado_movimientos set dtim_emo_fecha = " + var_fecha_fin + ", inte_emo_bloqueado = 0 where  vcha_mov_movimiento_id = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' and vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "' and vcha_emo_referencia = '" + rs!vcha_emo_referencia_folio + "' and inte_emo_numero = " + CStr(rs!inte_Sid_folio), cnn, adOpenDynamic, adLockOptimistic
                  If rs!vcha_Emo_afectacion = "ENTRADA" Then
                     rsaux1.Open "UPDATE TB_eNTRADAS SET DTIM_ENT_FECHA = " + var_fecha_fin + " where  vcha_mov_movimiento_id = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' and vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "' and inte_ENT_numero = " + CStr(rs!inte_Sid_folio), cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux1.Open "UPDATE  TB_SALIDAS SET DTIM_SAL_FECHA = " + var_fecha_fin + " where  vcha_mov_movimiento_id = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' and vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "'  and inte_SAL_numero = " + CStr(rs!inte_Sid_folio), cnn, adOpenDynamic, adLockOptimistic
                  End If
               End If
               rsaux.Close
               rs.MoveNext
         Wend
         rs.Close
         'traspasos apartado
         
         
         rs.Open "select distinct vcha_Emo_referencia_folio, VCHA_ALM_ALMACEN_ID, dtim_Emo_fecha from tb_Cargar_movimientos_compucaja WHERE VCHA_EMO_FP_CODIGO = '9004'", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = 'T' and VCHA_ALM_ALMACEN_ID = 'ALAP' AND  VCHA_EMO_ALMACEN_DESTINO = '" + rs!VCHA_ALM_ALMACEN_ID + "' and vcha_emo_referencia = '" + rs!vcha_emo_referencia_folio + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux.EOF Then
                  var_almacen_Destino = rs!VCHA_ALM_ALMACEN_ID
                  var_clave_movimiento = "T"
                  var_almacen_origen = "ALAP"
                  var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, 0, "", "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, "", "", rs!vcha_emo_referencia_folio, "", "", "", "", 0, 0, 0, "1", 1)
                  var_numero_folio = var_numero_folio_regreso
                  rsaux1.Open "update tb_Cargar_movimientos_compucaja set inte_sid_folio_TRASPASO = " + CStr(var_numero_folio) + " where  vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "' and vcha_emo_referencia_folio = '" + rs!vcha_emo_referencia_folio + "'", cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux.Close
               rs.MoveNext
         Wend
         rs.Close
         
         
         
         
         rs.Open "select * from tb_Cargar_movimientos_compucaja where VCHA_EMO_FP_CODIGO = '9004' and inte_sid_folio_TRASPASO is not null", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux1.Open "select * from tb_temporal_entradas where vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_ent_numero = " + CStr(rs!inte_Sid_folio_tRASPASO) + " and vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux1.EOF Then
                  var_año_iva = Year(rs!dtim_emo_fecha)
                  If var_año_iva < 2010 Then
                     var_precio = rs!floa_Emo_precio / 1.15
                  Else
                     var_precio = rs!floa_Emo_precio / 1.16
                  End If
                  
                  var_costo = 0
                  rsaux9.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + rsaux1!VCHA_ALM_ALMACEN_ID + "' and vcha_Art_Articulo_id = '" + rs!vcha_Art_articulo_id + "' ", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux9.EOF Then
                     var_costo = IIf(IsNull(rsaux9!FLOA_eXI_COSTO), 0, rsaux9!FLOA_eXI_COSTO)
                  Else
                     rsaux8.Open "select * from avl_precios where art_codigo = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux8.EOF Then
                        var_costo = rsaux!art_ultimocosto
                     End If
                     rsaux8.Close
                  End If
                  rsaux9.Close
                  
                  
                  var_cadena = "insert into tb_temporal_entradas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_Art_Articulo_id, floa_ent_Cantidad, floa_ent_costo, floa_ent_precio, inte_ent_año, vcha_ent_almacen_origen)"
                  var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + rs!VCHA_ALM_ALMACEN_ID + "','T', " + CStr(rs!inte_Sid_folio_tRASPASO) + ",'" + rs!vcha_Art_articulo_id + "', " + CStr(rs!floa_emo_Cantidad) + "," + CStr(var_costo) + "," + CStr(var_precio) + ",2005,'ALAP')"
                  rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  var_cadena = "insert into tb_entradas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_Art_Articulo_id, floa_ent_Cantidad, floa_ent_costo, floa_ent_precio, inte_ent_año, vcha_ent_almacen_origen)"
                  var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + rs!VCHA_ALM_ALMACEN_ID + "','T', " + CStr(rs!inte_Sid_folio_tRASPASO) + ",'" + rs!vcha_Art_articulo_id + "', " + CStr(rs!floa_emo_Cantidad) + "," + CStr(var_costo) + "," + CStr(Round(var_precio, 2)) + ",2005,'ALAP')"
                  rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  
                  var_cadena = "insert into tb_temporal_salidas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_Sal_numero, vcha_Art_Articulo_id, floa_sal_Cantidad, floa_sal_costo, floa_sal_precio, inte_Sal_año)"
                  var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','ALAP','T', " + CStr(rs!inte_Sid_folio_tRASPASO) + ",'" + rs!vcha_Art_articulo_id + "', " + CStr(rs!floa_emo_Cantidad) + "," + CStr(var_costo) + "," + CStr(var_precio) + ",2005)"
                  rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  var_cadena = "insert into tb_salidas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_Sal_numero, vcha_Art_Articulo_id, floa_sal_Cantidad, floa_sal_costo, floa_sal_precio, inte_Sal_año)"
                  var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','ALAP','T', " + CStr(rs!inte_Sid_folio_tRASPASO) + ",'" + rs!vcha_Art_articulo_id + "', " + CStr(rs!floa_emo_Cantidad) + "," + CStr(var_costo) + "," + CStr(var_precio) + ",2005)"
                  rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               
               
               Else
                  rsaux.Open "update tb_Temporal_entradas set floa_ent_Cantidad =  floa_ent_cantidad + " + CStr(rs!floa_emo_Cantidad) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_ent_numero = " + CStr(rs!inte_Sid_folio_tRASPASO) + " and vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux.Open "update tb_entradas set floa_ent_Cantidad =  floa_ent_cantidad + " + CStr(rs!floa_emo_Cantidad) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_ent_numero = " + CStr(rs!inte_Sid_folio_tRASPASO) + " and vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  
                  rsaux.Open "update tb_Temporal_salidas set floa_Sal_Cantidad =  floa_sal_cantidad + " + CStr(rs!floa_emo_Cantidad) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_Sal_numero = " + CStr(rs!inte_Sid_folio_tRASPASO) + " and vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux.Open "update tb_salidas set floa_Sal_Cantidad =  floa_sal_cantidad + " + CStr(rs!floa_emo_Cantidad) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_Sal_numero = " + CStr(rs!inte_Sid_folio_tRASPASO) + " and vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux1.Close
               rs.MoveNext
         Wend
         rs.Close
         
         
         rs.Open "select distinct vcha_Emo_referencia_folio, vcha_alm_almacen_id, vcha_mov_movimiento_id, vcha_emo_referencia_folio, dtim_Emo_fecha, vcha_emo_afectacion, inte_sid_folio_TRASPASO, dtim_Emo_fecha from tb_Cargar_movimientos_compucaja where inte_Sid_folio_tRASPASO is not null", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               var_dia = CStr(Day(rs!dtim_emo_fecha))
               var_mes = CStr(Month(rs!dtim_emo_fecha))
               var_año = CStr(Year(rs!dtim_emo_fecha))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_fin = "{d '" + CStr(var_año) + "-" + CStr(var_mes) + "-" + CStr(var_dia) + "'}"
               
               rsaux.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = 'T' and vcha_alm_almacen_id = 'ALAP' and vcha_emo_referencia = '" + rs!vcha_emo_referencia_folio + "' and inte_emo_numero = " + CStr(rs!inte_Sid_folio_tRASPASO), cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  rsaux1.Open "update tb_Encabezado_movimientos set char_emo_Estatus = 'I' where  vcha_mov_movimiento_id = 'T' and vcha_alm_almacen_id = 'ALAP' and vcha_emo_referencia = '" + rs!vcha_emo_referencia_folio + "' and inte_emo_numero = " + CStr(rs!inte_Sid_folio_tRASPASO), cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "update tb_Encabezado_movimientos set dtim_emo_fecha = " + var_fecha_fin + ", inte_emo_bloqueado = 0  where  vcha_mov_movimiento_id = 'T' and vcha_alm_almacen_id = 'ALAP' and vcha_emo_referencia = '" + rs!vcha_emo_referencia_folio + "' and inte_emo_numero = " + CStr(rs!inte_Sid_folio_tRASPASO), cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "UPDATE TB_eNTRADAS SET DTIM_ENT_FECHA = " + var_fecha_fin + " where  vcha_mov_movimiento_id = 'T' and vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "' and inte_ENT_numero = " + CStr(rs!inte_Sid_folio_tRASPASO), cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "UPDATE  TB_SALIDAS SET DTIM_SAL_FECHA = " + var_fecha_fin + " where  vcha_mov_movimiento_id = 'T' and vcha_alm_almacen_id = 'ALAP'  and inte_SAL_numero = " + CStr(rs!inte_Sid_folio_tRASPASO), cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux.Close
               rs.MoveNext
         Wend
         rs.Close
         
         
         
         
         rs.Open "select distinct vcha_Emo_referencia_folio, vcha_alm_almacen_id, vcha_mov_movimiento_id, vcha_emo_referencia_folio, dtim_Emo_fecha, vcha_emo_afectacion from tb_Cargar_movimientos_compucaja  WHERE inte_sid_folio is not null", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux1.Open "insert into tb_folios_movimientos_sid (vcha_fol_folio_id) values ('" + rs!vcha_emo_referencia_folio + "')", cnn_compucaja, adOpenDynamic, adLockOptimistic
               rs.MoveNext
         Wend
         rs.Close
         
         
         
         
         
         
         rs.Open "DELETE FROM TB_CARGAR_APARTADOS_COMPUCAJA", cnn, adOpenDynamic, adLockOptimistic
         rs.Open "select * from via_detalle_apartado WHERE ap_estado IN (0,2) and folio_estatus not in (select vcha_fol_folio_id from tb_folios_apartados) and ap_fecha >= " + var_fecha_inicio + " order by folio", cnn_compucaja
         While Not rs.EOF
               var_dia = CStr(Day(rs!AP_FECHA))
               var_mes = CStr(Month(rs!AP_FECHA))
               var_año = CStr(Year(rs!AP_FECHA))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_fin = "{d '" + CStr(var_año) + "-" + CStr(var_mes) + "-" + CStr(var_dia) + "'}"
               var_cadena = "INSERT INTO TB_CARGAR_APARTADOS_COMPUCAJA (VCHA_EMO_REFERENCIA_FOLIO, DTIM_EMO_FECHA, VCHA_ART_ARTICULO_ID, FLOA_EMO_CANTIDAD, FLOA_EMO_PRECIO, INTE_EMO_AP_ESTADO, VCHA_FOL_FOLIO_ID) "
               var_cadena = var_cadena + " VALUES ('" + rs!FOLIO + "'," + var_fecha_fin + ",'" + rs!art_codigo + "'," + CStr(rs!DA_CANTIDAD) + "," + CStr(rs!DA_PRECIOVENTA / 1.16) + "," + CStr(rs!AP_ESTado) + ",'" + rs!FOLIO_ESTATUS + "')"
               rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               rs.MoveNext
         Wend
         rs.Close
         rs.Open "SELECT DISTINCT vcha_fol_folio_id, inte_emo_ap_Estado, dtim_Emo_fecha FROM TB_CARGAR_APARTADOS_COMPUCAJA", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux.Open "select * from tb_Encabezado_movimientos where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND vcha_mov_movimiento_id = 'T' and vcha_emo_referencia = '" + rs!vcha_fol_folio_id + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux.EOF Then
                  var_dia = CStr(Day(rs!dtim_emo_fecha))
                  var_mes = CStr(Month(rs!dtim_emo_fecha))
                  var_año = CStr(Year(rs!dtim_emo_fecha))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_fin = "{d '" + CStr(var_año) + "-" + CStr(var_mes) + "-" + CStr(var_dia) + "'}"
                  If rs!inte_emo_ap_Estado = 0 Then
                     var_almacen_Destino = "ALAP"
                     var_clave_movimiento = "T"
                     var_almacen_origen = "CC_1"
                  End If
                  If rs!inte_emo_ap_Estado = 2 Then
                     var_almacen_Destino = "CC_1"
                     var_clave_movimiento = "T"
                     var_almacen_origen = "ALAP"
                  End If
                  var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, 0, "", "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, "", "", rs!vcha_fol_folio_id, "", "", "", "", 0, 0, 0, "1", 1)
                  var_numero_folio = var_numero_folio_regreso
                  rsaux2.Open "SELECT * FROM TB_CARGAR_APARTADOS_COMPUCAJA WHERE VCHA_FOL_FOLIO_ID = '" + rs!vcha_fol_folio_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux2.EOF
                        rsaux1.Open "select * from tb_temporal_entradas where vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_ent_numero = " + CStr(var_numero_folio) + " and vcha_Art_articulo_id = '" + rsaux2!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If rsaux1.EOF Then
                           var_año_iva = Year(rs!dtim_emo_fecha)
                           var_precio = rsaux2!floa_Emo_precio
                           
                           var_costo = 0
                           rsaux9.Open "select * from tb_existencias where vcha_alm_almacen_id = 'ALAP' and vcha_Art_Articulo_id = '" + rsaux2!vcha_Art_articulo_id + "' ", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux9.EOF Then
                              var_costo = IIf(IsNull(rsaux9!FLOA_eXI_COSTO), 0, rsaux9!FLOA_eXI_COSTO)
                           Else
                              rsaux8.Open "select * from avl_precios where art_codigo = '" + rsaux2!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux8.EOF Then
                                 var_costo = rsaux!art_ultimocosto
                              End If
                              rsaux8.Close
                           End If
                           rsaux9.Close
                           
                           rsaux3.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux2!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux3.EOF Then
                              'var_costo = IIf(IsNull(rsaux3!mone_art_costo_Estandar), 0, rsaux3!mone_art_costo_Estandar)
                           Else
                              rsaux10.Open "select * from Articulos where art_codigo = '" + rsaux2!vcha_Art_articulo_id + "'", cnn_compucaja, adOpenDynamic, adLockOptimistic
                              var_año_iva = Year(rs!dtim_emo_fecha)
                              'If var_año_iva < 2010 Then
                              '   var_costo = rsaux10!art_ultimocosto / 1.15
                              'Else
                              '   var_costo = rsaux10!art_ultimocosto / 1.16
                              'End If
                              var_cadena = "insert into tb_articulos (vcha_Art_articulo_id, vcha_Art_nombre_español, mone_Art_costo_Estandar, mone_Art_precio_base, dtim_Art_fecha_alta, vcha_lic_licencia_id, vcha_art_numero_lic, inte_art_detenido, vcha_equ_equivalencia_id, vcha_art_codigo_externo)"
                              var_cadena = var_cadena + "        values ('" + rsaux2!vcha_Art_articulo_id + "','" + rsaux10!art_descripcion + "'," + CStr(var_costo) + ", " + CStr(rsaux2!floa_Emo_precio) + ",getdate(),'SIN LICENCIA','SIN LICENCIA',0,'" + rsaux10!art_codigo + "','" + rsaux2!vcha_Art_articulo_id + "')"
                              rsaux9.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              rsaux10.Close
                           End If
                           rsaux3.Close
                           var_cadena = "insert into tb_temporal_entradas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_Art_Articulo_id, floa_ent_Cantidad, floa_ent_costo, floa_ent_precio, inte_ent_año, vcha_ent_almacen_origen)"
                           var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + var_almacen_Destino + "','T', " + CStr(var_numero_folio) + ",'" + rsaux2!vcha_Art_articulo_id + "', " + CStr(rsaux2!floa_emo_Cantidad) + "," + CStr(var_costo) + "," + CStr(var_precio) + ",2005,'" + var_almacen_origen + "')"
                           rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                           var_cadena = "insert into tb_entradas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_Art_Articulo_id, floa_ent_Cantidad, floa_ent_costo, floa_ent_precio, inte_ent_año, vcha_ent_almacen_origen)"
                           var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + var_almacen_Destino + "','T', " + CStr(var_numero_folio) + ",'" + rsaux2!vcha_Art_articulo_id + "', " + CStr(rsaux2!floa_emo_Cantidad) + "," + CStr(var_costo) + "," + CStr(Round(var_precio, 2)) + ",2005,'" + var_almacen_origen + "')"
                           rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  
                           var_cadena = "insert into tb_temporal_salidas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_Sal_numero, vcha_Art_Articulo_id, floa_sal_Cantidad, floa_sal_costo, floa_sal_precio, inte_Sal_año)"
                           var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + var_almacen_origen + "','T', " + CStr(var_numero_folio) + ",'" + rsaux2!vcha_Art_articulo_id + "', " + CStr(rsaux2!floa_emo_Cantidad) + "," + CStr(var_costo) + "," + CStr(var_precio) + ",2005)"
                           rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                           var_cadena = "insert into tb_salidas (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_Sal_numero, vcha_Art_Articulo_id, floa_sal_Cantidad, floa_sal_costo, floa_sal_precio, inte_Sal_año)"
                           var_cadena = var_cadena + "           values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + var_almacen_origen + "','T', " + CStr(var_numero_folio) + ",'" + rsaux2!vcha_Art_articulo_id + "', " + CStr(rsaux2!floa_emo_Cantidad) + "," + CStr(var_costo) + "," + CStr(var_precio) + ",2005)"
                           rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux10.Open "update tb_Temporal_entradas set floa_ent_Cantidad =  floa_ent_cantidad + " + CStr(rsaux2!floa_emo_Cantidad) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_ent_numero = " + CStr(var_numero_folio) + " and vcha_Art_articulo_id = '" + rsaux2!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           rsaux10.Open "update tb_entradas set floa_ent_Cantidad =  floa_ent_cantidad + " + CStr(rsaux2!floa_emo_Cantidad) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_ent_numero = " + CStr(var_numero_folio) + " and vcha_Art_articulo_id = '" + rsaux2!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  
                           rsaux10.Open "update tb_Temporal_salidas set floa_Sal_Cantidad =  floa_sal_cantidad + " + CStr(rsaux2!floa_emo_Cantidad) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_Sal_numero = " + CStr(var_numero_folio) + " and vcha_Art_articulo_id = '" + rsaux2!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           rsaux10.Open "update tb_salidas set floa_Sal_Cantidad =  floa_sal_cantidad + " + CStr(rsaux2!floa_emo_Cantidad) + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_Sal_numero = " + CStr(var_numero_folio) + " and vcha_Art_articulo_id = '" + rsaux2!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        
                        rsaux10.Open "update tb_entradas set dtim_ent_fecha  = " + var_fecha_fin + " where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_ent_numero = " + CStr(var_numero_folio) + " and vcha_Art_articulo_id = '" + rsaux2!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        rsaux10.Open "update tb_salidas set dtim_sal_fecha  = " + var_fecha_fin + "  where  vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_Sal_numero = " + CStr(var_numero_folio) + " and vcha_Art_articulo_id = '" + rsaux2!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        rsaux1.Close
                       
                        rsaux2.MoveNext
                  Wend
                  rsaux10.Open "update tb_Encabezado_movimientos set char_emo_estatus = 'I', dtim_emo_fecha = " + var_fecha_fin + ", inte_Emo_bloqueado = 0 where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'T' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  rsaux10.Open "insert into tb_folios_apartados (vcha_fol_folio_id) values ('" + rs!vcha_fol_folio_id + "')", cnn_compucaja, adOpenDynamic, adLockOptimistic
                  rsaux2.Close
               End If
               rsaux.Close
               rs.MoveNext
         Wend
         rs.Close
         
         
         
         'MsgBox "Se a terminado de cargar los movimientos"
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
    
End Sub


Private Sub Form_Load()
   Timer1.Enabled = True
   var_i = 0
   Top = 3000
   Left = 3500
   txt_inicio = Date
   txt_fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub Timer1_Timer()
   var_i = var_i + 1
   Call cmd_cargar_Click
   Me.txt_fecha = Now
End Sub

Private Sub txt_fin_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar la fecha"
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes = CDate(Me.txt_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_fin_LostFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = ""
End Sub

Private Sub txt_inicio_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar la fecha"
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_LostFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = ""
End Sub
