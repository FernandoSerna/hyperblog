VERSION 5.00
Begin VB.Form frmsubir_polizas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subir polizas"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_folio 
      Height          =   315
      Left            =   2025
      TabIndex        =   4
      Top             =   885
      Width           =   555
   End
   Begin VB.TextBox txt_destino 
      Height          =   315
      Left            =   1455
      TabIndex        =   3
      Top             =   870
      Width           =   555
   End
   Begin VB.TextBox txt_archivo 
      Enabled         =   0   'False
      Height          =   300
      Left            =   840
      TabIndex        =   2
      Top             =   870
      Width           =   600
   End
   Begin VB.TextBox txt_origen 
      Enabled         =   0   'False
      Height          =   285
      Left            =   105
      TabIndex        =   1
      Top             =   885
      Width           =   630
   End
   Begin VB.CommandButton cmd_subir_polizas 
      Caption         =   "Subir polizas"
      Height          =   750
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4395
   End
End
Attribute VB_Name = "frmsubir_polizas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_subir_polizas_Click()
     var_clave_movimiento = "EI"
     Dim var_fecha_movimiento As Date
     If rs.State = 1 Then
        rs.Close
     End If
     rs.Open "select * from archivo_ei", cnn, adOpenDynamic, adLockOptimistic
     While Not rs.EOF
           If rsaux.State = 1 Then
              rsaux.Close
           End If
           rsaux.Open "select * from tb_encabezado_movimientos where vcha_emo_referencia = '" + rs!codigo + "' and vcha_mov_movimiento_id = 'EI'", cnn, adOpenDynamic, adLockOptimistic
           'MsgBox rs!CODIGO
           If Not rsaux.EOF Then
           var_fecha_movimiento = rsaux!DTIM_EMO_FECHa
           Me.txt_folio = rsaux!INTE_EMO_NUMERO
           var_numero_folio = rsaux!INTE_EMO_NUMERO
           Me.txt_archivo = rs!codigo
           rsaux2.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + rsaux!VCHA_PRO_PROVEEDOR_ID + "'", cnn, adOpenDynamic, adLockOptimistic
           rsaux3.Open "select * from tb_almacenes where vcha_Alm_almacen_id = '" + rsaux!VCHA_ALM_ALMACEN_ID + "'", cnn, adOpenDynamic, adLockOptimistic
           Me.txt_destino = rsaux3!VCHA_ALM_NOMBRE
           rsaux3.Close
           Me.txt_origen = rsaux2!VCHA_UOR_NOMBRE
           rsaux2.Close
           If Not rsaux.EOF Then
              var_numero_folio = rsaux!INTE_EMO_NUMERO
              If var_clave_movimiento = "EI" Or var_clave_movimiento = "ETA" Then
                 If var_almacen_Destino = "RETEX" Then
                    var_cadena = "SELECT dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ENTRADAS.INTE_ENT_NUMERO, Sum(dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD) AS CANTIDAD, SUM(dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_CANTIDAD_ENVIADA * dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_COSTO / dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD * dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD) AS COSTO, SUM(dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_CANTIDAD_ENVIADA * dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_PRECIO / dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD * dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD) AS Precio FROM  dbo.TB_ENTRADAS INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND "
                    var_cadena = var_cadena + " dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENTRADAS.INTE_ENT_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN dbo.TB_ARCHIVO_COMPARACION ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA = dbo.TB_ARCHIVO_COMPARACION.VCHA_COM_REFERENCIA AND dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') "
                    var_cadena = var_cadena + " AND (dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') AND"
                    var_cadena = var_cadena + " (dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = 'EI') AND (dbo.TB_ENTRADAS.INTE_ENT_NUMERO = " + CStr(var_numero_folio) + ") GROUP BY dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ENTRADAS.INTE_ENT_NUMERO "
                 Else
                    If var_empresa = "31" Then
                       var_cadena = " SELECT dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ENTRADAS.INTE_ENT_NUMERO, SUM(dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD) AS CANTIDAD, SUM(dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD * dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_COSTO) AS COSTO, SUM(dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD * dbo.TB_ENTRADAS.FLOA_ENT_PRECIO) As Precio FROM dbo.TB_ENTRADAS INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENTRADAS.INTE_ENT_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN "
                       var_cadena = var_cadena + " dbo.TB_ARCHIVO_COMPARACION ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA = dbo.TB_ARCHIVO_COMPARACION.VCHA_COM_REFERENCIA AND dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') AND (dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = 'EI') AND (dbo.TB_ENTRADAS.INTE_ENT_NUMERO = " + CStr(var_numero_folio) + ") "
                       var_cadena = var_cadena + " GROUP BY dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ENTRADAS.INTE_ENT_NUMERO                "
                    Else
                       var_cadena = "SELECT  dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ENTRADAS.INTE_ENT_NUMERO, SUM(dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD) AS CANTIDAD, SUM(dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD * dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_COSTO) AS COSTO, SUM(dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD * dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_PRECIO) As Precio FROM dbo.TB_ENTRADAS INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND Dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENTRADAS.INTE_ENT_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN "
                       var_cadena = var_cadena + " dbo.TB_ARCHIVO_COMPARACION ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ALM_ALMACEN_ID AND  dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA = dbo.TB_ARCHIVO_COMPARACION.VCHA_COM_REFERENCIA AND dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') AND"
                       var_cadena = var_cadena + " (dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "') AND (dbo.TB_ENTRADAS.INTE_ENT_NUMERO = " + CStr(var_numero_folio) + ") GROUP BY dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ENTRADAS.INTE_ENT_NUMERO"
                    End If
                 End If
                 rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                 If var_clave_movimiento = "ETA" Then
                    concepto_1 = "ENTRADA TRASPASO " + CStr(var_numero_folio) + " "
                    CONCEPTO_2 = "ORIGEN " + Me.txt_origen + " " + Me.txt_archivo
                    CONCEPTO_3 = "RECEPCION POR TRASPASO " + Me.txt_destino
                    If var_empresa = "06" Then
                       If var_almacen_Destino = "ABPT" Then
                          rsaux11.Open "select * from tb_generador_polizas where poliza_id = '59'", cnnoracle, adOpenDynamic, adLockOptimistic
                       End If
                       If var_almacen_Destino = "Q0Z" Then
                          rsaux11.Open "select * from tb_generador_polizas where poliza_id = '19'", cnnoracle, adOpenDynamic, adLockOptimistic
                       End If
                       If var_almacen_Destino = "MPCOL" Then
                          rsaux11.Open "select * from tb_generador_polizas where poliza_id = '24'", cnnoracle, adOpenDynamic, adLockOptimistic
                       End If
                       If var_almacen_Destino = "MPCOC" Then
                          rsaux11.Open "select * from tb_generador_polizas where poliza_id = '28'", cnnoracle, adOpenDynamic, adLockOptimistic
                       End If
                       If var_almacen_Destino = "MPEDR" Then
                          rsaux11.Open "select * from tb_generador_polizas where poliza_id = '23'", cnnoracle, adOpenDynamic, adLockOptimistic
                       End If
                       If var_almacen_Destino = "PTMU" Or var_almacen_Destino = "CMU" Or var_almacen_Destino = "PMU" Then
                          rsaux11.Open "select * from tb_generador_polizas where poliza_id = '60'", cnnoracle, adOpenDynamic, adLockOptimistic
                       End If
                    End If
                    If var_empresa = "18" Then
                       rsaux11.Open "select * from tb_generador_polizas where empresa_id = '" + var_empresa + "' and poliza_id = '17'", cnnoracle, adOpenDynamic, adLockOptimistic
                    End If
                 End If
                 If var_clave_movimiento = "EI" Then
                    rsaux11.Open "select * from tb_generador_polizas where empresa_id = '" + var_empresa + "'", cnnoracle, adOpenDynamic, adLockOptimistic
                    concepto_1 = "FACTURA INTERCOMPAÑIA " + Me.txt_archivo
                    CONCEPTO_2 = "FACTURA NUM: " + Me.txt_archivo
                    CONCEPTO_3 = "POLIZA FACTURAS INTERCOMPAÑIA"
                 End If
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
                          If rsaux10.EOF Then
                             var_importe_precio = 0
                          Else
                             var_importe_precio = rsaux10!Precio
                          End If
                       Else
                          If rsaux10.EOF Then
                             var_importe_precio = 0
                          Else
                             var_importe_precio = IIf(IsNull(rsaux10!Costo), 0, rsaux10!Costo)
                          End If
                       End If
                       var_cadena = "InsERT INTO IN_TB_POLIZAS_INT (STATUS, SET_OF_BOOKS_ID, USER_JE_SOURCE_NAME, USER_JE_CATEGORY_NAME, ACCOUNTING_DATE, CURRENCY_CODE, DATE_CREATED, ACTUAL_FLAG,  SEGMENT1, SEGMENT2, SEGMENT3, SEGMENT4, SEGMENT5, SEGMENT6, SEGMENT7, ENTERED_DR, ENTERED_CR, ACCOUNTED_DR, ACCOUNTED_CR, GROUP_ID, REFERENCE4, REFERENCE5, REFERENCe10, REFERENCE1, REFERENCE2, CREATED_BY)"
                       var_fecha_movimiento = Format(var_fecha_movimiento, "Short Date")
                       If var_cargo_poliza = 1 Then
                          var_cadena = var_cadena + " VALUES ('NEW', " + CStr(var_juego_libros_poliza) + ",'" + var_origen_poliza + "','" + var_categoria_poliza + "',TO_DATE('" + CStr(var_fecha_movimiento) + "','DD/MM/YYYY'),'" + var_moneda_poliza + "',TO_DATE('" + CStr(var_fecha_movimiento) + "','DD/MM/YYYY'),'A','" + var_segmento1_poliza + "','" + var_segmento2_poliza + "','" + var_segmento3_poliza + "','" + var_segmento4_poliza + "','" + var_segmento5_poliza + "','" + var_segmento6_poliza + "','" + var_segmento7_poliza + "'," + CStr(IIf(IsNull(var_importe_precio), 0, var_importe_precio)) + ",0," + CStr(IIf(IsNull(var_importe_precio), 0, var_importe_precio)) + ",0,1,'" + concepto_1 + "','" + CONCEPTO_2 + "','" + var_descripcion_poliza + "','" + CONCEPTO_3 + "','" + CONCEPTO_3 + "',1143)"
                       Else
                          var_cadena = var_cadena + " VALUES ('NEW', " + CStr(var_juego_libros_poliza) + ",'" + var_origen_poliza + "','" + var_categoria_poliza + "',TO_DATE('" + CStr(var_fecha_movimiento) + "','DD/MM/YYYY'),'" + var_moneda_poliza + "',TO_DATE('" + CStr(var_fecha_movimiento) + "','DD/MM/YYYY'),'A','" + var_segmento1_poliza + "','" + var_segmento2_poliza + "','" + var_segmento3_poliza + "','" + var_segmento4_poliza + "','" + var_segmento5_poliza + "','" + var_segmento6_poliza + "','" + var_segmento7_poliza + "',0," + CStr(IIf(IsNull(var_importe_precio), 0, var_importe_precio)) + ",0," + CStr(IIf(IsNull(var_importe_precio), 0, var_importe_precio)) + ",1,'" + concepto_1 + "','" + CONCEPTO_2 + "','" + var_descripcion_poliza + "','" + CONCEPTO_3 + "','" + CONCEPTO_3 + "',1143)"
                          
                       End If
                       'MsgBox var_cadena
                       rsaux9.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                       rsaux11.MoveNext
                 Wend
                 rsaux11.Close
                 rsaux11.Open "select sq_id_facturas.nextval from dual", cnnoracle, adOpenDynamic, adLockOptimistic
                 var_consecutivo = rsaux11(0).Value
                 rsaux11.Close
                 rsaux11.Open "SELECT TOP 1 * FROM TB_ARCHIVO_COMPARACION WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND VCHA_COM_REFERENCIA = '" + Me.txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
                 var_proveedor_oracle = rsaux11!VCHA_COM_PROVEEDOR
                 rsaux11.Close
                 rsaux11.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_unidad_id = '" + var_proveedor_oracle + "'", cnn, adOpenDynamic, adLockOptimistic
                 var_empresa_emite = rsaux11!VCHA_EMP_EMPRESA_ID
                 var_proveedor_oracle_2 = rsaux11!vcha_uor_proveedor_oracle
                 rsaux11.Close
                 rsaux11.Open "select * from tb_empresas_cruzadas_oracle where vcha_emp_Empresa_emite = '" + var_empresa_emite + "' and vcha_emp_empresa_recibe = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                 var_unidad_oracle = rsaux11!vcha_emp_organizacion
                 rsaux11.Close
                 rsaux11.Open "SELECT vendor_site_id FROM po_vendor_sites_all@perpvia.vianney.com.mx Where vendor_id = '" + var_proveedor_oracle_2 + "'  AND vendor_site_id in (4070,2803,1125,1200,1202,1126,1519,1327,1520,3545,1127,1674,4383,2668,1529,1326,2669,2925,3755,1268,1324,3768,2392,9016,3332,1462, 1473,1737,1992, 1991,4618,5454,6407,4380,10626,9899,7244) AND ORG_ID = '" + var_unidad_oracle + "'", cnnoracle, adOpenDynamic, adLockOptimistic
                 var_clave_proveedor_oracle = rsaux11!vendor_site_id
                 rsaux11.Close
                 var_cadena = "SELECT     SUM(dbo.TB_TEMPORAL_ENTRADAS.FLOA_ENT_CANTIDAD * (dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_COSTO*FLOA_COM_CANTIDAD_ENVIADA/FLOA_ENT_CANTIDAD )) AS Expr1"
                 var_cadena = var_cadena + " FROM         dbo.TB_TEMPORAL_ENTRADAS INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_TEMPORAL_ENTRADAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_TEMPORAL_ENTRADAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_TEMPORAL_ENTRADAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_TEMPORAL_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND Dbo.TB_TEMPORAL_ENTRADAS.INTE_ENT_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN dbo.TB_ARCHIVO_COMPARACION ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ALM_ALMACEN_ID AND"
                 var_cadena = var_cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA = dbo.TB_ARCHIVO_COMPARACION.VCHA_COM_REFERENCIA AND dbo.TB_TEMPORAL_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ART_ARTICULO_ID WHERE  (dbo.TB_TEMPORAL_ENTRADAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_TEMPORAL_ENTRADAS.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') AND (dbo.TB_TEMPORAL_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "') AND (dbo.TB_TEMPORAL_ENTRADAS.INTE_ENT_NUMERO = " + Me.txt_folio + ") and FLOA_ENT_CANTIDAD > 0"
                 rsaux11.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                 var_importe_total = rsaux11(0).Value
                 rsaux11.Close
                 var_importe_total = var_importe_total * 1.16
                    
                 var_cadena = "insert into IN_TB_FACTURAS_INT (INVOICE_ID,INVOICE_NUM,INVOICE_TYPE_LOOKUP_CODE,VENDOR_ID,VENDOR_SITE_ID,INVOICE_AMOUNT,INVOICE_CURRENCY_CODE,EXCHANGE_RATE_TYPE,EXCHANGE_DATE,EXCHANGE_RATE,Description,Source,GL_DATE,INVOICE_DATE,ORG_ID) values (" + CStr(var_consecutivo) + ",'" + Me.txt_archivo + "','STANDARD'," + CStr(var_proveedor_oracle_2) + "," + CStr(var_clave_proveedor_oracle) + "," + CStr(IIf(IsNull(var_importe_total), 0, var_importe_total)) + ",'MXP',null,null,null,'FACTURA DE RECEPCION NUM: " + Me.txt_archivo + "','FACTURA INTERCOMPAÑIAS',TO_DATE('" + CStr(var_fecha_movimiento) + "','DD/MM/YYYY'),TO_DATE('" + CStr(var_fecha_movimiento) + "','DD/MM/YYYY')," + var_unidad_oracle + ")"
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
                 rsaux11.Open "select awt_group_id from po_vendors@perpvia.vianney.com.mx Where vendor_id = " + CStr(var_proveedor_oracle), cnnoracle, adOpenDynamic, adLockOptimistic
                 AWT_GROUP_ID = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                 rsaux11.Close
                 If TAX_CODE = 0 Then
                    If AWT_GROUP_ID = 0 Then
                       var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",1,'ITEM'," + CStr(var_subimporte) + ", TO_DATE('" + CStr(var_fecha_movimiento) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                    Else
                       var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",1,'ITEM'," + CStr(var_subimporte) + ", TO_DATE('" + CStr(var_fecha_movimiento) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "',NULL,NULL," + CStr(AWT_GROUP_ID) + ",NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                    End If
                 Else
                    If AWT_GROUP_ID = 0 Then
                       var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",1,'ITEM'," + CStr(IIf(IsNull(var_subimporte), 0, var_subimporte)) + ", TO_DATE('" + CStr(var_fecha_movimiento) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "','" + CStr(TAX_CODE) + "',NULL,NULL,NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                    Else
                       var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",1,'ITEM'," + CStr(IIf(IsNull(var_subimporte), 0, var_subimporte)) + ", TO_DATE('" + CStr(var_fecha_movimiento) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "','" + CStr(TAX_CODE) + "',NULL," + CStr(AWT_GROUP_ID) + ",NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                    End If
                 End If
                 rsaux11.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                 rsaux11.Open "select sq_id_lineas_factura.nextval from dual", cnnoracle, adOpenDynamic, adLockOptimistic
                 var_consecutivo_linea = rsaux11(0).Value
                 rsaux11.Close
                 var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description,AMOUNT_INCLUDES_TAX_FLAG,TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",2,'TAX'," + CStr(IIf(IsNull(var_importe_iva), 0, var_importe_iva)) + ", TO_DATE('" + CStr(var_fecha_movimiento) + "','DD/MM/YYYY'),'IMPUESTO',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                 rsaux11.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                 If var_empresa = "18" Then
                 End If
                 rsaux10.Close
             End If
          End If
          End If
          rsaux.Close
          rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub
