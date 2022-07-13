VERSION 5.00
Begin VB.Form frmsubir_polizas_intercompañias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subir poliza"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6330
      Picture         =   "frmsubir_polizas_intercompañias.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton Command9 
      Height          =   315
      Left            =   90
      Picture         =   "frmsubir_polizas_intercompañias.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Desmarcar Todos Alt + D"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmsubir_polizas_intercompañias.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   45
      TabIndex        =   15
      Top             =   360
      Width           =   6675
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos del movimiento "
      Height          =   1620
      Left            =   105
      TabIndex        =   10
      Top             =   480
      Width           =   6570
      Begin VB.TextBox txt_archivo 
         Enabled         =   0   'False
         Height          =   345
         Left            =   4380
         TabIndex        =   5
         Top             =   345
         Width           =   1350
      End
      Begin VB.TextBox txt_estatus 
         Enabled         =   0   'False
         Height          =   345
         Left            =   3090
         TabIndex        =   4
         Top             =   345
         Width           =   360
      End
      Begin VB.TextBox txt_piezas 
         Enabled         =   0   'False
         Height          =   345
         Left            =   4395
         TabIndex        =   9
         Top             =   1125
         Width           =   1350
      End
      Begin VB.TextBox txt_fecha 
         Enabled         =   0   'False
         Height          =   345
         Left            =   900
         TabIndex        =   8
         Top             =   1125
         Width           =   2430
      End
      Begin VB.TextBox txt_nombre_almacen 
         Enabled         =   0   'False
         Height          =   345
         Left            =   2280
         TabIndex        =   7
         Top             =   735
         Width           =   4230
      End
      Begin VB.TextBox txt_almacen 
         Enabled         =   0   'False
         Height          =   345
         Left            =   900
         TabIndex        =   6
         Top             =   735
         Width           =   1350
      End
      Begin VB.TextBox txt_numero 
         Height          =   345
         Left            =   900
         TabIndex        =   3
         Top             =   345
         Width           =   1350
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Factura:"
         Height          =   195
         Left            =   3690
         TabIndex        =   17
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estatus:"
         Height          =   195
         Left            =   2460
         TabIndex        =   16
         Top             =   420
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total piezas:"
         Height          =   195
         Left            =   3420
         TabIndex        =   14
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   210
         TabIndex        =   13
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Almacén:"
         Height          =   195
         Left            =   210
         TabIndex        =   12
         Top             =   810
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   210
         TabIndex        =   11
         Top             =   420
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmsubir_polizas_intercompañias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
Dim var_almacen_Destino As String
   If Trim(Me.txt_almacen) <> "" Then
      If Me.txt_estatus = "I" Then
         var_almacen_Destino = Me.txt_almacen
         var_clave_movimiento = "EI"
         var_numero_folio = CDbl(Me.txt_numero)
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
         'MsgBox var_cadena
         If rsaux10.State = 1 Then
            rsaux10.Close
         End If
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
                  var_cadena = var_cadena + " VALUES ('NEW', " + CStr(var_juego_libros_poliza) + ",'" + var_origen_poliza + "','" + var_categoria_poliza + "',TO_DATE('" + CStr(CDate(txt_fecha)) + "','DD/MM/YYYY'),'" + var_moneda_poliza + "',TO_DATE('" + CStr(CDate(txt_fecha)) + "','DD/MM/YYYY'),'A','" + var_segmento1_poliza + "','" + var_segmento2_poliza + "','" + var_segmento3_poliza + "','" + var_segmento4_poliza + "','" + var_segmento5_poliza + "','" + var_segmento6_poliza + "','" + var_segmento7_poliza + "'," + CStr(var_importe_precio) + ",0," + CStr(var_importe_precio) + ",0,1,'FACTURAS INTERCOMPAÑIA " + txt_archivo + "','FACTURA NUM: " + Me.txt_archivo + "','" + var_descripcion_poliza + "','POLIZA FACTURAS INTERCOMPAÑIA','POLIZA FACTURAS INTERCOMPAÑIA',1143)"
               Else
                  var_cadena = var_cadena + " VALUES ('NEW', " + CStr(var_juego_libros_poliza) + ",'" + var_origen_poliza + "','" + var_categoria_poliza + "',TO_DATE('" + CStr(CDate(txt_fecha)) + "','DD/MM/YYYY'),'" + var_moneda_poliza + "',TO_DATE('" + CStr(CDate(txt_fecha)) + "','DD/MM/YYYY'),'A','" + var_segmento1_poliza + "','" + var_segmento2_poliza + "','" + var_segmento3_poliza + "','" + var_segmento4_poliza + "','" + var_segmento5_poliza + "','" + var_segmento6_poliza + "','" + var_segmento7_poliza + "',0," + CStr(var_importe_precio) + ",0," + CStr(var_importe_precio) + ",1,'FACTURAS INTERCOMPAÑIA " + txt_archivo + "','FACTURA NUM: " + Me.txt_archivo + "','" + var_descripcion_poliza + "','POLIZA FACTURAS INTERCOMPAÑIA','POLIZA FACTURAS INTERCOMPAÑIA',1143)"
               End If
               'MsgBox var_cadena
               rsaux9.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
               rsaux11.MoveNext
         Wend
         rsaux11.Close
         cc = 1
         If cc = 1 Then
         rsaux11.Open "select sq_id_facturas.nextval from dual", cnnoracle, adOpenDynamic, adLockOptimistic
         var_consecutivo = rsaux11(0).Value
         rsaux11.Close
         rsaux11.Open "SELECT TOP 1 * FROM TB_ARCHIVO_COMPARACION WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_MOV_MOVIMIENTO_ID = 'EI' AND VCHA_COM_REFERENCIA = '" + Me.txt_archivo + "'", cnn, adOpenDynamic, adLockOptimistic
         var_proveedor_oracle = rsaux11!VCHA_COM_PROVEEDOR
         rsaux11.Close
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
         rsaux11.Open "SELECT vendor_site_id FROM po_vendor_sites_all@perpvia.vianney.com.mx Where vendor_id = '" + var_proveedor_oracle_2 + "'  AND vendor_site_id in (4070,2803,1125,1200,1202,1126,1519,1327,1520,3545,1127,1674,4383,2668,1529,1326,2669,2925,3755,1268,1324,3768,2392,9016,3332,1462, 1473,1737,1992, 1991,4618,5454,6407,4380,10626,9899,7244) AND ORG_ID = '" + var_unidad_oracle + "'", cnnoracle, adOpenDynamic, adLockOptimistic
         var_clave_proveedor_oracle = rsaux11!vendor_site_id
         rsaux11.Close
                        
                           'rsaux11.Open "select sum(FLOA_ENT_CANTIDAD * floa_ENT_costo) FROM  tb_temporal_entradas where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
         var_cadena = "SELECT     SUM(dbo.TB_TEMPORAL_ENTRADAS.FLOA_ENT_CANTIDAD * (dbo.TB_ARCHIVO_COMPARACION.FLOA_COM_COSTO*FLOA_COM_CANTIDAD_ENVIADA/FLOA_ENT_CANTIDAD )) AS Expr1"
         var_cadena = var_cadena + " FROM         dbo.TB_TEMPORAL_ENTRADAS INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_TEMPORAL_ENTRADAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_TEMPORAL_ENTRADAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_TEMPORAL_ENTRADAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_TEMPORAL_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND Dbo.TB_TEMPORAL_ENTRADAS.INTE_ENT_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN dbo.TB_ARCHIVO_COMPARACION ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ALM_ALMACEN_ID AND"
         var_cadena = var_cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA = dbo.TB_ARCHIVO_COMPARACION.VCHA_COM_REFERENCIA AND dbo.TB_TEMPORAL_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARCHIVO_COMPARACION.VCHA_ART_ARTICULO_ID WHERE  (dbo.TB_TEMPORAL_ENTRADAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_TEMPORAL_ENTRADAS.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') AND (dbo.TB_TEMPORAL_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "') AND (dbo.TB_TEMPORAL_ENTRADAS.INTE_ENT_NUMERO = " + Me.txt_numero + ") "
         rsaux11.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         var_importe_total = rsaux11(0).Value
         rsaux11.Close
         var_importe_total = var_importe_total * 1.16
                     
         var_cadena = "insert into IN_TB_FACTURAS_INT (INVOICE_ID,INVOICE_NUM,INVOICE_TYPE_LOOKUP_CODE,VENDOR_ID,VENDOR_SITE_ID,INVOICE_AMOUNT,INVOICE_CURRENCY_CODE,EXCHANGE_RATE_TYPE,EXCHANGE_DATE,EXCHANGE_RATE,Description,Source,GL_DATE,INVOICE_DATE,ORG_ID) values (" + CStr(var_consecutivo) + ",'" + Me.txt_archivo + "','STANDARD'," + CStr(var_proveedor_oracle_2) + "," + CStr(var_clave_proveedor_oracle) + "," + CStr(var_importe_total) + ",'MXP',null,null,null,'FACTURA DE RECEPCION NUM: " + Me.txt_archivo + "','FACTURA INTERCOMPAÑIAS',TO_DATE('" + CStr(CDate(txt_fecha)) + "','DD/MM/YYYY'),TO_DATE('" + CStr(CDate(txt_fecha)) + "','DD/MM/YYYY')," + var_unidad_oracle + ")"
         'MsgBox var_cadena
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
               var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",1,'ITEM'," + CStr(var_subimporte) + ", TO_DATE('" + CStr(CDate(txt_fecha)) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
            Else
               var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",1,'ITEM'," + CStr(var_subimporte) + ", TO_DATE('" + CStr(CDate(txt_fecha)) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "',NULL,NULL," + CStr(AWT_GROUP_ID) + ",NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
            End If
         Else
            If AWT_GROUP_ID = 0 Then
               var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",1,'ITEM'," + CStr(var_subimporte) + ", TO_DATE('" + CStr(CDate(txt_fecha)) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "','" + CStr(TAX_CODE) + "',NULL,NULL,NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
            Else
               var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",1,'ITEM'," + CStr(var_subimporte) + ", TO_DATE('" + CStr(CDate(txt_fecha)) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "','" + CStr(TAX_CODE) + "',NULL," + CStr(AWT_GROUP_ID) + ",NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
            End If
         End If
                        
        'MsgBox var_cadena
         rsaux11.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
         rsaux11.Open "select sq_id_lineas_factura.nextval from dual", cnnoracle, adOpenDynamic, adLockOptimistic
         var_consecutivo_linea = rsaux11(0).Value
         rsaux11.Close
                        
         var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description,AMOUNT_INCLUDES_TAX_FLAG,TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + ",2,'TAX'," + CStr(var_importe_iva) + ", TO_DATE('" + CStr(CDate(txt_fecha)) + "','DD/MM/YYYY'),'IMPUESTO',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
         rsaux11.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
         End If 'del cc
         MsgBox "Se a terminado de subir la poliza", vbOKOnly, "ATENCION"
         Me.txt_numero = ""
         Me.txt_almacen = ""
         Me.txt_nombre_almacen = ""
         Me.txt_estatus = ""
         Me.txt_archivo = ""
         Me.txt_fecha = ""
         Me.txt_piezas = ""
      Else
         MsgBox "El movimiento no a sido cerrado", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un movimiento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command9_Click()
   Me.txt_estatus = ""
   Me.txt_almacen = ""
   Me.txt_numero = ""
   Me.txt_nombre_almacen = ""
   Me.txt_piezas = ""
   Me.txt_archivo = ""
   Me.txt_numero.SetFocus
End Sub

Private Sub Form_Load()
   Top = 2500
   Left = 2000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub txt_numero_Change()
   Me.txt_estatus = ""
   Me.txt_almacen = ""
   Me.txt_nombre_almacen = ""
   Me.txt_fecha = ""
   Me.txt_piezas = ""
   Me.txt_archivo = ""
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_aceptar_pedidos.SetFocus
   End If
End Sub

Private Sub txt_numero_LostFocus()
    If Trim(Me.txt_numero) <> "" Then
       If IsNumeric(Me.txt_numero) Then
          rs.Open "SELECT * FROM TB_ENCABEZADO_MOVIMIENTOS WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = 'EI' AND INTE_EMO_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
          If Not rs.EOF Then
             rsaux.Open "SELECT * FROM TB_ALMACENES WHERE VCHA_ALM_ALMACEN_ID = '" + rs!VCHA_ALM_ALMACEN_ID + "'", cnn, adOpenDynamic, adLockOptimistic
             Me.txt_almacen = rs!VCHA_ALM_ALMACEN_ID
             Me.txt_nombre_almacen = rsaux!VCHA_ALM_NOMBRE
             Me.txt_fecha = CStr(rs!DTIM_EMO_FECHa)
             Me.txt_archivo = rs!vcha_Emo_referencia
             rsaux.Close
             rsaux.Open "SELECT SUM(FLOA_ENT_CANTIDAD) FROM TB_ENTRADAS WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = 'EI' AND INTE_ENT_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
             If Not rsaux.EOF Then
                Me.txt_piezas = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
             End If
             rsaux.Close
             Me.txt_estatus = IIf(IsNull(rs!char_Emo_estatus), "", rs!char_Emo_estatus)
          Else
             MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
          End If
          rs.Close
       Else
          MsgBox "Número de movimiento incorrecto", vbOKOnly, "ATENCION"
          Me.txt_almacen = ""
          Me.txt_nombre_almacen = ""
          Me.txt_fecha = ""
          Me.txt_piezas = ""
          Me.txt_estatus = ""
          Me.txt_archivo = ""
       End If
    Else
       Me.txt_almacen = ""
       Me.txt_nombre_almacen = ""
       Me.txt_fecha = ""
       Me.txt_piezas = ""
       Me.txt_estatus = ""
       Me.txt_archivo = ""
    End If
End Sub
