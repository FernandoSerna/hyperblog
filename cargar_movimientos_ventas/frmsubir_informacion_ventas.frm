VERSION 5.00
Begin VB.Form frmsubir_informacion_ventas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subir información ventas"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   555
      Left            =   135
      TabIndex        =   14
      Top             =   5430
      Width           =   2820
   End
   Begin VB.Frame Frame2 
      Caption         =   " Periodo notas de crédito "
      Height          =   1035
      Left            =   135
      TabIndex        =   9
      Top             =   3525
      Width           =   2895
      Begin VB.TextBox txt_fecha_fin 
         Height          =   330
         Left            =   915
         TabIndex        =   13
         Top             =   615
         Width           =   1380
      End
      Begin VB.TextBox txt_fecha_inicio 
         Height          =   330
         Left            =   915
         TabIndex        =   12
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   675
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   105
         TabIndex        =   10
         Top             =   285
         Width           =   420
      End
   End
   Begin VB.CommandButton cmd_subir_notas_credito 
      Caption         =   "Subir notas crédito"
      Height          =   570
      Left            =   75
      TabIndex        =   8
      Top             =   4635
      Width           =   2970
   End
   Begin VB.Frame Frame1 
      Caption         =   " Servidor "
      Height          =   1620
      Left            =   165
      TabIndex        =   3
      Top             =   90
      Width           =   2835
      Begin VB.OptionButton opt_tienda_cantia 
         Caption         =   "Tienda Cantia"
         Height          =   330
         Left            =   105
         TabIndex        =   7
         Top             =   1215
         Width           =   2550
      End
      Begin VB.OptionButton opt_vergel 
         Caption         =   "Vergel"
         Height          =   330
         Left            =   105
         TabIndex        =   6
         Top             =   915
         Width           =   2550
      End
      Begin VB.OptionButton opt_distribucion 
         Caption         =   "Distribución"
         Height          =   330
         Left            =   105
         TabIndex        =   5
         Top             =   615
         Width           =   2550
      End
      Begin VB.OptionButton opt_cdindustrial 
         Caption         =   "Ciudad Industrial"
         Height          =   330
         Left            =   105
         TabIndex        =   4
         Top             =   300
         Width           =   2550
      End
   End
   Begin VB.TextBox Text1 
      Height          =   510
      Left            =   120
      TabIndex        =   1
      Top             =   2925
      Width           =   2940
   End
   Begin VB.CommandButton cmd_subir_informacion 
      Caption         =   "Subir información"
      Height          =   690
      Left            =   105
      TabIndex        =   0
      Top             =   1845
      Width           =   2970
   End
   Begin VB.Label lbl_accion 
      Height          =   345
      Left            =   225
      TabIndex        =   2
      Top             =   2550
      Width           =   2760
   End
End
Attribute VB_Name = "frmsubir_informacion_ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim cnn_cdindustrial As ADODB.Connection
   Dim cnn_distribucion As ADODB.Connection
   Dim cnn_recuperacion As ADODB.Connection
   Dim cnn_cantia As ADODB.Connection
   Dim cnnoracle As ADODB.Connection
   Dim rs1 As ADODB.Recordset
   Dim rs2 As ADODB.Recordset
   Dim rs3 As ADODB.Recordset
   Dim rs4 As ADODB.Recordset
   Dim rs5 As ADODB.Recordset
   Dim rs6 As ADODB.Recordset
   Dim rs7 As ADODB.Recordset


Private Sub cantia_devoluciones()
      xz = 2
      If xz = 2 Then
         
         var_dia = CStr(Day(CDate(Me.txt_fecha_inicio)))
         var_mes = CStr(Month(CDate(Me.txt_fecha_inicio)))
         var_año = CStr(Year(CDate(Me.txt_fecha_inicio)))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         VAR_FECHA_INICIO = var_año + "-" + var_mes + "-" + var_dia
    
         var_dia = CStr(Day(CDate(Me.txt_fecha_fin)))
         var_mes = CStr(Month(CDate(Me.txt_fecha_fin)))
         var_año = CStr(Year(CDate(Me.txt_fecha_fin)))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         VAR_FECHA_FIN = var_año + "-" + var_mes + "-" + var_dia
         
         rs1.Open "DELETE FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
         var_origen = 6
         If rs1.State = 1 Then
            rs1.Close
         End If
         var_cadena = " SELECT     16 AS FLOA_CAR_PORCENTAJE_IVA, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, ISNULL(dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID, 'SIN EMPRESA') AS VCHA_EMP_EMPRESA_ID, ISNULL(dbo.TB_EMPRESAS.VCHA_EMP_NOMBRE, 'SIN EMPRESA') AS VCHA_EMP_NOMBRE, ISNULL(dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID, 'SIN UNIDAD') AS VCHA_UOR_UNIDAD_ID, ISNULL(dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_NOMBRE, 'SIN UNIDAD') AS VCHA_UOR_NOMBRE, REPLACE(REPLACE(ISNULL(dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_NOMBRE, 'SIN ESTABLECIMIENTO'), '''', '´'), '´', '') AS VCHA_ESB_NOMBRE, { fn WEEK(dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA) } AS SEMANA, 'NC' AS VTA_DESCRIPCION_DOCUMENTO, ISNULL(dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID, '') AS VCHA_ART_ARTICULO_ID, ISNULL(dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, '') AS VCHA_ART_NOMBRE_ESPAÑOL, ISNULL(dbo.TB_CATALOGOS.VCHA_CAT_CATALOGO_ID, 'SIN CATALOGO') AS VCHA_CAT_CATALOGO_ID, "
         var_cadena = var_cadena + " ISNULL(dbo.TB_CATALOGOS.VCHA_CAT_NOMBRE, ' SIN CATALOGO ') AS VCHA_CAT_NOMBRE, ISNULL(dbo.TB_DISEÑOS.VCHA_DIS_DISEÑO_ID, 'SIN DISEÑO') AS VCHA_DIS_DISEÑO_ID, ISNULL(dbo.TB_DISEÑOS.VCHA_DIS_NOMBRE, ' SIN DISEÑO ') AS VCHA_DIS_NOMBRE, ISNULL(dbo.TB_LINEAS.VCHA_LIN_LINEA_ID, 'SIN LINEA') AS VCHA_LIN_LINEA_ID, ISNULL(dbo.TB_LINEAS.VCHA_LIN_NOMBRE, 'SIN LINEA') AS VCHA_LIN_NOMBRE, ISNULL(dbo.TB_TALLAS.VCHA_TAL_TALLA_ID, 'SIN TALLA') AS VCHA_TAL_TALLA_ID, ISNULL(dbo.TB_TALLAS.VCHA_TAL_NOMBRE, 'SIN TALLA') AS VCHA_TAL_NOMBRE, ISNULL(dbo.TB_LICENCIAS.VCHA_LIC_LICENCIA_ID, 'SIN LICENCIA') AS VCHA_LIC_LICENCIA_ID, ISNULL(dbo.TB_LICENCIAS.VCHA_LIC_NOMBRE, 'SIN LICENCIA') AS VCHA_LIC_NOMBRE, ISNULL(dbo.TB_ARTICULOS.VCHA_ART_NUMERO_LIC, 'SIN NUMERO DE LICENCIA') AS VCHA_ART_NUMERO_LIC, ISNULL(dbo.TB_ARTICULOS.MONE_ART_PRECIO_BASE, 0) AS MONE_ART_PRECIO_BASE, ISNULL(dbo.TB_ARTICULOS.DTIM_ART_FECHA_ALTA, "
         var_cadena = var_cadena + " GETDATE()) AS DTIM_ART_FECHA_ALTA, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO AS inte_car_numero, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_ZON_DESCRIPCION, 'SIN ZONA') AS VCHA_ZON_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, ISNULL(dbo.VW_CLIENTES.VCHA_CAN_CANAL_VENTA_ID, 'SIN CANAL') AS VCHA_CAN_CANAL_VENTA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_CAN_NOMBRE, 'SIN CANAL') AS VCHA_CAN_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_AGE_AGENTE_ID, 'SIN AGENTE') AS VCHA_AGE_AGENTE_ID, ISNULL(dbo.VW_CLIENTES.VCHA_AGE_NOMBRE, 'SIN AGENTE') AS VCHA_AGE_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_RUT_RUTA_ID, 'SIN RUTA') AS VCHA_RUT_RUTA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_RUT_NOMBRE, 'SIN RUTA') AS VCHA_RUT_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_ZON_ZONA_ID, 'SIN ZONA') AS VCHA_ZON_ZONA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_ZON_DESCRIPCION, 'SIN ZONA') AS VCHA_ZON_DESCRIPCION, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID, 'SIN CLIENTE') AS VCHA_CLI_CLAVE_ID,"
         var_cadena = var_cadena + " REPLACE(REPLACE(ISNULL(dbo.VW_CLIENTES.VCHA_CLI_NOMBRE, 'SIN CLIENTE'), '''', '´'), '´', '') AS VCHA_CLI_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_CLAVE_UNIFICADA_ID, '0') AS VCHA_CLI_CLAVE_UNIFICADA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_RFC, 'SIN RFC') AS VCHA_CLI_RFC, ISNULL(dbo.VW_CLIENTES.VCHA_TIT_TITULAR_ID, 'SIN TITULAR') AS VCHA_TIT_TITULAR_ID, REPLACE(REPLACE(ISNULL(dbo.VW_CLIENTES.VCHA_TIT_NOMBRE, 'SIN TITULAR'), '''', '´'), '´', '') AS VCHA_TIT_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_GAC_GRUPO_ACTUAL_ID, 'SIN GRUPO ') AS VCHA_GAC_GRUPO_ACTUAL_ID, REPLACE(REPLACE(ISNULL(dbo.VW_CLIENTES.VCHA_GAC_NOMBRE, 'SIN GRUPO'), '''', '´'), '´', '') AS VCHA_GAC_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_CP, 'SIN CP') AS VCHA_CLI_CP, ISNULL(dbo.VW_CLIENTES.VCHA_EST_ESTADO_ID, 'SIN ESTADO') AS VCHA_EST_ESTADO_ID, ISNULL(dbo.VW_CLIENTES.VCHA_EST_NOMBRE, 'SIN ESTADO') AS VCHA_EST_NOMBRE, inte_ent_consecutivo_tabla, "
         var_cadena = var_cadena + " ISNULL(dbo.VW_CLIENTES.VCHA_CIU_CIUDAD_ID, 'SIN CIUDAD') AS VCHA_CIU_CIUDAD_ID, ISNULL(dbo.VW_CLIENTES.VCHA_CIU_NOMBRE, 'SIN CIUDAD') AS VCHA_CIU_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_MUN_MUNICIPIO_ID, 'SIN MUNICIPIO') AS VCHA_MUN_MUNICIPIO_ID, ISNULL(dbo.VW_CLIENTES.VCHA_MUN_NOMBRE, 'SIN MUNICIPIO') AS VCHA_MUN_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_COL_COLONIA_ID, 'SIN COLONIA') AS VCHA_COL_COLONIA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_COL_NOMBRE, 'SIN COLONIA') AS VCHA_COL_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_MON_MONEDA_ID, 'SIN MONEDA') AS VCHA_MON_MONEDA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_MON_DIVISA, 'SIN MONEDA') AS VCHA_MON_DIVISA, ISNULL(dbo.VW_CLIENTES.VCHA_ZON_DESCRIPCION, 'SIN ZONA') AS Expr1, ISNULL(dbo.VW_CLIENTES.VCHA_PAI_PAIS_ID, '') AS VCHA_PAI_PAIS_ID, ISNULL(dbo.VW_CLIENTES.VCHA_PAI_NOMBRE, '') AS VCHA_PAI_NOMBRE, dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, dbo.TB_ENTRADAS.FLOA_ENT_COSTO, dbo.TB_ENTRADAS.FLOA_ENT_PRECIO "
         var_cadena = var_cadena + " FROM dbo.TB_CATALOGOS RIGHT OUTER JOIN dbo.TB_LINEAS RIGHT OUTER JOIN dbo.TB_ENTRADAS INNER JOIN dbo.TB_UNIDADESORGANIZACIONALES INNER JOIN dbo.TB_EMPRESAS INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID ON dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID ON dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENTRADAS.INTE_ENT_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN "
         var_cadena = var_cadena + " dbo.TB_ALMACENES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID LEFT OUTER JOIN dbo.TB_ESTABLECIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID = dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID LEFT OUTER Join dbo.VW_CLIENTES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_CLI_CLAVE_ID = dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID LEFT OUTER JOIN dbo.TB_LICENCIAS ON dbo.TB_ARTICULOS.VCHA_LIC_LICENCIA_ID = dbo.TB_LICENCIAS.VCHA_LIC_LICENCIA_ID LEFT OUTER JOIN dbo.TB_TALLAS ON dbo.TB_ARTICULOS.VCHA_TAL_TALLA_ID = dbo.TB_TALLAS.VCHA_TAL_TALLA_ID ON dbo.TB_LINEAS.VCHA_LIN_LINEA_ID = dbo.TB_ARTICULOS.VCHA_LIN_LINEA_ID LEFT OUTER JOIN dbo.TB_DISEÑOS ON dbo.TB_ARTICULOS.VCHA_DIS_DISEÑO_ID = dbo.TB_DISEÑOS.VCHA_DIS_DISEÑO_ID ON dbo.TB_CATALOGOS.VCHA_CAT_CATALOGO_ID = dbo.TB_ARTICULOS.VCHA_ART_CATALOGO_VIGENTE WHERE (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'CC_4') AND "
         var_cadena = var_cadena + " (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= CONVERT(DATETIME, '" + VAR_FECHA_INICIO + "', 102)) AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA < CONVERT(DATETIME, '" + VAR_FECHA_FIN + "', 102)) AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN = 0) "
         rs1.Open var_cadena, cnn_cdindustrial, adOpenDynamic, adLockOptimistic
         While Not rs1.EOF
               var_dia = CStr(Day(CDate(rs1!DTIM_emo_FECHA)))
               var_mes = CStr(Month(CDate(rs1!DTIM_emo_FECHA)))
               var_año = CStr(Year(CDate(rs1!DTIM_emo_FECHA)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_factura = var_dia + "/" + var_mes + "/" + var_año
               var_dia = CStr(Day(CDate(rs1!DTIM_ART_FECHA_ALTA)))
               var_mes = CStr(Month(CDate(rs1!DTIM_ART_FECHA_ALTA)))
               var_año = CStr(Year(CDate(rs1!DTIM_ART_FECHA_ALTA)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_alta = var_dia + "/" + var_mes + "/" + var_año
               var_canal_venta = Trim(rs1!VCHA_CAN_CANAL_VENTA_ID)
               If var_canal_venta = "" Then
                  var_canal_venta = "SIN CANAL"
               End If
               var_nombre_canal = Trim(rs1!VCHA_CAN_NOMBRE)
               If var_nombre_canal = "" Then
                  var_nombre_canal = "SIN CANAL"
               End If
               var_agente = Trim(rs1!VCHA_AGE_AGENTE_ID)
               If var_agente = "" Then
                  var_agente = "SIN AGENTE"
               End If
               var_nombre_agente = Trim(rs1!VCHA_AGE_NOMBRE)
               If var_nombre_agente = "" Then
                  var_nombre_agente = "SIN AGENTE"
               End If
               var_ruta = Trim(rs1!VCHA_RUT_RUTA_ID)
               If var_ruta = "" Then
                  var_ruta = "SIN RUTA"
               End If
               var_nombre_ruta = Trim(rs1!VCHA_RUT_NOMBRE)
               If var_nombre_ruta = "" Then
                  var_nombre_ruta = "SIN RUTA"
               End If
               var_zona = Trim(rs1!VCHA_ZON_ZONA_ID)
               If var_zona = "" Then
                  var_zona = "SIN ZONA"
               End If
               var_nombre_zona = Trim(rs1!VCHA_ZON_NOMBRE)
               If var_nombre_zona = "" Then
                  var_nombre_zona = "SIN ZONA"
               End If
               var_cliente = Trim(rs1!VCHA_CLI_CLAVE_ID)
               If var_cliente = "" Then
                  var_cliente = "SIN CLIENTE"
               End If
               var_cliente_unfo = Trim(VCHA_CLI_CLAVE_UNIFICADA_ID)
               If var_cliente_unfo = "" Then
                  var_cliente_unfo = "0"
               End If
               var_nombre_cliente = Trim(rs1!VCHA_CLI_NOMBRE)
               If var_nombre_cliente = "" Then
                  var_nombre_cliente = "SIN CLIENTE"
               End If
               var_rfc = Trim(rs1!VCHA_CLI_RFC)
               var_titular = Trim(rs1!VCHA_TIT_TITULAR_ID)
               If var_titular = "" Then
                  var_titular = "SIN TITULAR"
               End If
               var_nombre_titular = Trim(rs1!VCHA_TIT_NOMBRE)
               If var_nombre_titular = "" Then
                  var_nombre_titular = "SIN TITULAR"
               End If
               var_grupo = Trim(rs1!VCHA_GAC_GRUPO_ACTUAL_ID)
               If var_grupo = "" Then
                  var_grupo = "SIN GRUPO"
               End If
               var_nombre_grupo = Trim(rs1!VCHA_GAC_NOMBRE)
               If var_nombre_grupo = "" Then
                  var_nombre_grupo = "SIN GRUPO"
               End If
               var_establecimiento = ""
               If var_establecimiento = "" Then
                  var_establecimiento = "SIN ESTABLECIMIENTO"
               End If
               var_nombre_establecimiento = ""
               If var_nombre_establecimiento = "" Then
                  var_nombre_establecimiento = "SIN ESTABLECIMIENTO"
               End If
               var_cp = Trim(rs1!VCHA_CLI_CP)
               If var_cp = "" Then
                  var_cp = "SIN CP"
               End If
               var_estado = Trim(rs1!VCHA_EST_ESTADO_ID)
               If var_estado = "" Then
                  var_estado = "SIN ESTADO"
               End If
               var_nombre_estado = Trim(rs1!VCHA_EST_NOMBRE)
               If var_nombre_estado = "" Then
                  var_nombre_estado = "SIN ESTADO"
               End If
               var_ciudad = Trim(rs1!VCHA_CIU_CIUDAD_ID)
               If var_ciudad = "" Then
                  var_ciudad = "SIN CIUDAD"
               End If
               var_nombre_ciudad = Trim(rs1!VCHA_CIU_NOMBRE)
               If var_nombre_ciudad = "" Then
                  var_nombre_ciudad = "SIN CIUDAD"
               End If
               var_municipio = Trim(rs1!VCHA_MUN_MUNICIPIO_ID)
               If var_municipio = "" Then
                  var_municipio = "SIN MUNICIPIO"
               End If
               var_nombre_municipio = Trim(rs1!VCHA_MUN_NOMBRE)
               If var_nombre_municipio = "" Then
                  var_nombre_municipio = "SIN MUNICIPIO"
               End If
               var_colonia = Trim(rs1!vcha_col_colonia_id)
               If var_colonia = "" Then
                  var_colonia = "SIN COLONIA"
               End If
               var_nombre_colonia = Trim(rs1!VCHA_COL_NOMBRE)
               If var_nombre_colonia = "" Then
                  var_nombre_colonia = "SIN COLONIA"
               End If
               var_articulo = Trim(rs1!VCHA_ART_ARTICULO_ID)
               If var_articulo = "" Then
                  var_articulo = "SIN ARTICULO"
               End If
               var_nombre_articulo = Trim(rs1!VCHA_aRT_NOMBRE_ESPAÑOL)
               If var_nombre_articulo = "" Then
                  var_nombre_articulo = "SIN ARTICULO"
               End If
               var_catalogo = Trim(rs1!VCHA_CAT_CATALOGO_ID)
               If var_catalogo = "" Then
                  var_catalogo = "SIN CATALOGO"
               End If
               var_nombre_catalogo = Trim(rs1!VCHA_CAT_NOMBRE)
               If var_nombre_catalogo = "" Then
                  var_nombre_catalogo = "SIN CATALOGO"
               End If
               var_diseño = Trim(rs1!VCHA_DIS_DISEÑO_ID)
               If var_diseño = "" Then
                  var_diseño = "SIN DISEÑO"
               End If
               var_nombre_diseño = Trim(rs1!VCHA_DIS_NOMBRE)
               If var_nombre_diseño = "" Then
                  var_nombre_diseño = "SIN DISEÑO"
               End If
               var_linea = Trim(rs1!VCHA_LIN_LINEA_ID)
               If var_linea = "" Then
                  var_linea = "SIN LINEA"
               End If
               var_nombre_linea = Trim(rs1!VCHA_LIN_NOMBRE)
               If var_nombre_linea = "" Then
                  var_nombre_linea = "SIN LINEA"
               End If
               var_talla = Trim(rs1!vcha_Tal_talla_id)
               If var_talla = "" Then
                  var_talla = "SIN TALLA"
               End If
               var_nombre_talla = Trim(rs1!VCHA_TAL_NOMBRE)
               If var_nombre_talla = "" Then
                  var_nombre_talla = "SIN TALLA"
               End If
               var_licencia = Trim(rs1!vcha_lic_licencia_id)
               If var_licencia = "" Then
                  var_licencia = "SIN LICENCIA"
               End If
               var_nombre_licencia = Trim(rs1!vcha_lic_nombre)
               If var_nombre_licencia = "" Then
                  var_nombre_licencia = "SIN LICENCIA"
               End If
               var_numero_licencia = Trim(rs1!VCHA_ART_NUMERO_LIC)
               If var_numero_licencia = "" Then
                  var_numero_licencia = "SIN NUMERO DE LICENCIA"
               End If
               var_pais = Trim(rs1!VCHA_PAI_PAIS_ID)
               If var_pais = "" Then
                  var_pais = "SIN PAIS"
               End If
               var_nombre_pais = Trim(rs1!vcha_pai_nombre)
               If var_nombre_pais = "" Then
                  var_nombre_pais = "SIN PAIS"
               End If
        
         
               var_descuento_1 = 0
               var_descuento_2 = 0
               VAR_PRECIO = rs1!FLOA_ent_PRECIO - (rs1!FLOA_ent_PRECIO * (var_descuento_1 / 100))
               VAR_PRECIO = VAR_PRECIO - (VAR_PRECIO * (var_descuento_2 / 100))
               VAR_PORCENTAJE_IVA = IIf(IsNull(rs1!FLOA_CAR_PORCENTAJE_IVA), 0, rs1!FLOA_CAR_PORCENTAJE_IVA)
               var_precio_2 = VAR_PRECIO * ((1 + (VAR_PORCENTAJE_IVA / 100)))
               var_importe_iva = var_precio_2 - VAR_PRECIO
               VAR_PRECIO = var_precio_2
               var_precio_base = IIf(IsNull(rs1!MONE_ART_PRECIO_BASE), 0, rs1!MONE_ART_PRECIO_BASE) * ((1 + (VAR_PORCENTAJE_IVA / 100)))
                         
               var_cadena = "insert into  VT_TB_VENTAS  (VTA_INDICE_ID, VTA_ORIGEN_ID, VTA_EMPRESA_ID, VTA_EMPRESA, VTA_UNIDAD_ORGANIZACIONAL_ID, VTA_UNIDAD_ORGANIZACIONAL,VTA_TIENDA_ID,VTA_TIENDA,VTA_CANAL_ID, VTA_CANAL, VTA_REGION,VTA_AGENTE_ID,VTA_AGENTE,VTA_RUTA_ID,VTA_RUTA,VTA_ZONA_ID,VTA_ZONA,VTA_CLIENTE_ID,VTA_CLIENTE_ID_UNFO,VTA_CLIENTE,VTA_RFC_CLIENTE,VTA_TITULAR_ID,VTA_TITULAR,VTA_GRUPO_ID,VTA_GRUPO, "
               var_cadena = var_cadena + " VTA_ESTABLECIMIENTO_ID,VTA_ESTABLECIMIENTO,VTA_CODIGO_POSTAL,VTA_ESTADO_ID,VTA_ESTADO,VTA_CIUDAD_ID,VTA_CIUDAD,VTA_MUNICIPIO_ID,VTA_MUNICIPIO, VTA_COLONIA_ID,VTA_COLONIA,VTA_FECHA, VTA_SEMANA,VTA_ID_ENCABEZADO_SID,VTA_ID_DETALLE_SID,VTA_MOVIMIENTO_ID,VTA_TIPO_MOVIMIENTO_ID, VTA_DOCUMENTO, VTA_DESCRIPCION_DOCUMENTO, VTA_NUMERO_DOCUMENTO,VTA_SERIE,VTA_PLAZO,VTA_ARTICULO_ID,                 VTA_ARTICULO, "
               var_cadena = var_cadena + "VTA_CATALOGO_ID,VTA_CATALOGO,VTA_DISENIO_ID,VTA_DISENIO, VTA_FAMILIA_ID, VTA_FAMILIA, VTA_LINEA_ID,VTA_LINEA, VTA_TALLA_ID,VTA_TALLA,VTA_LICENCIA_ID,VTA_LICENCIA,VTA_NUMERO_LICENCIA, "
               var_cadena = var_cadena + "VTA_PRECIO_BASE, VTA_IMPUESTO_PORCIENTO ,VTA_FECHA_ALTA_CODIGO, VTA_CANTIDAD, VTA_COSTO, VTA_PRECIO_LISTA_A, VTA_PRECIO_LISTA_B, VTA_PRECIO_LISTA_C, VTA_DESCUENTO1 ,VTA_DESCUENTO2, VTA_DESCUENTO3, VTA_TIPO_CAMBIO, VTA_MONEDA ,VTA_IMPORTE,VTA_FLETE_ID,VTA_IMPORTE_FLETE,VTA_IMPUESTO_FLETE,VTA_MOVIMIENTO_INDEX,VTA_TOTAL_MOVIMIENTOS,VTA_CLAVE_REGISTRO, VTA_PAIS_ID, VTA_PAIS, VTA_IMPUESTO, precio) VALUES"
               var_cadena = var_cadena + "(sq_ventas_id.nextval," + CStr(var_origen) + ",'" + Trim(rs1!VCHA_EMP_EMPRESA_ID) + "', '" + Trim(rs1!VCHA_EMP_NOMBRE) + "','" + Trim(rs1!VCHA_UOR_UNIDAD_ID) + "','" + Trim(rs1!VCHA_UOR_NOMBRE) + "','" + Trim(rs1!VCHA_ALM_ALMACEN_ID) + "', '" + Trim(rs1!VCHA_ALM_NOMBRE) + "','" + Trim(var_canal_venta) + "','" + Trim(var_nombre_canal) + "','','" + Trim(var_agente) + "', '" + Trim(var_nombre_agente) + "','" + Trim(var_ruta) + "','" + Trim(var_nombre_ruta) + "', '" + Trim(var_zona) + "', '" + Trim(var_nombre_zona) + "', '" + Trim(var_cliente) + "','" + Trim(VCHA_CLI_CLAVE_UNIFICADA_ID) + "','" + Trim(var_nombre_cliente) + "', '" + Trim(var_rfc) + "','" + Trim(var_titular) + "', '" + Trim(var_nombre_titular) + "', '" + Trim(var_grupo) + "', '" + Trim(var_nombre_grupo) + "',"
               var_cadena = var_cadena + "'" + Trim(var_establecimiento) + "', '" + Trim(var_nombre_establecimiento) + "', '" + Trim(var_cp) + "','" + Trim(var_estado) + "', '" + Trim(var_nombre_estado) + "', '" + Trim(var_ciudad) + "','" + Trim(var_nombre_ciudad) + "', '" + Trim(var_municipio) + "','" + Trim(var_nombre_municipio) + "','" + Trim(var_colonia) + "', '" + Trim(var_nombre_colonia) + "', TO_DATE('" + var_fecha_factura + "','DD/MM/YYYY')," + CStr(rs1!SEMANA) + ", '" + CStr(rs1!inte_ent_consecutivo_tabla) + Trim("sqlquezada2") + Trim("sicantia") + "', '" + CStr(rs1!inte_ent_consecutivo_tabla) + Trim("sqlquezada2") + Trim("sidcantia") + "', 'NC',  'NC',          'NC',         'NC'                 , " + CStr(rs1!inte_Car_numero) + ",'TC',0,'" + Trim(rs1!VCHA_ART_ARTICULO_ID) + "', '" + Trim(rs1!VCHA_aRT_NOMBRE_ESPAÑOL) + "',"
               var_cadena = var_cadena + "'" + Trim(var_catalogo) + "','" + Trim(var_nombre_catalogo) + "', '" + Trim(var_diseño) + "', '" + Trim(var_nombre_diseño) + "','','','" + Trim(var_linea) + "','" + Trim(var_nombre_linea) + "','" + Trim(var_talla) + "','" + Trim(var_nombre_talla) + "','" + Trim(var_licencia) + "', '" + Trim(var_nombre_licencia) + "','" + Trim(var_numero_licencia) + "',"
               
               var_cadena = var_cadena + CStr(var_precio_base) + "," + CStr(rs1!FLOA_CAR_PORCENTAJE_IVA) + ",TO_DATE('" + var_fecha_alta + "','DD/MM/YYYY')," + CStr(rs1!floa_ent_cantidad) + "," + CStr(IIf(IsNull(rs1!FLOA_ent_COSTO), 0, rs1!FLOA_ent_COSTO)) + "," + CStr(VAR_PRECIO) + ",0,0," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(1) + ",'" + Trim(rs1!vcha_mon_divisa) + "'," + CStr(rs1!floa_ent_cantidad * VAR_PRECIO) + ",0,0,0,'','','" + Trim(rs1!VCHA_EMP_EMPRESA_ID) + Trim(rs1!VCHA_UOR_UNIDAD_ID) + "NC" + Trim("TC") + CStr(rs1!inte_Car_numero) + "_" + CStr(rs1!inte_ent_consecutivo_tabla) + "','" + Trim(var_pais) + "','" + Trim(var_nombre_pais) + "', " + CStr(rs1!floa_ent_cantidad * var_importe_iva) + "," + CStr(rs1!FLOA_ent_PRECIO) + ")"
               
            
               rs2.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
            
               rs3.Open "SELECT * FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS WHERE  VCHA_EMP_EMPRESA_ID = '" + rs1!VCHA_EMP_EMPRESA_ID + "' AND VCHA_CAR_TIPO_DOCUMENTO = 'NC' AND VCHA_CAR_DOCUMENTO = 'NC' AND VCHA_SER_SERIE_ID = 'TC' AND VCHA_CLI_CLAVE_ID = '" + rs1!VCHA_CLI_CLAVE_ID + "' AND INTE_CAR_NUMERO = " + CStr(rs1!inte_Car_numero), cnn_cdindustrial, adOpenDynamic, adLockOptimistic
               If rs3.EOF Then
                  rs4.Open "INSERT INTO TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS (VCHA_EMP_EMPRESA_ID,  VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO) VALUES ('" + rs1!VCHA_EMP_EMPRESA_ID + "','NC','NC','TC','" + rs1!VCHA_CLI_CLAVE_ID + "'," + CStr(rs1!inte_Car_numero) + ")", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
               End If
               rs3.Close
               
               rs1.MoveNext
               var_i = var_i + 1
               Text1 = var_i
               Me.Refresh
               Me.Text1.Refresh
               'Me.lbl_accion.Refresh
         Wend
         rs1.Close
      
         rs2.Open "select distinct VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
         var_i = 0
         While Not rs2.EOF
               rs4.Open "UPDATE TB_ENCABEZADO_MOVIMIENTOS SET INTE_EMO_NUMERO_ORIGEN = 1 WHeRE vcha_emp_empresa_id = '" + rs2!VCHA_EMP_EMPRESA_ID + "' and vcha_mov_movimiento_id = 'CC_4' and inte_emo_numero = " + CStr(rs2!inte_Car_numero), cnn_cdindustrial, adOpenDynamic, adLockOptimistic
               rs2.MoveNext
               var_i = var_i + 1
               Text1.Text = var_id
               Me.Refresh
               Me.Text1.Refresh
               Me.lbl_accion.Refresh
         Wend
         rs2.Close
         x = 1
         If x = 1 Then
            var_origen = 6
            If rs4.State = 1 Then
               rs4.Close
            End If
            rs4.Open "SELECT VCHA_EMP_EMPRESA_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, vcha_cli_clave_id from TB_ENCABEZADO_movimientos WHERE INTE_EMO_NUMERO_ORIGEN = 2 and VCHA_MOV_MOVIMIENTO_id = 'CC_4'", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
            While Not rs4.EOF
                  
                  rs1.Open "DELETE FROM vt_tb_ventas where VTA_ORIGEN_ID = " + CStr(var_origen) + " and VTA_EMPRESA_ID = '" + rs4!VCHA_EMP_EMPRESA_ID + "' And vta_documento = 'NC' and vta_serie = 'TC' and vta_numero_documento = " + CStr(rs4!INTE_EMO_NUMERO) + " and vta_cliente_id = '" + rs4!VCHA_CLI_CLAVE_ID + "'", cnnoracle, adOpenDynamic, adLockOptimistic
                  rs4.MoveNext
            Wend
            rs4.Close
            rs1.Open "DELETE FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
            var_origen = 6
            If rs1.State = 1 Then
               rs1.Close
            End If
            var_cadena = " SELECT     16 AS FLOA_CAR_PORCENTAJE_IVA, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, ISNULL(dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID, 'SIN EMPRESA') AS VCHA_EMP_EMPRESA_ID, ISNULL(dbo.TB_EMPRESAS.VCHA_EMP_NOMBRE, 'SIN EMPRESA') AS VCHA_EMP_NOMBRE, ISNULL(dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID, 'SIN UNIDAD') AS VCHA_UOR_UNIDAD_ID, ISNULL(dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_NOMBRE, 'SIN UNIDAD') AS VCHA_UOR_NOMBRE, REPLACE(REPLACE(ISNULL(dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_NOMBRE, 'SIN ESTABLECIMIENTO'), '''', '´'), '´', '') AS VCHA_ESB_NOMBRE, { fn WEEK(dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA) } AS SEMANA, 'NC' AS VTA_DESCRIPCION_DOCUMENTO, ISNULL(dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID, '') AS VCHA_ART_ARTICULO_ID, ISNULL(dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, '') AS VCHA_ART_NOMBRE_ESPAÑOL, ISNULL(dbo.TB_CATALOGOS.VCHA_CAT_CATALOGO_ID, 'SIN CATALOGO') AS VCHA_CAT_CATALOGO_ID, "
            var_cadena = var_cadena + " ISNULL(dbo.TB_CATALOGOS.VCHA_CAT_NOMBRE, ' SIN CATALOGO ') AS VCHA_CAT_NOMBRE, ISNULL(dbo.TB_DISEÑOS.VCHA_DIS_DISEÑO_ID, 'SIN DISEÑO') AS VCHA_DIS_DISEÑO_ID, ISNULL(dbo.TB_DISEÑOS.VCHA_DIS_NOMBRE, ' SIN DISEÑO ') AS VCHA_DIS_NOMBRE, ISNULL(dbo.TB_LINEAS.VCHA_LIN_LINEA_ID, 'SIN LINEA') AS VCHA_LIN_LINEA_ID, ISNULL(dbo.TB_LINEAS.VCHA_LIN_NOMBRE, 'SIN LINEA') AS VCHA_LIN_NOMBRE, ISNULL(dbo.TB_TALLAS.VCHA_TAL_TALLA_ID, 'SIN TALLA') AS VCHA_TAL_TALLA_ID, ISNULL(dbo.TB_TALLAS.VCHA_TAL_NOMBRE, 'SIN TALLA') AS VCHA_TAL_NOMBRE, ISNULL(dbo.TB_LICENCIAS.VCHA_LIC_LICENCIA_ID, 'SIN LICENCIA') AS VCHA_LIC_LICENCIA_ID, ISNULL(dbo.TB_LICENCIAS.VCHA_LIC_NOMBRE, 'SIN LICENCIA') AS VCHA_LIC_NOMBRE, ISNULL(dbo.TB_ARTICULOS.VCHA_ART_NUMERO_LIC, 'SIN NUMERO DE LICENCIA') AS VCHA_ART_NUMERO_LIC, ISNULL(dbo.TB_ARTICULOS.MONE_ART_PRECIO_BASE, 0) AS MONE_ART_PRECIO_BASE, ISNULL(dbo.TB_ARTICULOS.DTIM_ART_FECHA_ALTA, "
            var_cadena = var_cadena + " GETDATE()) AS DTIM_ART_FECHA_ALTA, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO AS inte_car_numero, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_ZON_DESCRIPCION, 'SIN ZONA') AS VCHA_ZON_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, ISNULL(dbo.VW_CLIENTES.VCHA_CAN_CANAL_VENTA_ID, 'SIN CANAL') AS VCHA_CAN_CANAL_VENTA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_CAN_NOMBRE, 'SIN CANAL') AS VCHA_CAN_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_AGE_AGENTE_ID, 'SIN AGENTE') AS VCHA_AGE_AGENTE_ID, ISNULL(dbo.VW_CLIENTES.VCHA_AGE_NOMBRE, 'SIN AGENTE') AS VCHA_AGE_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_RUT_RUTA_ID, 'SIN RUTA') AS VCHA_RUT_RUTA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_RUT_NOMBRE, 'SIN RUTA') AS VCHA_RUT_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_ZON_ZONA_ID, 'SIN ZONA') AS VCHA_ZON_ZONA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_ZON_DESCRIPCION, 'SIN ZONA') AS VCHA_ZON_DESCRIPCION, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID, 'SIN CLIENTE') AS VCHA_CLI_CLAVE_ID,"
            var_cadena = var_cadena + " REPLACE(REPLACE(ISNULL(dbo.VW_CLIENTES.VCHA_CLI_NOMBRE, 'SIN CLIENTE'), '''', '´'), '´', '') AS VCHA_CLI_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_CLAVE_UNIFICADA_ID, '0') AS VCHA_CLI_CLAVE_UNIFICADA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_RFC, 'SIN RFC') AS VCHA_CLI_RFC, ISNULL(dbo.VW_CLIENTES.VCHA_TIT_TITULAR_ID, 'SIN TITULAR') AS VCHA_TIT_TITULAR_ID, REPLACE(REPLACE(ISNULL(dbo.VW_CLIENTES.VCHA_TIT_NOMBRE, 'SIN TITULAR'), '''', '´'), '´', '') AS VCHA_TIT_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_GAC_GRUPO_ACTUAL_ID, 'SIN GRUPO ') AS VCHA_GAC_GRUPO_ACTUAL_ID, REPLACE(REPLACE(ISNULL(dbo.VW_CLIENTES.VCHA_GAC_NOMBRE, 'SIN GRUPO'), '''', '´'), '´', '') AS VCHA_GAC_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_CP, 'SIN CP') AS VCHA_CLI_CP, ISNULL(dbo.VW_CLIENTES.VCHA_EST_ESTADO_ID, 'SIN ESTADO') AS VCHA_EST_ESTADO_ID, ISNULL(dbo.VW_CLIENTES.VCHA_EST_NOMBRE, 'SIN ESTADO') AS VCHA_EST_NOMBRE, inte_ent_consecutivo_tabla, "
            var_cadena = var_cadena + " ISNULL(dbo.VW_CLIENTES.VCHA_CIU_CIUDAD_ID, 'SIN CIUDAD') AS VCHA_CIU_CIUDAD_ID, ISNULL(dbo.VW_CLIENTES.VCHA_CIU_NOMBRE, 'SIN CIUDAD') AS VCHA_CIU_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_MUN_MUNICIPIO_ID, 'SIN MUNICIPIO') AS VCHA_MUN_MUNICIPIO_ID, ISNULL(dbo.VW_CLIENTES.VCHA_MUN_NOMBRE, 'SIN MUNICIPIO') AS VCHA_MUN_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_COL_COLONIA_ID, 'SIN COLONIA') AS VCHA_COL_COLONIA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_COL_NOMBRE, 'SIN COLONIA') AS VCHA_COL_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_MON_MONEDA_ID, 'SIN MONEDA') AS VCHA_MON_MONEDA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_MON_DIVISA, 'SIN MONEDA') AS VCHA_MON_DIVISA, ISNULL(dbo.VW_CLIENTES.VCHA_ZON_DESCRIPCION, 'SIN ZONA') AS Expr1, ISNULL(dbo.VW_CLIENTES.VCHA_PAI_PAIS_ID, '') AS VCHA_PAI_PAIS_ID, ISNULL(dbo.VW_CLIENTES.VCHA_PAI_NOMBRE, '') AS VCHA_PAI_NOMBRE, dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, dbo.TB_ENTRADAS.FLOA_ENT_COSTO, dbo.TB_ENTRADAS.FLOA_ENT_PRECIO "
            var_cadena = var_cadena + " FROM dbo.TB_CATALOGOS RIGHT OUTER JOIN dbo.TB_LINEAS RIGHT OUTER JOIN dbo.TB_ENTRADAS INNER JOIN dbo.TB_UNIDADESORGANIZACIONALES INNER JOIN dbo.TB_EMPRESAS INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID ON dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID ON dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENTRADAS.INTE_ENT_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN "
            var_cadena = var_cadena + " dbo.TB_ALMACENES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID LEFT OUTER JOIN dbo.TB_ESTABLECIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID = dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID LEFT OUTER Join dbo.VW_CLIENTES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_CLI_CLAVE_ID = dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID LEFT OUTER JOIN dbo.TB_LICENCIAS ON dbo.TB_ARTICULOS.VCHA_LIC_LICENCIA_ID = dbo.TB_LICENCIAS.VCHA_LIC_LICENCIA_ID LEFT OUTER JOIN dbo.TB_TALLAS ON dbo.TB_ARTICULOS.VCHA_TAL_TALLA_ID = dbo.TB_TALLAS.VCHA_TAL_TALLA_ID ON dbo.TB_LINEAS.VCHA_LIN_LINEA_ID = dbo.TB_ARTICULOS.VCHA_LIN_LINEA_ID LEFT OUTER JOIN dbo.TB_DISEÑOS ON dbo.TB_ARTICULOS.VCHA_DIS_DISEÑO_ID = dbo.TB_DISEÑOS.VCHA_DIS_DISEÑO_ID ON dbo.TB_CATALOGOS.VCHA_CAT_CATALOGO_ID = dbo.TB_ARTICULOS.VCHA_ART_CATALOGO_VIGENTE WHERE (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'CC_4') AND "
            var_cadena = var_cadena + " (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= CONVERT(DATETIME, '" + VAR_FECHA_INICIO + "', 102)) AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA < CONVERT(DATETIME, '" + VAR_FECHA_FIN + "', 102)) AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN = 2) "
            rs1.Open var_cadena, cnn_cdindustrial, adOpenDynamic, adLockOptimistic
            While Not rs1.EOF
                  var_dia = CStr(Day(CDate(rs1!DTIM_emo_FECHA)))
                  var_mes = CStr(Month(CDate(rs1!DTIM_emo_FECHA)))
                  var_año = CStr(Year(CDate(rs1!DTIM_emo_FECHA)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_factura = var_dia + "/" + var_mes + "/" + var_año
                  var_dia = CStr(Day(CDate(rs1!DTIM_ART_FECHA_ALTA)))
                  var_mes = CStr(Month(CDate(rs1!DTIM_ART_FECHA_ALTA)))
                  var_año = CStr(Year(CDate(rs1!DTIM_ART_FECHA_ALTA)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_alta = var_dia + "/" + var_mes + "/" + var_año
                  var_canal_venta = Trim(rs1!VCHA_CAN_CANAL_VENTA_ID)
                  If var_canal_venta = "" Then
                     var_canal_venta = "SIN CANAL"
                  End If
                  var_nombre_canal = Trim(rs1!VCHA_CAN_NOMBRE)
                  If var_nombre_canal = "" Then
                     var_nombre_canal = "SIN CANAL"
                  End If
                  var_agente = Trim(rs1!VCHA_AGE_AGENTE_ID)
                  If var_agente = "" Then
                     var_agente = "SIN AGENTE"
                  End If
                  var_nombre_agente = Trim(rs1!VCHA_AGE_NOMBRE)
                  If var_nombre_agente = "" Then
                     var_nombre_agente = "SIN AGENTE"
                  End If
                  var_ruta = Trim(rs1!VCHA_RUT_RUTA_ID)
                  If var_ruta = "" Then
                     var_ruta = "SIN RUTA"
                  End If
                  var_nombre_ruta = Trim(rs1!VCHA_RUT_NOMBRE)
                  If var_nombre_ruta = "" Then
                     var_nombre_ruta = "SIN RUTA"
                  End If
                  var_zona = Trim(rs1!VCHA_ZON_ZONA_ID)
                  If var_zona = "" Then
                     var_zona = "SIN ZONA"
                  End If
                  var_nombre_zona = Trim(rs1!VCHA_ZON_NOMBRE)
                  If var_nombre_zona = "" Then
                     var_nombre_zona = "SIN ZONA"
                  End If
                  var_cliente = Trim(rs1!VCHA_CLI_CLAVE_ID)
                  If var_cliente = "" Then
                     var_cliente = "SIN CLIENTE"
                  End If
                  var_cliente_unfo = Trim(VCHA_CLI_CLAVE_UNIFICADA_ID)
                  If var_cliente_unfo = "" Then
                     var_cliente_unfo = "0"
                  End If
                  var_nombre_cliente = Trim(rs1!VCHA_CLI_NOMBRE)
                  If var_nombre_cliente = "" Then
                     var_nombre_cliente = "SIN CLIENTE"
                  End If
                  var_rfc = Trim(rs1!VCHA_CLI_RFC)
                  var_titular = Trim(rs1!VCHA_TIT_TITULAR_ID)
                  If var_titular = "" Then
                     var_titular = "SIN TITULAR"
                  End If
                  var_nombre_titular = Trim(rs1!VCHA_TIT_NOMBRE)
                  If var_nombre_titular = "" Then
                     var_nombre_titular = "SIN TITULAR"
                  End If
                  var_grupo = Trim(rs1!VCHA_GAC_GRUPO_ACTUAL_ID)
                  If var_grupo = "" Then
                     var_grupo = "SIN GRUPO"
                  End If
                  var_nombre_grupo = Trim(rs1!VCHA_GAC_NOMBRE)
                  If var_nombre_grupo = "" Then
                     var_nombre_grupo = "SIN GRUPO"
                  End If
                  var_establecimiento = ""
                  If var_establecimiento = "" Then
                     var_establecimiento = "SIN ESTABLECIMIENTO"
                  End If
                  var_nombre_establecimiento = ""
                  If var_nombre_establecimiento = "" Then
                     var_nombre_establecimiento = "SIN ESTABLECIMIENTO"
                  End If
                  var_cp = Trim(rs1!VCHA_CLI_CP)
                  If var_cp = "" Then
                     var_cp = "SIN CP"
                  End If
                  var_estado = Trim(rs1!VCHA_EST_ESTADO_ID)
                  If var_estado = "" Then
                     var_estado = "SIN ESTADO"
                  End If
                  var_nombre_estado = Trim(rs1!VCHA_EST_NOMBRE)
                  If var_nombre_estado = "" Then
                     var_nombre_estado = "SIN ESTADO"
                  End If
                  var_ciudad = Trim(rs1!VCHA_CIU_CIUDAD_ID)
                  If var_ciudad = "" Then
                     var_ciudad = "SIN CIUDAD"
                  End If
                  var_nombre_ciudad = Trim(rs1!VCHA_CIU_NOMBRE)
                  If var_nombre_ciudad = "" Then
                     var_nombre_ciudad = "SIN CIUDAD"
                  End If
                  var_municipio = Trim(rs1!VCHA_MUN_MUNICIPIO_ID)
                  If var_municipio = "" Then
                     var_municipio = "SIN MUNICIPIO"
                  End If
                  var_nombre_municipio = Trim(rs1!VCHA_MUN_NOMBRE)
                  If var_nombre_municipio = "" Then
                     var_nombre_municipio = "SIN MUNICIPIO"
                  End If
                  var_colonia = Trim(rs1!vcha_col_colonia_id)
                  If var_colonia = "" Then
                     var_colonia = "SIN COLONIA"
                  End If
                  var_nombre_colonia = Trim(rs1!VCHA_COL_NOMBRE)
                  If var_nombre_colonia = "" Then
                     var_nombre_colonia = "SIN COLONIA"
                  End If
                  var_articulo = Trim(rs1!VCHA_ART_ARTICULO_ID)
                  If var_articulo = "" Then
                     var_articulo = "SIN ARTICULO"
                  End If
                  var_nombre_articulo = Trim(rs1!VCHA_aRT_NOMBRE_ESPAÑOL)
                  If var_nombre_articulo = "" Then
                     var_nombre_articulo = "SIN ARTICULO"
                  End If
                  var_catalogo = Trim(rs1!VCHA_CAT_CATALOGO_ID)
                  If var_catalogo = "" Then
                     var_catalogo = "SIN CATALOGO"
                  End If
                  var_nombre_catalogo = Trim(rs1!VCHA_CAT_NOMBRE)
                  If var_nombre_catalogo = "" Then
                     var_nombre_catalogo = "SIN CATALOGO"
                  End If
                  var_diseño = Trim(rs1!VCHA_DIS_DISEÑO_ID)
                  If var_diseño = "" Then
                     var_diseño = "SIN DISEÑO"
                  End If
                  var_nombre_diseño = Trim(rs1!VCHA_DIS_NOMBRE)
                  If var_nombre_diseño = "" Then
                     var_nombre_diseño = "SIN DISEÑO"
                  End If
                  var_linea = Trim(rs1!VCHA_LIN_LINEA_ID)
                  If var_linea = "" Then
                     var_linea = "SIN LINEA"
                  End If
                  var_nombre_linea = Trim(rs1!VCHA_LIN_NOMBRE)
                  If var_nombre_linea = "" Then
                     var_nombre_linea = "SIN LINEA"
                  End If
                  var_talla = Trim(rs1!vcha_Tal_talla_id)
                  If var_talla = "" Then
                     var_talla = "SIN TALLA"
                  End If
                  var_nombre_talla = Trim(rs1!VCHA_TAL_NOMBRE)
                  If var_nombre_talla = "" Then
                     var_nombre_talla = "SIN TALLA"
                  End If
                  var_licencia = Trim(rs1!vcha_lic_licencia_id)
                  If var_licencia = "" Then
                     var_licencia = "SIN LICENCIA"
                  End If
                  var_nombre_licencia = Trim(rs1!vcha_lic_nombre)
                  If var_nombre_licencia = "" Then
                     var_nombre_licencia = "SIN LICENCIA"
                  End If
                  var_numero_licencia = Trim(rs1!VCHA_ART_NUMERO_LIC)
                  If var_numero_licencia = "" Then
                     var_numero_licencia = "SIN NUMERO DE LICENCIA"
                  End If
                  var_pais = Trim(rs1!VCHA_PAI_PAIS_ID)
                  If var_pais = "" Then
                     var_pais = "SIN PAIS"
                  End If
                  var_nombre_pais = Trim(rs1!vcha_pai_nombre)
                  If var_nombre_pais = "" Then
                     var_nombre_pais = "SIN PAIS"
                  End If
        
         
                  var_descuento_1 = 0
                  var_descuento_2 = 0
                  VAR_PRECIO = rs1!FLOA_ent_PRECIO - (rs1!FLOA_ent_PRECIO * (var_descuento_1 / 100))
                  VAR_PRECIO = VAR_PRECIO - (VAR_PRECIO * (var_descuento_2 / 100))
                  VAR_PORCENTAJE_IVA = IIf(IsNull(rs1!FLOA_CAR_PORCENTAJE_IVA), 0, rs1!FLOA_CAR_PORCENTAJE_IVA)
                  var_precio_2 = VAR_PRECIO * ((1 + (VAR_PORCENTAJE_IVA / 100)))
                  var_importe_iva = var_precio_2 - VAR_PRECIO
                  VAR_PRECIO = var_precio_2
                  var_precio_base = IIf(IsNull(rs1!MONE_ART_PRECIO_BASE), 0, rs1!MONE_ART_PRECIO_BASE) * ((1 + (VAR_PORCENTAJE_IVA / 100)))
                         
                  var_cadena = "insert into  VT_TB_VENTAS  (VTA_INDICE_ID, VTA_ORIGEN_ID, VTA_EMPRESA_ID, VTA_EMPRESA, VTA_UNIDAD_ORGANIZACIONAL_ID, VTA_UNIDAD_ORGANIZACIONAL,VTA_TIENDA_ID,VTA_TIENDA,VTA_CANAL_ID, VTA_CANAL, VTA_REGION,VTA_AGENTE_ID,VTA_AGENTE,VTA_RUTA_ID,VTA_RUTA,VTA_ZONA_ID,VTA_ZONA,VTA_CLIENTE_ID,VTA_CLIENTE_ID_UNFO,VTA_CLIENTE,VTA_RFC_CLIENTE,VTA_TITULAR_ID,VTA_TITULAR,VTA_GRUPO_ID,VTA_GRUPO, "
                  var_cadena = var_cadena + " VTA_ESTABLECIMIENTO_ID,VTA_ESTABLECIMIENTO,VTA_CODIGO_POSTAL,VTA_ESTADO_ID,VTA_ESTADO,VTA_CIUDAD_ID,VTA_CIUDAD,VTA_MUNICIPIO_ID,VTA_MUNICIPIO, VTA_COLONIA_ID,VTA_COLONIA,VTA_FECHA, VTA_SEMANA,VTA_ID_ENCABEZADO_SID,VTA_ID_DETALLE_SID,VTA_MOVIMIENTO_ID,VTA_TIPO_MOVIMIENTO_ID, VTA_DOCUMENTO, VTA_DESCRIPCION_DOCUMENTO, VTA_NUMERO_DOCUMENTO,VTA_SERIE,VTA_PLAZO,VTA_ARTICULO_ID,                 VTA_ARTICULO, "
                  var_cadena = var_cadena + "VTA_CATALOGO_ID,VTA_CATALOGO,VTA_DISENIO_ID,VTA_DISENIO, VTA_FAMILIA_ID, VTA_FAMILIA, VTA_LINEA_ID,VTA_LINEA, VTA_TALLA_ID,VTA_TALLA,VTA_LICENCIA_ID,VTA_LICENCIA,VTA_NUMERO_LICENCIA, "
                  var_cadena = var_cadena + "VTA_PRECIO_BASE, VTA_IMPUESTO_PORCIENTO ,VTA_FECHA_ALTA_CODIGO, VTA_CANTIDAD, VTA_COSTO, VTA_PRECIO_LISTA_A, VTA_PRECIO_LISTA_B, VTA_PRECIO_LISTA_C, VTA_DESCUENTO1 ,VTA_DESCUENTO2, VTA_DESCUENTO3, VTA_TIPO_CAMBIO, VTA_MONEDA ,VTA_IMPORTE,VTA_FLETE_ID,VTA_IMPORTE_FLETE,VTA_IMPUESTO_FLETE,VTA_MOVIMIENTO_INDEX,VTA_TOTAL_MOVIMIENTOS,VTA_CLAVE_REGISTRO, VTA_PAIS_ID, VTA_PAIS, VTA_IMPUESTO, precio) VALUES"
                  var_cadena = var_cadena + "(sq_ventas_id.nextval," + CStr(var_origen) + ",'" + Trim(rs1!VCHA_EMP_EMPRESA_ID) + "', '" + Trim(rs1!VCHA_EMP_NOMBRE) + "','" + Trim(rs1!VCHA_UOR_UNIDAD_ID) + "','" + Trim(rs1!VCHA_UOR_NOMBRE) + "','" + Trim(rs1!VCHA_ALM_ALMACEN_ID) + "', '" + Trim(rs1!VCHA_ALM_NOMBRE) + "','" + Trim(var_canal_venta) + "','" + Trim(var_nombre_canal) + "','','" + Trim(var_agente) + "', '" + Trim(var_nombre_agente) + "','" + Trim(var_ruta) + "','" + Trim(var_nombre_ruta) + "', '" + Trim(var_zona) + "', '" + Trim(var_nombre_zona) + "', '" + Trim(var_cliente) + "','" + Trim(VCHA_CLI_CLAVE_UNIFICADA_ID) + "','" + Trim(var_nombre_cliente) + "', '" + Trim(var_rfc) + "','" + Trim(var_titular) + "', '" + Trim(var_nombre_titular) + "', '" + Trim(var_grupo) + "', '" + Trim(var_nombre_grupo) + "',"
                  var_cadena = var_cadena + "'" + Trim(var_establecimiento) + "', '" + Trim(var_nombre_establecimiento) + "', '" + Trim(var_cp) + "','" + Trim(var_estado) + "', '" + Trim(var_nombre_estado) + "', '" + Trim(var_ciudad) + "','" + Trim(var_nombre_ciudad) + "', '" + Trim(var_municipio) + "','" + Trim(var_nombre_municipio) + "','" + Trim(var_colonia) + "', '" + Trim(var_nombre_colonia) + "', TO_DATE('" + var_fecha_factura + "','DD/MM/YYYY')," + CStr(rs1!SEMANA) + ", '" + CStr(rs1!inte_ent_consecutivo_tabla) + Trim("sqlquezada2") + Trim("sicantia") + "', '" + CStr(rs1!inte_ent_consecutivo_tabla) + Trim("sqlquezada2") + Trim("sidcantia") + "', 'NC',  'NC',          'NC',         'NC'                 , " + CStr(rs1!inte_Car_numero) + ",'TC',0,'" + Trim(rs1!VCHA_ART_ARTICULO_ID) + "', '" + Trim(rs1!VCHA_aRT_NOMBRE_ESPAÑOL) + "',"
                  var_cadena = var_cadena + "'" + Trim(var_catalogo) + "','" + Trim(var_nombre_catalogo) + "', '" + Trim(var_diseño) + "', '" + Trim(var_nombre_diseño) + "','','','" + Trim(var_linea) + "','" + Trim(var_nombre_linea) + "','" + Trim(var_talla) + "','" + Trim(var_nombre_talla) + "','" + Trim(var_licencia) + "', '" + Trim(var_nombre_licencia) + "','" + Trim(var_numero_licencia) + "',"
                  
                  var_cadena = var_cadena + CStr(var_precio_base) + "," + CStr(rs1!FLOA_CAR_PORCENTAJE_IVA) + ",TO_DATE('" + var_fecha_alta + "','DD/MM/YYYY')," + CStr(rs1!floa_ent_cantidad) + "," + CStr(IIf(IsNull(rs1!FLOA_ent_COSTO), 0, rs1!FLOA_ent_COSTO)) + "," + CStr(VAR_PRECIO) + ",0,0," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(1) + ",'" + Trim(rs1!vcha_mon_divisa) + "'," + CStr(rs1!floa_ent_cantidad * VAR_PRECIO) + ",0,0,0,'','','" + Trim(rs1!VCHA_EMP_EMPRESA_ID) + Trim(rs1!VCHA_UOR_UNIDAD_ID) + "NC" + Trim("TC") + CStr(rs1!inte_Car_numero) + "_" + CStr(rs1!inte_ent_consecutivo_tabla) + "','" + Trim(var_pais) + "','" + Trim(var_nombre_pais) + "', " + CStr(rs1!floa_ent_cantidad * var_importe_iva) + "," + CStr(rs1!FLOA_ent_PRECIO) + ")"
                  
            
                  rs2.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
            
                  rs3.Open "SELECT * FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS WHERE  VCHA_EMP_EMPRESA_ID = '" + rs1!VCHA_EMP_EMPRESA_ID + "' AND VCHA_CAR_TIPO_DOCUMENTO = 'NC' AND VCHA_CAR_DOCUMENTO = 'NC' AND VCHA_SER_SERIE_ID = 'TC' AND VCHA_CLI_CLAVE_ID = '" + rs1!VCHA_CLI_CLAVE_ID + "' AND INTE_CAR_NUMERO = " + CStr(rs1!inte_Car_numero), cnn_cdindustrial, adOpenDynamic, adLockOptimistic
                  If rs3.EOF Then
                     rs4.Open "INSERT INTO TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS (VCHA_EMP_EMPRESA_ID,  VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO) VALUES ('" + rs1!VCHA_EMP_EMPRESA_ID + "','NC','NC','TC','" + rs1!VCHA_CLI_CLAVE_ID + "'," + CStr(rs1!inte_Car_numero) + ")", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
                  End If
                  rs3.Close
               
                  rs1.MoveNext
                  var_i = var_i + 1
                  Text1 = var_i
                  Me.Refresh
                  Me.Text1.Refresh
                  'Me.lbl_accion.Refresh
            Wend
            rs1.Close
      
            rs2.Open "select distinct VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
            var_i = 0
            While Not rs2.EOF
                  rs4.Open "UPDATE TB_ENCABEZADO_MOVIMIENTOS SET INTE_EMO_NUMERO_ORIGEN = 1 WHeRE vcha_emp_empresa_id = '" + rs2!VCHA_EMP_EMPRESA_ID + "' and vcha_mov_movimiento_id = 'CC_4' and inte_emo_numero = " + CStr(rs2!inte_Car_numero), cnn_cdindustrial, adOpenDynamic, adLockOptimistic
                  rs2.MoveNext
                  var_i = var_i + 1
                  Text1.Text = var_id
                  Me.Refresh
                  Me.Text1.Refresh
                  Me.lbl_accion.Refresh
            Wend
            rs2.Close
         End If
      End If ' xz
End Sub
Private Sub Command1_Click()
   Dim var_canal_venta As String
   Dim var_nombre_canal As String
   Dim var_agente As String
   Dim var_nombre_agente As String
   Dim var_ruta As String
   Dim var_nombre_ruta As String
   Dim var_zona As String
   Dim var_nombre_zona As String
   Dim var_cliente As String
   Dim var_cliente_unfo As String
   Dim var_nombre_cliente As String
   Dim var_rfc As String
   Dim var_titular As String
   Dim var_nombre_titular As String
   Dim var_grupo As String
   Dim var_nombre_grupo As String
   Dim var_establecimiento As String
   Dim var_nombre_establecimiento As String
   Dim var_cp As String
   Dim var_estado As String
   Dim var_nombre_estado As String
   Dim var_ciudad As String
   Dim var_nombre_ciudad As String
   Dim var_municipio As String
   Dim var_nombre_municipio As String
   Dim var_colonia As String
   Dim var_nombre_colonia As String
   Dim var_articulo As String
   Dim var_nombre_articulo As String
   Dim var_catalogo As String
   Dim var_nombre_catalogo As String
   Dim var_diseño As String
   Dim var_nombre_diseño As String
   Dim var_linea As String
   Dim var_nombre_linea As String
   Dim var_talla As String
   Dim var_nombre_talla As String
   Dim var_licencia As String
   Dim var_nombre_licencia As String
   Dim var_numero_licencia As String
   Dim var_pais As String
   Dim var_nombre_pais As String
   
   
   
   If Me.opt_cdindustrial.Value = True Then
      cnn_cdindustrial.Open "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=SID;Data Source=ADMCDINDUSTRIAL"
      var_origen = 2
   End If
   If Me.opt_distribucion.Value = True Then
      cnn_cdindustrial.Open "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=vianney;Data Source=DISTRIBUCION"
      var_origen = 3
   End If
   If Me.opt_vergel.Value = True Then
      cnn_cdindustrial.Open "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=SIDtextilera;Data Source=SQLQUEZADA2"
      var_origen = 4
   End If
   If Me.opt_tienda_cantia.Value = True Then
      cnn_cdindustrial.Open "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=SIDcantia;Data Source=SQLQUEZADA2"
      var_origen = 5
   End If
   
End Sub

Private Sub cmd_subir_informacion_Click()
   Dim var_cadena As String
   Dim var_canal_venta As String
   Dim var_nombre_canal As String
   Dim var_agente As String
   Dim var_nombre_agente As String
   Dim var_ruta As String
   Dim var_nombre_ruta As String
   Dim var_zona As String
   Dim var_nombre_zona As String
   Dim var_cliente As String
   Dim var_cliente_unfo As String
   Dim var_nombre_cliente As String
   Dim var_rfc As String
   Dim var_titular As String
   Dim var_nombre_titular As String
   Dim var_grupo As String
   Dim var_nombre_grupo As String
   Dim var_establecimiento As String
   Dim var_nombre_establecimiento As String
   Dim var_cp As String
   Dim var_estado As String
   Dim var_nombre_estado As String
   Dim var_ciudad As String
   Dim var_nombre_ciudad As String
   Dim var_municipio As String
   Dim var_nombre_municipio As String
   Dim var_colonia As String
   Dim var_nombre_colonia As String
   Dim var_articulo As String
   Dim var_nombre_articulo As String
   Dim var_catalogo As String
   Dim var_nombre_catalogo As String
   Dim var_diseño As String
   Dim var_nombre_diseño As String
   Dim var_linea As String
   Dim var_nombre_linea As String
   Dim var_talla As String
   Dim var_nombre_talla As String
   Dim var_licencia As String
   Dim var_nombre_licencia As String
   Dim var_numero_licencia As String
   Dim var_pais As String
   Dim var_nombre_pais As String
   
   Dim var_origen As Integer
   
   Set cnn_cdindustrial = CreateObject("ADODB.connection")
   Set cnn_distribucion = CreateObject("ADODB.connection")
   Set cnn_recuperacion = CreateObject("ADODB.connection")
   Set cnn_cantia = CreateObject("ADODB.connection")
   Set cnnoracle = CreateObject("ADODB.connection")
   Set rs1 = CreateObject("ADODB.recordset")
   Set rs2 = CreateObject("ADODB.recordset")
   Set rs3 = CreateObject("ADODB.recordset")
   Set rs4 = CreateObject("ADODB.recordset")
   Set rs5 = CreateObject("ADODB.recordset")
   Set rs6 = CreateObject("ADODB.recordset")
   Set rs7 = CreateObject("ADODB.recordset")
   
   For var_j = 1 To 4
       'MsgBox var_j
       If var_j = 1 Then
          Me.opt_cdindustrial.Value = True
       End If
       If var_j = 2 Then
          Me.opt_distribucion.Value = True
       End If
       If var_j = 3 Then
          Me.opt_vergel.Value = True
       End If
       If var_j = 4 Then
          Me.opt_tienda_cantia.Value = True
       End If
       If Me.opt_cdindustrial.Value = True Then
           If cnn_cdindustrial.State = 1 Then
              cnn_cdindustrial.Close
           End If
           cnn_cdindustrial.Open "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=SID;Data Source=ADMCDINDUSTRIAL"
           var_origen = 2
       End If
       If Me.opt_distribucion.Value = True Then
           If cnn_cdindustrial.State = 1 Then
              cnn_cdindustrial.Close
           End If
          cnn_cdindustrial.Open "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=vianney;Data Source=DISTRIBUCION"
          var_origen = 3
       End If
       If Me.opt_vergel.Value = True Then
           If cnn_cdindustrial.State = 1 Then
              cnn_cdindustrial.Close
           End If
           cnn_cdindustrial.Open "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=SIDtextilera;Data Source=SQLQUEZADA2"
           var_origen = 4
       End If
       If Me.opt_tienda_cantia.Value = True Then
           If cnn_cdindustrial.State = 1 Then
              cnn_cdindustrial.Close
           End If
          cnn_cdindustrial.Open "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=SIDcantia;Data Source=SQLQUEZADA2"
          var_origen = 5
       End If
       If cnnoracle.State = 1 Then
          cnnoracle.Close
       End If
       cnnoracle.Open "Provider=OraOLEDB.Oracle.1;User ID=distribucion;Data Source=AP;Extended Properties=;Persist Security Info=True;Password=distribucion"

       cnn_cdindustrial.CommandTimeout = 360
   
       
       
       
      If var_origen = 2 Then
         rs4.Open "SELECT VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO from tb_encabezado_cartera WHERE inte_int_interface = 0 and vcha_car_documento = 'FA' and (VCHA_ORC_CP = 'SRVDISENO' or VCHA_ORC_CP = 'SQLQUEZADA2') AND (VCHA_ORC_CP_ENTREGA = 'SIDTEXTILERA')  and dtim_car_fecha >= {d '2010-01-01'}", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
      End If
      If var_origen = 4 Then
         rs4.Open "SELECT VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO from tb_encabezado_cartera WHERE inte_int_interface = 0 and vcha_car_documento = 'FA' and (VCHA_ORC_CP = 'ADMCDINDUSTRIAL') AND (VCHA_ORC_CP_ENTREGA = 'SID') and dtim_car_fecha >= {d '2010-01-01'}", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
      End If
      If var_origen = 3 Then
         rs4.Open "SELECT VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO from tb_encabezado_cartera WHERE inte_int_interface = 0 and vcha_car_documento = 'FA' and (VCHA_ORC_CP = 'DISTRIBUCION') AND (VCHA_ORC_CP_ENTREGA = 'VIANNEY') and dtim_car_fecha >= {d '2010-01-01'}", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
      End If
      If var_origen = 5 Then
         rs4.Open "SELECT VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO from tb_encabezado_cartera WHERE inte_int_interface = 0 and vcha_car_documento = 'FA' and (VCHA_ORC_CP = 'SRVDISENO' or VCHA_ORC_CP = 'SQLQUEZADA2') AND (VCHA_ORC_CP_ENTREGA = 'SIDCANTIA') and dtim_car_fecha >= {d '2010-01-01'}", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
      End If
      
      While Not rs4.EOF
            'MsgBox var_origen
            rs1.Open "DELETE FROM vt_tb_ventas where VTA_ORIGEN_ID = " + CStr(var_origen) + " and VTA_EMPRESA_ID = '" + rs4!VCHA_EMP_EMPRESA_ID + "' And vta_documento = 'FA' and vta_serie = '" + rs4!VCHA_SER_SERIE_ID + "' and vta_numero_documento = " + CStr(rs4!inte_Car_numero) + " and vta_cliente_id = '" + rs4!VCHA_CLI_CLAVE_ID + "'", cnnoracle, adOpenDynamic, adLockOptimistic
            rs4.MoveNext
      Wend
      rs4.Close
       
       
       
       
       
       
       
       'rs1.Open "SELECT  COUNT(*) FROM VW_SUBIR_INFORMACION_VENTAS_DETALLE WHERE (INTERFACE_DETALLE = 0 or interface_encabezado = 0) ", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
       'If Not rs1.EOF Then
       '   MsgBox "Total de registros " + CStr(rs1(0).Value), vbOKOnly, "ATENCION"
       'End If
       'rs1.Close
       var_i = 0
       rs1.Open "DELETE FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
   
       'lbl_accion = "Subiendo información"
       'Me.lbl_accion.Refresh
       'MsgBox "aqui empieza el query"
       'rs1.Open "SELECT * FROM VW_SUBIR_INFORMACION_VENTAS_DETALLE WHERE (interface_encabezado = 0) order by inte_sal_consecutivo_tabla", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
       var_dia = CStr(Day(CDate(Me.txt_fecha_inicio)))
       var_mes = CStr(Month(CDate(Me.txt_fecha_inicio)))
       var_año = CStr(Year(CDate(Me.txt_fecha_inicio)))
       If Len(Trim(var_dia)) = 1 Then
          var_dia = "0" + var_dia
       End If
       If Len(Trim(var_mes)) = 1 Then
          var_mes = "0" + var_mes
       End If
       VAR_FECHA_INICIO = var_año + "-" + var_mes + "-" + var_dia
       
       var_dia = CStr(Day(CDate(Me.txt_fecha_fin)))
       var_mes = CStr(Month(CDate(Me.txt_fecha_fin)))
       var_año = CStr(Year(CDate(Me.txt_fecha_fin)))
       If Len(Trim(var_dia)) = 1 Then
          var_dia = "0" + var_dia
       End If
       If Len(Trim(var_mes)) = 1 Then
          var_mes = "0" + var_mes
       End If
       VAR_FECHA_FIN = var_año + "-" + var_mes + "-" + var_dia
   
   
      var_cadena = "SELECT     ISNULL(dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID, 'SIN EMPRESA') AS VCHA_EMP_EMPRESA_ID, ISNULL(dbo.TB_EMPRESAS.VCHA_EMP_NOMBRE, 'SIN EMPRESA') AS VCHA_EMP_NOMBRE, ISNULL(dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID, 'SIN UNIDAD') AS VCHA_UOR_UNIDAD_ID, ISNULL(dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_NOMBRE, 'SIN UNIDAD') AS VCHA_UOR_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_CAN_CANAL_VENTA_ID, 'SIN CANAL') AS VCHA_CAN_CANAL_VENTA_ID,  ISNULL(dbo.VW_CLIENTES.VCHA_CAN_NOMBRE, 'SIN CANAL') AS VCHA_CAN_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_AGE_AGENTE_ID, 'SIN AGENTE') AS VCHA_AGE_AGENTE_ID, ISNULL(dbo.VW_CLIENTES.VCHA_AGE_NOMBRE, 'SIN AGENTE') AS VCHA_AGE_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_RUT_RUTA_ID, 'SIN RUTA') AS VCHA_RUT_RUTA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_RUT_NOMBRE, 'SIN RUTA') AS VCHA_RUT_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_ZON_ZONA_ID, 'SIN ZONA') AS VCHA_ZON_ZONA_ID, "
      var_cadena = var_cadena + " ISNULL(dbo.VW_CLIENTES.VCHA_ZON_DESCRIPCION, 'SIN ZONA') AS VCHA_ZON_DESCRIPCION, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID, 'SIN CLIENTE') AS VCHA_CLI_CLAVE_ID, replace(replace(ISNULL(dbo.VW_CLIENTES.VCHA_CLI_NOMBRE, 'SIN CLIENTE'),'''','´'),'´','') AS VCHA_CLI_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_CLAVE_UNIFICADA_ID, '0') AS VCHA_CLI_CLAVE_UNIFICADA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_RFC, 'SIN RFC') AS VCHA_CLI_RFC, ISNULL(dbo.VW_CLIENTES.VCHA_TIT_TITULAR_ID, 'SIN TITULAR') AS VCHA_TIT_TITULAR_ID, "
      var_cadena = var_cadena + " replace(replace(ISNULL(dbo.VW_CLIENTES.VCHA_TIT_NOMBRE, 'SIN TITULAR'),'''','´'),'´','') AS VCHA_TIT_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_GAC_GRUPO_ACTUAL_ID, 'SIN GRUPO ') AS VCHA_GAC_GRUPO_ACTUAL_ID, replace(replace(ISNULL(dbo.VW_CLIENTES.VCHA_GAC_NOMBRE, 'SIN GRUPO'),'''','´'),'´','')  AS VCHA_GAC_NOMBRE, replace(replace(ISNULL(dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_NOMBRE, 'SIN ESTABLECIMIENTO'),'''','´'),'´','') AS VCHA_ESB_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_CP, 'SIN CP') AS VCHA_CLI_CP, ISNULL(dbo.VW_CLIENTES.VCHA_EST_ESTADO_ID, 'SIN ESTADO') "
      var_cadena = var_cadena + " AS VCHA_EST_ESTADO_ID, ISNULL(dbo.VW_CLIENTES.VCHA_EST_NOMBRE, 'SIN ESTADO') AS VCHA_EST_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_CIU_CIUDAD_ID, 'SIN CIUDAD') AS VCHA_CIU_CIUDAD_ID, ISNULL(dbo.VW_CLIENTES.VCHA_CIU_NOMBRE, 'SIN CIUDAD') AS VCHA_CIU_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_MUN_MUNICIPIO_ID, 'SIN MUNICIPIO') AS VCHA_MUN_MUNICIPIO_ID, ISNULL(dbo.VW_CLIENTES.VCHA_MUN_NOMBRE, 'SIN MUNICIPIO') AS VCHA_MUN_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_COL_COLONIA_ID, 'SIN COLONIA') AS VCHA_COL_COLONIA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_COL_NOMBRE, 'SIN COLONIA') AS VCHA_COL_NOMBRE, ISNULL(dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, GETDATE()) AS DTIM_CAR_FECHA, { fn WEEK(dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA) } AS SEMANA, ISNULL(dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID, 'SIN MOVIMIENTO') AS VCHA_MOV_MOVIMIENTO_ID, ISNULL(dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_TIPO_DOCUMENTO, "
      var_cadena = var_cadena + "'SIN TIPO DOCUMENTO') AS VCHA_CAR_TIPO_DOCUMENTO, ISNULL(dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO, 'SIN DOCUMENTO') AS VCHA_CAR_DOCUMENTO, 'FACTURA' AS VTA_DESCRIPCION_DOCUMENTO, ISNULL(dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO, 0) AS INTE_CAR_NUMERO, ISNULL(dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID, '') AS VCHA_SER_SERIE_ID, ISNULL(dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, 0) AS INTE_CAR_PLAZO, ISNULL(dbo.tb_ARTICULOS.VCHA_ART_ARTICULO_ID, '') AS VCHA_ART_ARTICULO_ID, ISNULL(dbo.tb_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, '') AS VCHA_ART_NOMBRE_ESPAÑOL, ISNULL(dbo.TB_CATALOGOS.VCHA_CAT_CATALOGO_ID, 'SIN CATALOGO') AS VCHA_CAT_CATALOGO_ID, ISNULL(dbo.TB_CATALOGOS.VCHA_CAT_NOMBRE, ' SIN CATALOGO ') AS VCHA_CAT_NOMBRE, ISNULL(dbo.TB_DISEÑOS.VCHA_DIS_DISEÑO_ID, 'SIN DISEÑO') AS VCHA_DIS_DISEÑO_ID, ISNULL(dbo.TB_DISEÑOS.VCHA_DIS_NOMBRE, ' SIN DISEÑO ') AS VCHA_DIS_NOMBRE, "
      var_cadena = var_cadena + " ISNULL(dbo.TB_LINEAS.VCHA_LIN_LINEA_ID, 'SIN LINEA') AS VCHA_LIN_LINEA_ID, ISNULL(dbo.TB_LINEAS.VCHA_LIN_NOMBRE, 'SIN LINEA') AS VCHA_LIN_NOMBRE, ISNULL(dbo.TB_TALLAS.VCHA_TAL_TALLA_ID, 'SIN TALLA') AS VCHA_TAL_TALLA_ID, ISNULL(dbo.TB_TALLAS.VCHA_TAL_NOMBRE, 'SIN TALLA') AS VCHA_TAL_NOMBRE, ISNULL(dbo.TB_LICENCIAS.VCHA_LIC_LICENCIA_ID, 'SIN LICENCIA') AS VCHA_LIC_LICENCIA_ID, ISNULL(dbo.TB_LICENCIAS.VCHA_LIC_NOMBRE, 'SIN LICENCIA') AS VCHA_LIC_NOMBRE, ISNULL(dbo.tb_ARTICULOS.VCHA_ART_NUMERO_LIC, 'SIN NUMERO DE LICENCIA') AS VCHA_ART_NUMERO_LIC, ISNULL(dbo.tb_ARTICULOS.MONE_ART_PRECIO_BASE, 0) AS MONE_ART_PRECIO_BASE, ISNULL(dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_PORCENTAJE_IVA, 0) AS FLOA_CAR_PORCENTAJE_IVA, ISNULL(dbo.tb_ARTICULOS.DTIM_ART_FECHA_ALTA, GETDATE()) AS DTIM_ART_FECHA_ALTA, ISNULL(dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, 0) AS FLOA_SAL_CANTIDAD, dbo.TB_SALIDAS.FLOA_SAL_COSTO, "
      var_cadena = var_cadena + " dbo.TB_SALIDAS.FLOA_SAL_PRECIO AS FLOA_SAL_PRECIO, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2, dbo.TB_SALIDAS.FLOA_SAL_PROMOCION_1, ISNULL(dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO, 1) AS FLOA_CAR_TIPO_CAMBIO, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO, ISNULL(dbo.VW_CLIENTES.VCHA_MON_MONEDA_ID, 'SIN MONEDA') AS VCHA_MON_MONEDA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_MON_DIVISA, 'SIN MONEDA') AS VCHA_MON_DIVISA, ISNULL(dbo.TB_ENCABEZADO_CARTERA.VCHA_ESB_ESTABLECIMIENTO_ID, 'SIN ESTABLECIMIENTO') AS VCHA_ESB_ESTABLECIMIENTO_ID, dbo.TB_ENCABEZADO_CARTERA.ID_DOCUMENTO, dbo.TB_ENCABEZADO_CARTERA.N_SERVIDOR, dbo.TB_ENCABEZADO_CARTERA.N_BASEDATOS, dbo.TB_ENCABEZADO_CARTERA.INTE_INT_INTERFACE AS INTERFACE_ENCABEZADO, dbo.TB_SALIDAS.INTE_INT_INTERFACE AS INTERFACE_DETALLE, dbo.TB_SALIDAS.INTE_SAL_CONSECUTIVO_TABLA, "
      var_cadena = var_cadena + " dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_ZON_DESCRIPCION, 'SIN ZONA') AS VCHA_ZON_NOMBRE, dbo.TB_SALIDAS.N_SERVIDOR AS SERVIDOR_DETALLE, dbo.TB_SALIDAS.N_BASEDATOS AS BASE_DATOS_DETALLE, dbo.TB_SALIDAS.INTE_SAL_AÑO, dbo.TB_SALIDAS.INTE_SAL_NUMERO, ISNULL(dbo.VW_CLIENTES.VCHA_PAI_PAIS_ID, '') AS VCHA_PAI_PAIS_ID, ISNULL(dbo.VW_CLIENTES.VCHA_PAI_NOMBRE, '') AS VCHA_PAI_NOMBRE FROM dbo.TB_CATALOGOS RIGHT OUTER JOIN dbo.TB_DISEÑOS RIGHT OUTER JOIN dbo.TB_ALMACENES INNER JOIN dbo.TB_UNIDADESORGANIZACIONALES INNER JOIN dbo.TB_SALIDAS ON dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID INNER JOIN dbo.tb_ARTICULOS ON dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.tb_ARTICULOS.VCHA_ART_ARTICULO_ID ON dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID INNER JOIN dbo.VW_CLIENTES INNER JOIN dbo.TB_ENCABEZADO_CARTERA INNER JOIN"
      var_cadena = var_cadena + " dbo.TB_EMPRESAS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID ON dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_SALIDAS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO LEFT OUTER JOIN dbo.TB_ESTABLECIMIENTOS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_ESB_ESTABLECIMIENTO_ID = dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID LEFT OUTER JOIN dbo.TB_LICENCIAS ON dbo.tb_ARTICULOS.VCHA_LIC_LICENCIA_ID = dbo.TB_LICENCIAS.VCHA_LIC_LICENCIA_ID LEFT OUTER JOIN dbo.TB_TALLAS ON dbo.tb_ARTICULOS.VCHA_TAL_TALLA_ID = dbo.TB_TALLAS.VCHA_TAL_TALLA_ID LEFT OUTER JOIN "
      var_cadena = var_cadena + " dbo.TB_LINEAS ON dbo.tb_ARTICULOS.VCHA_LIN_LINEA_ID = dbo.TB_LINEAS.VCHA_LIN_LINEA_ID ON dbo.TB_DISEÑOS.VCHA_DIS_DISEÑO_ID = dbo.tb_ARTICULOS.VCHA_DIS_DISEÑO_ID ON dbo.TB_CATALOGOS.VCHA_CAT_CATALOGO_ID = dbo.tb_ARTICULOS.VCHA_ART_CATALOGO_VIGENTE WHERE     (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA >= CONVERT(DATETIME, '" + VAR_FECHA_INICIO + "', 102)) AND (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA < CONVERT(DATETIME, '" + VAR_FECHA_FIN + "', 102)) And dbo.TB_ENCABEZADO_CARTERA.INTE_INT_INTERFACE = 0"
      rs1.Open var_cadena, cnn_cdindustrial, adOpenDynamic, adLockOptimistic
   
   
   
      'MsgBox rs1(0).Value
      'MsgBox "termina el query"
      'MsgBox rs1(0).Value
      While Not rs1.EOF
            var_dia = CStr(Day(CDate(rs1!DTIM_CAR_FECHA)))
            var_mes = CStr(Month(CDate(rs1!DTIM_CAR_FECHA)))
            var_año = CStr(Year(CDate(rs1!DTIM_CAR_FECHA)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_factura = var_dia + "/" + var_mes + "/" + var_año
            
            var_dia = CStr(Day(CDate(rs1!DTIM_ART_FECHA_ALTA)))
            var_mes = CStr(Month(CDate(rs1!DTIM_ART_FECHA_ALTA)))
            var_año = CStr(Year(CDate(rs1!DTIM_ART_FECHA_ALTA)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            
            var_fecha_alta = var_dia + "/" + var_mes + "/" + var_año
            
            
            var_canal_venta = Trim(rs1!VCHA_CAN_CANAL_VENTA_ID)
            If var_canal_venta = "" Then
               var_canal_venta = "SIN CANAL"
            End If
            var_nombre_canal = Trim(rs1!VCHA_CAN_NOMBRE)
            If var_nombre_canal = "" Then
               var_nombre_canal = "SIN CANAL"
            End If
            var_agente = Trim(rs1!VCHA_AGE_AGENTE_ID)
            If var_agente = "" Then
               var_agente = "SIN AGENTE"
            End If
            var_nombre_agente = Trim(rs1!VCHA_AGE_NOMBRE)
            If var_nombre_agente = "" Then
               var_nombre_agente = "SIN AGENTE"
            End If
            var_ruta = Trim(rs1!VCHA_RUT_RUTA_ID)
            If var_ruta = "" Then
               var_ruta = "SIN RUTA"
            End If
            var_nombre_ruta = Trim(rs1!VCHA_RUT_NOMBRE)
            If var_nombre_ruta = "" Then
               var_nombre_ruta = "SIN RUTA"
            End If
            var_zona = Trim(rs1!VCHA_ZON_ZONA_ID)
            If var_zona = "" Then
               var_zona = "SIN ZONA"
            End If
            var_nombre_zona = Trim(rs1!VCHA_ZON_NOMBRE)
            If var_nombre_zona = "" Then
               var_nombre_zona = "SIN ZONA"
            End If
            var_cliente = Trim(rs1!VCHA_CLI_CLAVE_ID)
            If var_cliente = "" Then
               var_cliente = "SIN CLIENTE"
            End If
            var_cliente_unfo = Trim(VCHA_CLI_CLAVE_UNIFICADA_ID)
            If var_cliente_unfo = "" Then
               var_cliente_unfo = "0"
            End If
            var_nombre_cliente = Trim(rs1!VCHA_CLI_NOMBRE)
            If var_nombre_cliente = "" Then
               var_nombre_cliente = "SIN CLIENTE"
            End If
            var_rfc = Trim(rs1!VCHA_CLI_RFC)
            var_titular = Trim(rs1!VCHA_TIT_TITULAR_ID)
            If var_titular = "" Then
               var_titular = "SIN TITULAR"
            End If
            var_nombre_titular = Trim(rs1!VCHA_TIT_NOMBRE)
            If var_nombre_titular = "" Then
               var_nombre_titular = "SIN TITULAR"
            End If
            var_grupo = Trim(rs1!VCHA_GAC_GRUPO_ACTUAL_ID)
            If var_grupo = "" Then
               var_grupo = "SIN GRUPO"
            End If
            var_nombre_grupo = Trim(rs1!VCHA_GAC_NOMBRE)
            If var_nombre_grupo = "" Then
               var_nombre_grupo = "SIN GRUPO"
            End If
            var_establecimiento = Trim(rs1!VCHA_ESB_ESTABLECIMIENTO_ID)
            If var_establecimiento = "" Then
               var_establecimiento = "SIN ESTABLECIMIENTO"
            End If
            var_nombre_establecimiento = Trim(rs1!VCHA_ESB_NOMBRE)
            If var_nombre_establecimiento = "" Then
               var_nombre_establecimiento = "SIN ESTABLECIMIENTO"
            End If
            var_cp = Trim(rs1!VCHA_CLI_CP)
            If var_cp = "" Then
               var_cp = "SIN CP"
            End If
             var_estado = Trim(rs1!VCHA_EST_ESTADO_ID)
            If var_estado = "" Then
               var_estado = "SIN ESTADO"
            End If
            var_nombre_estado = Trim(rs1!VCHA_EST_NOMBRE)
            If var_nombre_estado = "" Then
               var_nombre_estado = "SIN ESTADO"
            End If
            var_ciudad = Trim(rs1!VCHA_CIU_CIUDAD_ID)
            If var_ciudad = "" Then
               var_ciudad = "SIN CIUDAD"
            End If
            var_nombre_ciudad = Trim(rs1!VCHA_CIU_NOMBRE)
            If var_nombre_ciudad = "" Then
               var_nombre_ciudad = "SIN CIUDAD"
            End If
            var_municipio = Trim(rs1!VCHA_MUN_MUNICIPIO_ID)
            If var_municipio = "" Then
               var_municipio = "SIN MUNICIPIO"
            End If
            var_nombre_municipio = Trim(rs1!VCHA_MUN_NOMBRE)
            If var_nombre_municipio = "" Then
               var_nombre_municipio = "SIN MUNICIPIO"
            End If
            var_colonia = Trim(rs1!vcha_col_colonia_id)
            If var_colonia = "" Then
               var_colonia = "SIN COLONIA"
            End If
            var_nombre_colonia = Trim(rs1!VCHA_COL_NOMBRE)
            If var_nombre_colonia = "" Then
               var_nombre_colonia = "SIN COLONIA"
            End If
            var_articulo = Trim(rs1!VCHA_ART_ARTICULO_ID)
            If var_articulo = "" Then
               var_articulo = "SIN ARTICULO"
            End If
            var_nombre_articulo = Trim(rs1!VCHA_aRT_NOMBRE_ESPAÑOL)
            If var_nombre_articulo = "" Then
               var_nombre_articulo = "SIN ARTICULO"
            End If
            var_catalogo = Trim(rs1!VCHA_CAT_CATALOGO_ID)
            If var_catalogo = "" Then
               var_catalogo = "SIN CATALOGO"
            End If
            var_nombre_catalogo = Trim(rs1!VCHA_CAT_NOMBRE)
            If var_nombre_catalogo = "" Then
               var_nombre_catalogo = "SIN CATALOGO"
            End If
            var_diseño = Trim(rs1!VCHA_DIS_DISEÑO_ID)
            If var_diseño = "" Then
               var_diseño = "SIN DISEÑO"
            End If
            var_nombre_diseño = Trim(rs1!VCHA_DIS_NOMBRE)
            If var_nombre_diseño = "" Then
               var_nombre_diseño = "SIN DISEÑO"
            End If
            var_linea = Trim(rs1!VCHA_LIN_LINEA_ID)
            If var_linea = "" Then
               var_linea = "SIN LINEA"
            End If
            var_nombre_linea = Trim(rs1!VCHA_LIN_NOMBRE)
            If var_nombre_linea = "" Then
               var_nombre_linea = "SIN LINEA"
            End If
            var_talla = Trim(rs1!vcha_Tal_talla_id)
            If var_talla = "" Then
               var_talla = "SIN TALLA"
            End If
            var_nombre_talla = Trim(rs1!VCHA_TAL_NOMBRE)
            If var_nombre_talla = "" Then
               var_nombre_talla = "SIN TALLA"
            End If
            var_licencia = Trim(rs1!vcha_lic_licencia_id)
            If var_licencia = "" Then
               var_licencia = "SIN LICENCIA"
            End If
            var_nombre_licencia = Trim(rs1!vcha_lic_nombre)
            If var_nombre_licencia = "" Then
               var_nombre_licencia = "SIN LICENCIA"
            End If
            var_numero_licencia = Trim(rs1!VCHA_ART_NUMERO_LIC)
            If var_numero_licencia = "" Then
               var_numero_licencia = "SIN NUMERO DE LICENCIA"
            End If
            var_pais = Trim(rs1!VCHA_PAI_PAIS_ID)
            If var_pais = "" Then
               var_pais = "SIN PAIS"
            End If
            var_nombre_pais = Trim(rs1!vcha_pai_nombre)
            If var_nombre_pais = "" Then
               var_nombre_pais = "SIN PAIS"
            End If
         
            
            var_descuento_1 = IIf(IsNull(rs1!FLOA_SAL_DESCUENTO_1), 0, rs1!FLOA_SAL_DESCUENTO_1)
            var_descuento_2 = IIf(IsNull(rs1!FLOA_SAL_DESCUENTO_2), 0, rs1!FLOA_SAL_DESCUENTO_2)
            VAR_PRECIO = rs1!FLOA_SAL_PRECIO - (rs1!FLOA_SAL_PRECIO * (var_descuento_1 / 100))
            VAR_PRECIO = VAR_PRECIO - (VAR_PRECIO * (var_descuento_2 / 100))
            VAR_PORCENTAJE_IVA = IIf(IsNull(rs1!FLOA_CAR_PORCENTAJE_IVA), 0, rs1!FLOA_CAR_PORCENTAJE_IVA)
            var_precio_2 = VAR_PRECIO * ((1 + (VAR_PORCENTAJE_IVA / 100)))
            var_importe_iva = var_precio_2 - VAR_PRECIO
            VAR_PRECIO = var_precio_2
            var_precio_base = IIf(IsNull(rs1!MONE_ART_PRECIO_BASE), 0, rs1!MONE_ART_PRECIO_BASE) * ((1 + (VAR_PORCENTAJE_IVA / 100)))
                   
            var_cadena = "insert into  VT_TB_VENTAS  (VTA_INDICE_ID, VTA_ORIGEN_ID, VTA_EMPRESA_ID, VTA_EMPRESA, VTA_UNIDAD_ORGANIZACIONAL_ID, VTA_UNIDAD_ORGANIZACIONAL,VTA_TIENDA_ID,VTA_TIENDA,VTA_CANAL_ID, VTA_CANAL, VTA_REGION,VTA_AGENTE_ID,VTA_AGENTE,VTA_RUTA_ID,VTA_RUTA,VTA_ZONA_ID,VTA_ZONA,VTA_CLIENTE_ID,VTA_CLIENTE_ID_UNFO,VTA_CLIENTE,VTA_RFC_CLIENTE,VTA_TITULAR_ID,VTA_TITULAR,VTA_GRUPO_ID,VTA_GRUPO, "
            var_cadena = var_cadena + " VTA_ESTABLECIMIENTO_ID,VTA_ESTABLECIMIENTO,VTA_CODIGO_POSTAL,VTA_ESTADO_ID,VTA_ESTADO,VTA_CIUDAD_ID,VTA_CIUDAD,VTA_MUNICIPIO_ID,VTA_MUNICIPIO, VTA_COLONIA_ID,VTA_COLONIA,VTA_FECHA, VTA_SEMANA,VTA_ID_ENCABEZADO_SID,VTA_ID_DETALLE_SID,VTA_MOVIMIENTO_ID,VTA_TIPO_MOVIMIENTO_ID, VTA_DOCUMENTO, VTA_DESCRIPCION_DOCUMENTO, VTA_NUMERO_DOCUMENTO,VTA_SERIE,VTA_PLAZO,VTA_ARTICULO_ID,                 VTA_ARTICULO, "
            var_cadena = var_cadena + "VTA_CATALOGO_ID,VTA_CATALOGO,VTA_DISENIO_ID,VTA_DISENIO, VTA_FAMILIA_ID, VTA_FAMILIA, VTA_LINEA_ID,VTA_LINEA, VTA_TALLA_ID,VTA_TALLA,VTA_LICENCIA_ID,VTA_LICENCIA,VTA_NUMERO_LICENCIA, "
            var_cadena = var_cadena + "VTA_PRECIO_BASE, VTA_IMPUESTO_PORCIENTO ,VTA_FECHA_ALTA_CODIGO, VTA_CANTIDAD, VTA_COSTO, VTA_PRECIO_LISTA_A, VTA_PRECIO_LISTA_B, VTA_PRECIO_LISTA_C, VTA_DESCUENTO1 ,VTA_DESCUENTO2, VTA_DESCUENTO3, VTA_TIPO_CAMBIO, VTA_MONEDA ,VTA_IMPORTE,VTA_FLETE_ID,VTA_IMPORTE_FLETE,VTA_IMPUESTO_FLETE,VTA_MOVIMIENTO_INDEX,VTA_TOTAL_MOVIMIENTOS,VTA_CLAVE_REGISTRO, VTA_PAIS_ID, VTA_PAIS, VTA_IMPUESTO, precio) VALUES"
            var_cadena = var_cadena + "(sq_ventas_id.nextval," + CStr(var_origen) + ",'" + Trim(rs1!VCHA_EMP_EMPRESA_ID) + "', '" + Trim(rs1!VCHA_EMP_NOMBRE) + "','" + Trim(rs1!VCHA_UOR_UNIDAD_ID) + "','" + Trim(rs1!VCHA_UOR_NOMBRE) + "','" + Trim(rs1!VCHA_ALM_ALMACEN_ID) + "', '" + Trim(rs1!VCHA_ALM_NOMBRE) + "','" + Trim(var_canal_venta) + "','" + Trim(var_nombre_canal) + "','','" + Trim(var_agente) + "', '" + Trim(var_nombre_agente) + "','" + Trim(var_ruta) + "','" + Trim(var_nombre_ruta) + "', '" + Trim(var_zona) + "', '" + Trim(var_nombre_zona) + "', '" + Trim(var_cliente) + "','" + Trim(VCHA_CLI_CLAVE_UNIFICADA_ID) + "','" + Trim(var_nombre_cliente) + "', '" + Trim(var_rfc) + "','" + Trim(var_titular) + "', '" + Trim(var_nombre_titular) + "', '" + Trim(var_grupo) + "', '" + Trim(Mid(var_nombre_grupo, 1, 50)) + "',"
            var_cadena = var_cadena + "'" + Trim(var_establecimiento) + "', '" + Trim(var_nombre_establecimiento) + "', '" + Trim(var_cp) + "','" + Trim(var_estado) + "', '" + Trim(var_nombre_estado) + "', '" + Trim(var_ciudad) + "','" + Trim(var_nombre_ciudad) + "', '" + Trim(var_municipio) + "','" + Trim(var_nombre_municipio) + "','" + Trim(var_colonia) + "', '" + Trim(var_nombre_colonia) + "', TO_DATE('" + var_fecha_factura + "','DD/MM/YYYY')," + CStr(rs1!SEMANA) + ", '" + CStr(rs1!ID_DOCUMENTO) + Trim(rs1!N_SERVIDOR) + Trim(rs1!n_basedatos) + "', '" + CStr(rs1!inte_sal_consecutivo_tabla) + Trim(rs1!N_SERVIDOR) + Trim(rs1!n_basedatos) + "', '" + Trim(rs1!VCHA_MOV_MOVIMIENTO_ID) + "',  'FA',          'FA',         'FA'                 , " + CStr(rs1!inte_Car_numero) + ",'" + Trim(rs1!VCHA_SER_SERIE_ID) + "'," + CStr(rs1!INTE_CAR_PLAZO) + ",'" + Trim(rs1!VCHA_ART_ARTICULO_ID) + "', '" + Trim(rs1!VCHA_aRT_NOMBRE_ESPAÑOL) + "',"
            var_cadena = var_cadena + "'" + Trim(var_catalogo) + "','" + Trim(var_nombre_catalogo) + "', '" + Trim(var_diseño) + "', '" + Trim(var_nombre_diseño) + "','','','" + Trim(var_linea) + "','" + Trim(var_nombre_linea) + "','" + Trim(var_talla) + "','" + Trim(var_nombre_talla) + "','" + Trim(var_licencia) + "', '" + Trim(var_nombre_licencia) + "','" + Trim(var_numero_licencia) + "',"
            var_cadena = var_cadena + CStr(var_precio_base) + "," + CStr(rs1!FLOA_CAR_PORCENTAJE_IVA) + ",TO_DATE('" + var_fecha_alta + "','DD/MM/YYYY')," + CStr(rs1!floa_sal_cantidad) + "," + CStr(IIf(IsNull(rs1!FLOA_SAL_COSTO), 0, rs1!FLOA_SAL_COSTO)) + "," + CStr(VAR_PRECIO) + ",0,0," + CStr(rs1!FLOA_SAL_DESCUENTO_1) + "," + CStr(rs1!FLOA_SAL_DESCUENTO_2) + "," + CStr(rs1!FLOA_SAL_PROMOCION_1) + "," + CStr(rs1!FLOA_CAR_TIPO_CAMBIO) + ",'" + Trim(rs1!vcha_mon_divisa) + "'," + CStr(rs1!floa_sal_cantidad * VAR_PRECIO) + ",0,0,0,'','','" + Trim(rs1!VCHA_EMP_EMPRESA_ID) + Trim(rs1!VCHA_UOR_UNIDAD_ID) + "FA" + Trim(rs1!VCHA_SER_SERIE_ID) + CStr(rs1!inte_Car_numero) + "_" + CStr(rs1!inte_sal_consecutivo_tabla) + Trim(rs1!n_basedatos) + "','" + Trim(var_pais) + "','" + Trim(var_nombre_pais) + "', " + CStr(rs1!floa_sal_cantidad * var_importe_iva) + "," + CStr(rs1!FLOA_SAL_PRECIO) + ")"
            
            rs2.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
            rs3.Open "SELECT * FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS WHERE  VCHA_EMP_EMPRESA_ID = '" + rs1!VCHA_EMP_EMPRESA_ID + "' AND VCHA_CAR_TIPO_DOCUMENTO = 'FA' AND VCHA_CAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + rs1!VCHA_SER_SERIE_ID + "' AND VCHA_CLI_CLAVE_ID = '" + rs1!VCHA_CLI_CLAVE_ID + "' AND INTE_CAR_NUMERO = " + CStr(rs1!inte_Car_numero), cnn_cdindustrial, adOpenDynamic, adLockOptimistic
            If rs3.EOF Then
               rs4.Open "INSERT INTO TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS (VCHA_EMP_EMPRESA_ID,  VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO) VALUES ('" + rs1!VCHA_EMP_EMPRESA_ID + "','FA','FA','" + rs1!VCHA_SER_SERIE_ID + "','" + rs1!VCHA_CLI_CLAVE_ID + "'," + CStr(rs1!inte_Car_numero) + ")", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
            End If
            rs3.Close
            rs1.MoveNext
            var_i = var_i + 1
            Text1 = var_i
            Me.Refresh
            Me.Text1.Refresh
            'Me.lbl_accion.Refresh
      Wend
      'rs1.Close
      lbl_accion = "Actalizando estatus cartera"
      rs2.Open "select distinct VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
      var_i = 0
      While Not rs2.EOF
            rs3.Open "UPDATE TB_ENCABEZADO_cARTERA SET INTE_INT_INTERFACE = 1 WHERE VCHA_EMP_EMPRESA_ID = '" + rs2!VCHA_EMP_EMPRESA_ID + "' AND VCHA_CAR_TIPO_DOCUMENTO = 'FA' AND VCHA_CAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + rs2!VCHA_SER_SERIE_ID + "' AND VCHA_CLI_CLAVE_ID = '" + rs2!VCHA_CLI_CLAVE_ID + "' AND INTE_CAR_NUMERO = '" + CStr(rs2!inte_Car_numero) + "'", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
            rs2.MoveNext
            var_i = var_i + 1
            Text1.Text = var_id
            Me.Refresh
            Me.Text1.Refresh
            Me.lbl_accion.Refresh
      Wend
      rs2.Close
      'If rs1.Fields.Count > 0 Then
      '   rs1.MoveFirst
      'End If
      'While Not rs1.EOF
      '      rs3.Open "UPDATE TB_sALIDAS SET INTE_INT_INTERFACE = 1 WHERE VCHA_EMP_EMPRESA_ID = '" + rs1!vcha_emp_empresa_id + "' AND VCHA_UOR_UNIDAD_ID = '" + rs1!VCHA_UOR_UNIDAD_ID + "' AND VCHA_ALM_ALMACEN_ID = '" + rs1!VCHA_ALM_ALMACEN_ID + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + rs1!VCHA_MOV_MOVIMIENTO_ID + "' AND INTE_SAL_NUMERO = " + CStr(rs1!INTE_sAL_NUMERO) + " AND VCHA_ART_ARTICULO_ID = '" + rs1!VCHA_ART_aRTICULO_ID + "' AND INTE_SAL_AÑO = " + CStr(rs1!INTE_sAL_AÑO), cnn_cdindustrial, adOpenDynamic, adLockOptimistic
      '      rs1.MoveNext
      'Wend
      rs1.Close
      If rs4.State = 1 Then
         rs4.Close
      End If
      
      If var_origen = 2 Then
         rs4.Open "SELECT VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO from tb_encabezado_cartera WHERE inte_int_interface = 2 and vcha_car_documento = 'FA' and (VCHA_ORC_CP = 'SRVDISENO' or VCHA_ORC_CP = 'SQLQUEZADA2') AND (VCHA_ORC_CP_ENTREGA = 'SIDTEXTILERA')", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
      End If
      If var_origen = 4 Then
         rs4.Open "SELECT VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO from tb_encabezado_cartera WHERE inte_int_interface = 2 and vcha_car_documento = 'FA' and (VCHA_ORC_CP = 'ADMCDINDUSTRIAL') AND (VCHA_ORC_CP_ENTREGA = 'SID')", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
      End If
      If var_origen = 3 Then
         rs4.Open "SELECT VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO from tb_encabezado_cartera WHERE inte_int_interface = 2 and vcha_car_documento = 'FA' and (VCHA_ORC_CP = 'DISTRIBUCION') AND (VCHA_ORC_CP_ENTREGA = 'VIANNEY')", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
      End If
      If var_origen = 5 Then
         rs4.Open "SELECT VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO from tb_encabezado_cartera WHERE inte_int_interface = 2 and vcha_car_documento = 'FA' and (VCHA_ORC_CP = 'SRVDISENO' or VCHA_ORC_CP = 'SQLQUEZADA2') AND (VCHA_ORC_CP_ENTREGA = 'SIDCANTIA')", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
      End If
      
      While Not rs4.EOF
            'MsgBox "DELETE FROM vt_tb_ventas where VTA_ORIGEN_ID = " + CStr(var_origen) + " and VTA_EMPRESA_ID = '" + rs4!VCHA_EMP_EMPRESA_ID + "' And vta_documento = 'FA' and vta_serie = '" + rs4!VCHA_SER_SERIE_ID + "' and vta_numero_documento = " + CStr(rs4!INTE_CAR_NUMERO) + " and vta_cliente_id = '" + rs4!VCHA_CLI_CLAVE_ID + "'"
            rs1.Open "DELETE FROM vt_tb_ventas where VTA_ORIGEN_ID = " + CStr(var_origen) + " and VTA_EMPRESA_ID = '" + rs4!VCHA_EMP_EMPRESA_ID + "' And vta_documento = 'FA' and vta_serie = '" + rs4!VCHA_SER_SERIE_ID + "' and vta_numero_documento = " + CStr(rs4!inte_Car_numero) + " and vta_cliente_id = '" + rs4!VCHA_CLI_CLAVE_ID + "'", cnnoracle, adOpenDynamic, adLockOptimistic
            rs4.MoveNext
      Wend
      If rs4.RecordCount > 0 Then
         rs4.MoveFirst
         While Not rs4.EOF
               rs1.Open "DELETE FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
               rs1.Open "SELECT  * FROM VW_SUBIR_INFORMACION_VENTAS_DETALLE WHERE interface_encabezado = 2 and vcha_emp_empresa_id = '" + rs4!VCHA_EMP_EMPRESA_ID + "' And vcha_Car_documento = 'FA' and vcha_Ser_serie_id= '" + rs4!VCHA_SER_SERIE_ID + "' and inte_car_numero = " + CStr(rs4!inte_Car_numero) + " and vcha_cli_clave_id = '" + rs4!VCHA_CLI_CLAVE_ID + "'", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
               While Not rs1.EOF
                     var_dia = CStr(Day(CDate(rs1!DTIM_CAR_FECHA)))
                     var_mes = CStr(Month(CDate(rs1!DTIM_CAR_FECHA)))
                     var_año = CStr(Year(CDate(rs1!DTIM_CAR_FECHA)))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     var_fecha_factura = var_dia + "/" + var_mes + "/" + var_año
                     
                     var_dia = CStr(Day(CDate(rs1!DTIM_ART_FECHA_ALTA)))
                     var_mes = CStr(Month(CDate(rs1!DTIM_ART_FECHA_ALTA)))
                     var_año = CStr(Year(CDate(rs1!DTIM_ART_FECHA_ALTA)))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     
                     var_fecha_alta = var_dia + "/" + var_mes + "/" + var_año
                     
                     
                     
                     var_canal_venta = Trim(rs1!VCHA_CAN_CANAL_VENTA_ID)
                     If var_canal_venta = "" Then
                        var_canal_venta = "SIN CANAL"
                     End If
                     var_nombre_canal = Trim(rs1!VCHA_CAN_NOMBRE)
                     If var_nombre_canal = "" Then
                        var_nombre_canal = "SIN CANAL"
                     End If
                     var_agente = Trim(rs1!VCHA_AGE_AGENTE_ID)
                     If var_agente = "" Then
                        var_agente = "SIN AGENTE"
                     End If
                     var_nombre_agente = Trim(rs1!VCHA_AGE_NOMBRE)
                     If var_nombre_agente = "" Then
                        var_nombre_agente = "SIN AGENTE"
                     End If
                     var_ruta = Trim(rs1!VCHA_RUT_RUTA_ID)
                     If var_ruta = "" Then
                        var_ruta = "SIN RUTA"
                     End If
                     var_nombre_ruta = Trim(rs1!VCHA_RUT_NOMBRE)
                     If var_nombre_ruta = "" Then
                        var_nombre_ruta = "SIN RUTA"
                     End If
                     var_zona = Trim(rs1!VCHA_ZON_ZONA_ID)
                     If var_zona = "" Then
                        var_zona = "SIN ZONA"
                     End If
                     var_nombre_zona = Trim(rs1!VCHA_ZON_NOMBRE)
                     If var_nombre_zona = "" Then
                        var_nombre_zona = "SIN ZONA"
                     End If
                     var_cliente = Trim(rs1!VCHA_CLI_CLAVE_ID)
                     If var_cliente = "" Then
                        var_cliente = "SIN CLIENTE"
                     End If
                     var_cliente_unfo = Trim(VCHA_CLI_CLAVE_UNIFICADA_ID)
                     If var_cliente_unfo = "" Then
                        var_cliente_unfo = "0"
                     End If
                     var_nombre_cliente = Trim(rs1!VCHA_CLI_NOMBRE)
                     If var_nombre_cliente = "" Then
                        var_nombre_cliente = "SIN CLIENTE"
                     End If
                     var_rfc = Trim(rs1!VCHA_CLI_RFC)
                     var_titular = Trim(rs1!VCHA_TIT_TITULAR_ID)
                     If var_titular = "" Then
                        var_titular = "SIN TITULAR"
                     End If
                     var_nombre_titular = Trim(rs1!VCHA_TIT_NOMBRE)
                     If var_nombre_titular = "" Then
                        var_nombre_titular = "SIN TITULAR"
                     End If
                     var_grupo = Trim(rs1!VCHA_GAC_GRUPO_ACTUAL_ID)
                     If var_grupo = "" Then
                        var_grupo = "SIN GRUPO"
                     End If
                     var_nombre_grupo = Trim(rs1!VCHA_GAC_NOMBRE)
                     If var_nombre_grupo = "" Then
                        var_nombre_grupo = "SIN GRUPO"
                     End If
                     var_establecimiento = Trim(rs1!VCHA_ESB_ESTABLECIMIENTO_ID)
                     If var_establecimiento = "" Then
                        var_establecimiento = "SIN ESTABLECIMIENTO"
                     End If
                     var_nombre_establecimiento = Trim(rs1!VCHA_ESB_NOMBRE)
                     If var_nombre_establecimiento = "" Then
                        var_nombre_establecimiento = "SIN ESTABLECIMIENTO"
                     End If
                     var_cp = Trim(rs1!VCHA_CLI_CP)
                     If var_cp = "" Then
                        var_cp = "SIN CP"
                     End If
                     var_estado = Trim(rs1!VCHA_EST_ESTADO_ID)
                     If var_estado = "" Then
                        var_estado = "SIN ESTADO"
                     End If
                     var_nombre_estado = Trim(rs1!VCHA_EST_NOMBRE)
                     If var_nombre_estado = "" Then
                        var_nombre_estado = "SIN ESTADO"
                     End If
                     var_ciudad = Trim(rs1!VCHA_CIU_CIUDAD_ID)
                     If var_ciudad = "" Then
                        var_ciudad = "SIN CIUDAD"
                     End If
                     var_nombre_ciudad = Trim(rs1!VCHA_CIU_NOMBRE)
                     If var_nombre_ciudad = "" Then
                        var_nombre_ciudad = "SIN CIUDAD"
                     End If
                     var_municipio = Trim(rs1!VCHA_MUN_MUNICIPIO_ID)
                     If var_municipio = "" Then
                        var_municipio = "SIN MUNICIPIO"
                     End If
                     var_nombre_municipio = Trim(rs1!VCHA_MUN_NOMBRE)
                     If var_nombre_municipio = "" Then
                        var_nombre_municipio = "SIN MUNICIPIO"
                     End If
                     var_colonia = Trim(rs1!vcha_col_colonia_id)
                     If var_colonia = "" Then
                        var_colonia = "SIN COLONIA"
                     End If
                     var_nombre_colonia = Trim(rs1!VCHA_COL_NOMBRE)
                     If var_nombre_colonia = "" Then
                        var_nombre_colonia = "SIN COLONIA"
                     End If
                     var_articulo = Trim(rs1!VCHA_ART_ARTICULO_ID)
                     If var_articulo = "" Then
                        var_articulo = "SIN ARTICULO"
                     End If
                     var_nombre_articulo = Trim(rs1!VCHA_aRT_NOMBRE_ESPAÑOL)
                     If var_nombre_articulo = "" Then
                        var_nombre_articulo = "SIN ARTICULO"
                     End If
                     var_catalogo = Trim(rs1!VCHA_CAT_CATALOGO_ID)
                     If var_catalogo = "" Then
                        var_catalogo = "SIN CATALOGO"
                     End If
                     var_nombre_catalogo = Trim(rs1!VCHA_CAT_NOMBRE)
                     If var_nombre_catalogo = "" Then
                        var_nombre_catalogo = "SIN CATALOGO"
                     End If
                     var_diseño = Trim(rs1!VCHA_DIS_DISEÑO_ID)
                     If var_diseño = "" Then
                        var_diseño = "SIN DISEÑO"
                     End If
                     var_nombre_diseño = Trim(rs1!VCHA_DIS_NOMBRE)
                     If var_nombre_diseño = "" Then
                        var_nombre_diseño = "SIN DISEÑO"
                     End If
                     var_linea = Trim(rs1!VCHA_LIN_LINEA_ID)
                     If var_linea = "" Then
                        var_linea = "SIN LINEA"
                     End If
                     var_nombre_linea = Trim(rs1!VCHA_LIN_NOMBRE)
                     If var_nombre_linea = "" Then
                        var_nombre_linea = "SIN LINEA"
                     End If
                     var_talla = Trim(rs1!vcha_Tal_talla_id)
                     If var_talla = "" Then
                        var_talla = "SIN TALLA"
                     End If
                     var_nombre_talla = Trim(rs1!VCHA_TAL_NOMBRE)
                     If var_nombre_talla = "" Then
                        var_nombre_talla = "SIN TALLA"
                     End If
                     var_licencia = Trim(rs1!vcha_lic_licencia_id)
                     If var_licencia = "" Then
                        var_licencia = "SIN LICENCIA"
                     End If
                     var_nombre_licencia = Trim(rs1!vcha_lic_nombre)
                     If var_nombre_licencia = "" Then
                        var_nombre_licencia = "SIN LICENCIA"
                     End If
                     var_numero_licencia = Trim(rs1!VCHA_ART_NUMERO_LIC)
                     If var_numero_licencia = "" Then
                        var_numero_licencia = "SIN NUMERO DE LICENCIA"
                     End If
                     var_pais = Trim(rs1!VCHA_PAI_PAIS_ID)
                     If var_pais = "" Then
                        var_pais = "SIN PAIS"
                     End If
                     var_nombre_pais = Trim(rs1!vcha_pai_nombre)
                     If var_nombre_pais = "" Then
                        var_nombre_pais = "SIN PAIS"
                     End If
                     
                     
                     
                     
                     var_descuento_1 = IIf(IsNull(rs1!FLOA_SAL_DESCUENTO_1), 0, rs1!FLOA_SAL_DESCUENTO_1)
                     var_descuento_2 = IIf(IsNull(rs1!FLOA_SAL_DESCUENTO_2), 0, rs1!FLOA_SAL_DESCUENTO_2)
                     VAR_PRECIO = rs1!FLOA_SAL_PRECIO - (rs1!FLOA_SAL_PRECIO * (var_descuento_1 / 100))
                     VAR_PRECIO = VAR_PRECIO - (VAR_PRECIO * (var_descuento_2 / 100))
                     VAR_PORCENTAJE_IVA = IIf(IsNull(rs1!FLOA_CAR_PORCENTAJE_IVA), 0, rs1!FLOA_CAR_PORCENTAJE_IVA)
                     var_precio_2 = VAR_PRECIO * ((1 + (VAR_PORCENTAJE_IVA / 100)))
                     var_importe_iva = var_precio_2 - VAR_PRECIO
                     var_precio_base = IIf(IsNull(rs1!MONE_ART_PRECIO_BASE), 0, rs1!MONE_ART_PRECIO_BASE) * ((1 + (VAR_PORCENTAJE_IVA / 100)))
                     VAR_PRECIO = var_precio_2
                  
                  
                     var_cadena = "insert into  VT_TB_VENTAS  (VTA_INDICE_ID, VTA_ORIGEN_ID, VTA_EMPRESA_ID, VTA_EMPRESA, VTA_UNIDAD_ORGANIZACIONAL_ID, VTA_UNIDAD_ORGANIZACIONAL,VTA_TIENDA_ID,VTA_TIENDA,VTA_CANAL_ID, VTA_CANAL, VTA_REGION,VTA_AGENTE_ID,VTA_AGENTE,VTA_RUTA_ID,VTA_RUTA,VTA_ZONA_ID,VTA_ZONA,VTA_CLIENTE_ID,VTA_CLIENTE_ID_UNFO,VTA_CLIENTE,VTA_RFC_CLIENTE,VTA_TITULAR_ID,VTA_TITULAR,VTA_GRUPO_ID,VTA_GRUPO, "
                     var_cadena = var_cadena + " VTA_ESTABLECIMIENTO_ID,VTA_ESTABLECIMIENTO,VTA_CODIGO_POSTAL,VTA_ESTADO_ID,VTA_ESTADO,VTA_CIUDAD_ID,VTA_CIUDAD,VTA_MUNICIPIO_ID,VTA_MUNICIPIO, VTA_COLONIA_ID,VTA_COLONIA,VTA_FECHA, VTA_SEMANA,VTA_ID_ENCABEZADO_SID,VTA_ID_DETALLE_SID,VTA_MOVIMIENTO_ID,VTA_TIPO_MOVIMIENTO_ID, VTA_DOCUMENTO, VTA_DESCRIPCION_DOCUMENTO, VTA_NUMERO_DOCUMENTO,VTA_SERIE,VTA_PLAZO,VTA_ARTICULO_ID,                 VTA_ARTICULO, "
                     var_cadena = var_cadena + "VTA_CATALOGO_ID,VTA_CATALOGO,VTA_DISENIO_ID,VTA_DISENIO, VTA_FAMILIA_ID, VTA_FAMILIA, VTA_LINEA_ID,VTA_LINEA, VTA_TALLA_ID,VTA_TALLA,VTA_LICENCIA_ID,VTA_LICENCIA,VTA_NUMERO_LICENCIA, "
                     var_cadena = var_cadena + "VTA_PRECIO_BASE, VTA_IMPUESTO_PORCIENTO ,VTA_FECHA_ALTA_CODIGO, VTA_CANTIDAD, VTA_COSTO, VTA_PRECIO_LISTA_A, VTA_PRECIO_LISTA_B, VTA_PRECIO_LISTA_C, VTA_DESCUENTO1 ,VTA_DESCUENTO2, VTA_DESCUENTO3, VTA_TIPO_CAMBIO, VTA_MONEDA ,VTA_IMPORTE,VTA_IMPUESTO,VTA_FLETE_ID,VTA_IMPORTE_FLETE,VTA_IMPUESTO_FLETE,VTA_MOVIMIENTO_INDEX,VTA_TOTAL_MOVIMIENTOS,VTA_CLAVE_REGISTRO, VTA_PAIS_ID, VTA_PAIS) VALUES"
                     var_cadena = var_cadena + "(sq_ventas_id.nextval," + CStr(var_origen) + ",'" + Trim(rs1!VCHA_EMP_EMPRESA_ID) + "', '" + Trim(rs1!VCHA_EMP_NOMBRE) + "','" + Trim(rs1!VCHA_UOR_UNIDAD_ID) + "','" + Trim(rs1!VCHA_UOR_NOMBRE) + "','" + Trim(rs1!VCHA_ALM_ALMACEN_ID) + "', '" + Trim(rs1!VCHA_ALM_NOMBRE) + "','" + Trim(var_canal_venta) + "','" + Trim(var_nombre_canal) + "','','" + Trim(var_agente) + "', '" + Trim(var_nombre_agente) + "','" + Trim(var_ruta) + "','" + Trim(var_nombre_ruta) + "', '" + Trim(var_zona) + "', '" + Trim(var_nombre_zona) + "', '" + Trim(var_cliente) + "','" + Trim(VCHA_CLI_CLAVE_UNIFICADA_ID) + "','" + Trim(var_nombre_cliente) + "', '" + Trim(var_rfc) + "','" + Trim(var_titular) + "', '" + Trim(var_nombre_titular) + "', '" + Trim(var_grupo) + "', '" + Trim(var_nombre_grupo) + "',"
                     var_cadena = var_cadena + "'" + Trim(var_establecimiento) + "', '" + Trim(var_nombre_establecimiento) + "', '" + Trim(var_cp) + "','" + Trim(var_estado) + "', '" + Trim(var_nombre_estado) + "', '" + Trim(var_ciudad) + "','" + Trim(var_nombre_ciudad) + "', '" + Trim(var_municipio) + "','" + Trim(var_nombre_municipio) + "','" + Trim(var_colonia) + "', '" + Trim(var_nombre_colonia) + "', TO_DATE('" + var_fecha_factura + "','DD/MM/YYYY')," + CStr(rs1!SEMANA) + ", '" + CStr(rs1!ID_DOCUMENTO) + Trim(rs1!N_SERVIDOR) + Trim(rs1!n_basedatos) + "', '" + CStr(rs1!inte_sal_consecutivo_tabla) + Trim(rs1!N_SERVIDOR) + Trim(rs1!n_basedatos) + "', '" + Trim(rs1!VCHA_MOV_MOVIMIENTO_ID) + "',  'FA',          'FA',         'FA'                 , " + CStr(rs1!inte_Car_numero) + ",'" + Trim(rs1!VCHA_SER_SERIE_ID) + "'," + CStr(rs1!INTE_CAR_PLAZO) + ",'" + Trim(rs1!VCHA_ART_ARTICULO_ID) + "', '" + Trim(rs1!VCHA_aRT_NOMBRE_ESPAÑOL) + "',"
                     var_cadena = var_cadena + "'" + Trim(var_catalogo) + "','" + Trim(var_nombre_catalogo) + "', '" + Trim(var_diseño) + "', '" + Trim(var_nombre_diseño) + "','','','" + Trim(var_linea) + "','" + Trim(var_nombre_linea) + "','" + Trim(var_talla) + "','" + Trim(var_nombre_talla) + "','" + Trim(var_licencia) + "', '" + Trim(var_nombre_licencia) + "','" + Trim(var_numero_licencia) + "',"
                     var_cadena = var_cadena + CStr(var_precio_base) + "," + CStr(rs1!FLOA_CAR_PORCENTAJE_IVA) + ",TO_DATE('" + var_fecha_alta + "','DD/MM/YYYY')," + CStr(rs1!floa_sal_cantidad) + "," + CStr(IIf(IsNull(rs1!FLOA_SAL_COSTO), 0, rs1!FLOA_SAL_COSTO)) + "," + CStr(VAR_PRECIO) + "                               ,0,0," + CStr(rs1!FLOA_SAL_DESCUENTO_1) + "," + CStr(rs1!FLOA_SAL_DESCUENTO_2) + "," + CStr(rs1!FLOA_SAL_PROMOCION_1) + "," + CStr(rs1!FLOA_CAR_TIPO_CAMBIO) + ",'" + Trim(rs1!vcha_mon_divisa) + "'," + CStr(rs1!floa_sal_cantidad * VAR_PRECIO) + "," + CStr(rs1!floa_sal_cantidad * var_importe_iva) + ",0,0,0,'','','" + Trim(rs1!VCHA_EMP_EMPRESA_ID) + Trim(rs1!VCHA_UOR_UNIDAD_ID) + "FA" + Trim(rs1!VCHA_SER_SERIE_ID) + CStr(rs1!inte_Car_numero) + "_" + CStr(rs1!inte_sal_consecutivo_tabla) + "','" + Trim(var_pais) + "','" + Trim(var_nombre_pais) + "')"
                     rs2.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                     rs3.Open "SELECT * FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS WHERE  VCHA_EMP_EMPRESA_ID = '" + rs1!VCHA_EMP_EMPRESA_ID + "' AND VCHA_CAR_TIPO_DOCUMENTO = 'FA' AND VCHA_CAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + rs1!VCHA_SER_SERIE_ID + "' AND VCHA_CLI_CLAVE_ID = '" + rs1!VCHA_CLI_CLAVE_ID + "' AND INTE_CAR_NUMERO = " + CStr(rs1!inte_Car_numero), cnn_cdindustrial, adOpenDynamic, adLockOptimistic
                     If rs3.EOF Then
                        rs4.Open "INSERT INTO TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS (VCHA_EMP_EMPRESA_ID,  VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO) VALUES ('" + rs1!VCHA_EMP_EMPRESA_ID + "','FA','FA','" + rs1!VCHA_SER_SERIE_ID + "','" + rs1!VCHA_CLI_CLAVE_ID + "'," + CStr(rs1!inte_Car_numero) + ")", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
                     End If
                     rs3.Close
                     'rs3.Open "UPDATE TB_sALIDAS SET INTE_INT_INTERFACE = 1 WHERE N_SERVIDOR = '" + rs1!SERVIDOR_DETALLE + "' AND N_BASEDATOS = '" + rs1!BASE_DATOS_DETALLE + "' AND INTE_SAL_CONSECUTIVO_TABLA = " + CStr(rs1!INTE_SAL_CONSECUTIVO_TABLA), cnn_cdindustrial, adOpenDynamic, adLockOptimistic
                     rs1.MoveNext
                     var_i = var_i + 1
                     Text1 = var_i
                     Me.Refresh
                     Me.Text1.Refresh
               Wend
            
               If rs1.RecordCount > 0 Then
                  rs1.MoveFirst
               End If
               rs2.Open "select distinct VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
               While Not rs2.EOF
                     rs3.Open "UPDATE TB_ENCABEZADO_cARTERA SET INTE_INT_INTERFACE = 1 WHERE VCHA_EMP_EMPRESA_ID = '" + rs2!VCHA_EMP_EMPRESA_ID + "' AND VCHA_CAR_TIPO_DOCUMENTO = 'FA' AND VCHA_CAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + rs2!VCHA_SER_SERIE_ID + "' AND VCHA_CLI_CLAVE_ID = '" + rs2!VCHA_CLI_CLAVE_ID + "' AND INTE_CAR_NUMERO = '" + CStr(rs2!inte_Car_numero) + "'", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
                     rs2.MoveNext
               Wend
               rs2.Close
            
               'While Not rs1.EOF
               '    rs3.Open "UPDATE TB_sALIDAS SET INTE_INT_INTERFACE = 1 WHERE VCHA_EMP_EMPRESA_ID = '" + rs1!vcha_emp_empresa_id + "' AND VCHA_UOR_UNIDAD_ID = '" + rs1!VCHA_UOR_UNIDAD_ID + "' AND VCHA_ALM_ALMACEN_ID = '" + rs1!VCHA_ALM_ALMACEN_ID + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + rs1!VCHA_MOV_MOVIMIENTO_ID + "' AND INTE_SAL_NUMERO = " + CStr(rs1!INTE_sAL_NUMERO) + " AND VCHA_ART_ARTICULO_ID = '" + rs1!VCHA_ART_aRTICULO_ID + "' AND INTE_SAL_AÑO = " + CStr(rs1!INTE_sAL_AÑO), cnn_cdindustrial, adOpenDynamic, adLockOptimistic
               '    rs1.MoveNext
               'Wend
               'rs1.Close
            
            
               rs4.MoveNext
         Wend
      End If
      rs4.Close
   
   
      If var_origen = 5 Then
         rs1.Open "DELETE FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
         var_origen = 6
         If rs1.State = 1 Then
            rs1.Close
         End If
         var_cadena = " SELECT  16 as FLOA_CAR_PORCENTAJE_IVA, dtim_emo_Fecha, ISNULL(dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID, 'SIN EMPRESA') AS VCHA_EMP_EMPRESA_ID, ISNULL(dbo.TB_EMPRESAS.VCHA_EMP_NOMBRE, 'SIN EMPRESA') AS VCHA_EMP_NOMBRE, ISNULL(dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID, 'SIN UNIDAD') AS VCHA_UOR_UNIDAD_ID, ISNULL(dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_NOMBRE, 'SIN UNIDAD') AS VCHA_UOR_NOMBRE, REPLACE(REPLACE(ISNULL(dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_NOMBRE, 'SIN ESTABLECIMIENTO'), '''', '´'), '´', '') AS VCHA_ESB_NOMBRE, { fn WEEK(TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA) } AS SEMANA, 'FACTURA' AS VTA_DESCRIPCION_DOCUMENTO, ISNULL(dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID, '') AS VCHA_ART_ARTICULO_ID, ISNULL(dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, '') AS VCHA_ART_NOMBRE_ESPAÑOL, ISNULL(dbo.TB_CATALOGOS.VCHA_CAT_CATALOGO_ID, 'SIN CATALOGO') AS VCHA_CAT_CATALOGO_ID, ISNULL(dbo.TB_CATALOGOS.VCHA_CAT_NOMBRE, ' SIN CATALOGO ') AS VCHA_CAT_NOMBRE, ISNULL(dbo.TB_DISEÑOS.VCHA_DIS_DISEÑO_ID, "
         var_cadena = var_cadena + "'SIN DISEÑO') AS VCHA_DIS_DISEÑO_ID, ISNULL(dbo.TB_DISEÑOS.VCHA_DIS_NOMBRE, ' SIN DISEÑO ') AS VCHA_DIS_NOMBRE, ISNULL(dbo.TB_LINEAS.VCHA_LIN_LINEA_ID, 'SIN LINEA') AS VCHA_LIN_LINEA_ID, ISNULL(dbo.TB_LINEAS.VCHA_LIN_NOMBRE, 'SIN LINEA') AS VCHA_LIN_NOMBRE, ISNULL(dbo.TB_TALLAS.VCHA_TAL_TALLA_ID, 'SIN TALLA') AS VCHA_TAL_TALLA_ID, ISNULL(dbo.TB_TALLAS.VCHA_TAL_NOMBRE, 'SIN TALLA') AS VCHA_TAL_NOMBRE, ISNULL(dbo.TB_LICENCIAS.VCHA_LIC_LICENCIA_ID, 'SIN LICENCIA') AS VCHA_LIC_LICENCIA_ID, ISNULL(dbo.TB_LICENCIAS.VCHA_LIC_NOMBRE, 'SIN LICENCIA') AS VCHA_LIC_NOMBRE, ISNULL(dbo.TB_ARTICULOS.VCHA_ART_NUMERO_LIC, 'SIN NUMERO DE LICENCIA') AS VCHA_ART_NUMERO_LIC, ISNULL(dbo.TB_ARTICULOS.MONE_ART_PRECIO_BASE, 0) AS MONE_ART_PRECIO_BASE, ISNULL(dbo.TB_ARTICULOS.DTIM_ART_FECHA_ALTA, GETDATE()) AS DTIM_ART_FECHA_ALTA, ISNULL(dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, 0) AS FLOA_SAL_CANTIDAD, inte_emo_numero as inte_car_numero, "
         var_cadena = var_cadena + "dbo.TB_SALIDAS.FLOA_SAL_COSTO, dbo.TB_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2, dbo.TB_SALIDAS.FLOA_SAL_PROMOCION_1, dbo.TB_SALIDAS.INTE_INT_INTERFACE AS INTERFACE_DETALLE, dbo.TB_SALIDAS.INTE_SAL_CONSECUTIVO_TABLA, dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_ZON_DESCRIPCION, 'SIN ZONA') AS VCHA_ZON_NOMBRE, dbo.TB_SALIDAS.N_SERVIDOR AS SERVIDOR_DETALLE, dbo.TB_SALIDAS.N_BASEDATOS AS BASE_DATOS_DETALLE, dbo.TB_SALIDAS.INTE_SAL_AÑO, dbo.TB_SALIDAS.INTE_SAL_NUMERO, "
         var_cadena = var_cadena + "dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, ISNULL(dbo.VW_CLIENTES.VCHA_CAN_CANAL_VENTA_ID, 'SIN CANAL') AS VCHA_CAN_CANAL_VENTA_ID,  ISNULL(dbo.VW_CLIENTES.VCHA_CAN_NOMBRE, 'SIN CANAL') AS VCHA_CAN_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_AGE_AGENTE_ID, 'SIN AGENTE') AS VCHA_AGE_AGENTE_ID, ISNULL(dbo.VW_CLIENTES.VCHA_AGE_NOMBRE, 'SIN AGENTE') AS VCHA_AGE_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_RUT_RUTA_ID, 'SIN RUTA') AS VCHA_RUT_RUTA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_RUT_NOMBRE, 'SIN RUTA') AS VCHA_RUT_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_ZON_ZONA_ID, 'SIN ZONA') AS VCHA_ZON_ZONA_ID, "
         var_cadena = var_cadena + "ISNULL(dbo.VW_CLIENTES.VCHA_ZON_DESCRIPCION, 'SIN ZONA') AS VCHA_ZON_DESCRIPCION, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID, 'SIN CLIENTE') AS VCHA_CLI_CLAVE_ID, replace(replace(ISNULL(dbo.VW_CLIENTES.VCHA_CLI_NOMBRE, 'SIN CLIENTE'),'''','´'),'´','') AS VCHA_CLI_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_CLAVE_UNIFICADA_ID, '0') AS VCHA_CLI_CLAVE_UNIFICADA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_RFC, 'SIN RFC') AS VCHA_CLI_RFC, ISNULL(dbo.VW_CLIENTES.VCHA_TIT_TITULAR_ID, 'SIN TITULAR') AS VCHA_TIT_TITULAR_ID,     replace(replace(ISNULL(dbo.VW_CLIENTES.VCHA_TIT_NOMBRE, 'SIN TITULAR'),'''','´'),'´','') AS VCHA_TIT_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_GAC_GRUPO_ACTUAL_ID, 'SIN GRUPO ') AS VCHA_GAC_GRUPO_ACTUAL_ID, replace(replace(ISNULL(dbo.VW_CLIENTES.VCHA_GAC_NOMBRE, 'SIN GRUPO'),'''','´'),'´','')  AS VCHA_GAC_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_CP, 'SIN CP') AS VCHA_CLI_CP, ISNULL(dbo.VW_CLIENTES.VCHA_EST_ESTADO_ID, 'SIN ESTADO') "
         var_cadena = var_cadena + "AS VCHA_EST_ESTADO_ID, ISNULL(dbo.VW_CLIENTES.VCHA_EST_NOMBRE, 'SIN ESTADO') AS VCHA_EST_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_CIU_CIUDAD_ID, 'SIN CIUDAD') AS VCHA_CIU_CIUDAD_ID, ISNULL(dbo.VW_CLIENTES.VCHA_CIU_NOMBRE, 'SIN CIUDAD') AS VCHA_CIU_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_MUN_MUNICIPIO_ID, 'SIN MUNICIPIO') AS VCHA_MUN_MUNICIPIO_ID, ISNULL(dbo.VW_CLIENTES.VCHA_MUN_NOMBRE, 'SIN MUNICIPIO') AS VCHA_MUN_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_COL_COLONIA_ID, 'SIN COLONIA') AS VCHA_COL_COLONIA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_COL_NOMBRE, 'SIN COLONIA') AS VCHA_COL_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_MON_MONEDA_ID, 'SIN MONEDA') AS VCHA_MON_MONEDA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_MON_DIVISA, 'SIN MONEDA') AS VCHA_MON_DIVISA, ISNULL(dbo.VW_CLIENTES.VCHA_ZON_DESCRIPCION, 'SIN ZONA') AS VCHA_ZON_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_PAI_PAIS_ID, '') AS VCHA_PAI_PAIS_ID, ISNULL(dbo.VW_CLIENTES.VCHA_PAI_NOMBRE, '') AS VCHA_PAI_NOMBRE FROM dbo.TB_LINEAS RIGHT OUTER JOIN "
         var_cadena = var_cadena + "dbo.TB_ALMACENES INNER JOIN dbo.TB_UNIDADESORGANIZACIONALES INNER JOIN dbo.TB_SALIDAS ON dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID ON dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID INNER JOIN dbo.TB_ESTABLECIMIENTOS INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND "
         var_cadena = var_cadena + " dbo.TB_SALIDAS.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN dbo.TB_EMPRESAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID INNER JOIN dbo.VW_CLIENTES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_CLI_CLAVE_ID = dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID LEFT OUTER JOIN dbo.TB_LICENCIAS ON dbo.TB_ARTICULOS.VCHA_LIC_LICENCIA_ID = dbo.TB_LICENCIAS.VCHA_LIC_LICENCIA_ID LEFT OUTER JOIN dbo.TB_TALLAS ON dbo.TB_ARTICULOS.VCHA_TAL_TALLA_ID = dbo.TB_TALLAS.VCHA_TAL_TALLA_ID ON dbo.TB_LINEAS.VCHA_LIN_LINEA_ID = dbo.TB_ARTICULOS.VCHA_LIN_LINEA_ID LEFT OUTER JOIN dbo.TB_DISEÑOS ON dbo.TB_ARTICULOS.VCHA_DIS_DISEÑO_ID = dbo.TB_DISEÑOS.VCHA_DIS_DISEÑO_ID LEFT OUTER JOIN dbo.TB_CATALOGOS ON dbo.TB_ARTICULOS.VCHA_ART_CATALOGO_VIGENTE = dbo.TB_CATALOGOS.VCHA_CAT_CATALOGO_ID WHERE (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'CC_2') AND DTIM_EMO_FECHA >= CONVERT(DATETIME, '" + VAR_FECHA_INICIO + "', 102) AND "
         var_cadena = var_cadena + " DTIM_EMO_FECHA <  CONVERT(DATETIME, '" + VAR_FECHA_FIN + "', 102) and inte_emo_numero_origen = 0"
         rs1.Open var_cadena, cnn_cdindustrial, adOpenDynamic, adLockOptimistic
         While Not rs1.EOF
               var_dia = CStr(Day(CDate(rs1!DTIM_emo_FECHA)))
               var_mes = CStr(Month(CDate(rs1!DTIM_emo_FECHA)))
               var_año = CStr(Year(CDate(rs1!DTIM_emo_FECHA)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_factura = var_dia + "/" + var_mes + "/" + var_año
               var_dia = CStr(Day(CDate(rs1!DTIM_ART_FECHA_ALTA)))
               var_mes = CStr(Month(CDate(rs1!DTIM_ART_FECHA_ALTA)))
               var_año = CStr(Year(CDate(rs1!DTIM_ART_FECHA_ALTA)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_alta = var_dia + "/" + var_mes + "/" + var_año
               var_canal_venta = Trim(rs1!VCHA_CAN_CANAL_VENTA_ID)
               If var_canal_venta = "" Then
                  var_canal_venta = "SIN CANAL"
               End If
               var_nombre_canal = Trim(rs1!VCHA_CAN_NOMBRE)
               If var_nombre_canal = "" Then
                  var_nombre_canal = "SIN CANAL"
               End If
               var_agente = Trim(rs1!VCHA_AGE_AGENTE_ID)
               If var_agente = "" Then
                  var_agente = "SIN AGENTE"
               End If
               var_nombre_agente = Trim(rs1!VCHA_AGE_NOMBRE)
               If var_nombre_agente = "" Then
                  var_nombre_agente = "SIN AGENTE"
               End If
               var_ruta = Trim(rs1!VCHA_RUT_RUTA_ID)
               If var_ruta = "" Then
                  var_ruta = "SIN RUTA"
               End If
               var_nombre_ruta = Trim(rs1!VCHA_RUT_NOMBRE)
               If var_nombre_ruta = "" Then
                  var_nombre_ruta = "SIN RUTA"
               End If
               var_zona = Trim(rs1!VCHA_ZON_ZONA_ID)
               If var_zona = "" Then
                  var_zona = "SIN ZONA"
               End If
               var_nombre_zona = Trim(rs1!VCHA_ZON_NOMBRE)
               If var_nombre_zona = "" Then
                  var_nombre_zona = "SIN ZONA"
               End If
               var_cliente = Trim(rs1!VCHA_CLI_CLAVE_ID)
               If var_cliente = "" Then
                  var_cliente = "SIN CLIENTE"
               End If
               var_cliente_unfo = Trim(VCHA_CLI_CLAVE_UNIFICADA_ID)
               If var_cliente_unfo = "" Then
                  var_cliente_unfo = "0"
               End If
               var_nombre_cliente = Trim(rs1!VCHA_CLI_NOMBRE)
               If var_nombre_cliente = "" Then
                  var_nombre_cliente = "SIN CLIENTE"
               End If
               var_rfc = Trim(rs1!VCHA_CLI_RFC)
               var_titular = Trim(rs1!VCHA_TIT_TITULAR_ID)
               If var_titular = "" Then
                  var_titular = "SIN TITULAR"
               End If
               var_nombre_titular = Trim(rs1!VCHA_TIT_NOMBRE)
               If var_nombre_titular = "" Then
                  var_nombre_titular = "SIN TITULAR"
               End If
               var_grupo = Trim(rs1!VCHA_GAC_GRUPO_ACTUAL_ID)
               If var_grupo = "" Then
                  var_grupo = "SIN GRUPO"
               End If
               var_nombre_grupo = Trim(rs1!VCHA_GAC_NOMBRE)
               If var_nombre_grupo = "" Then
                  var_nombre_grupo = "SIN GRUPO"
               End If
               var_establecimiento = ""
               If var_establecimiento = "" Then
                  var_establecimiento = "SIN ESTABLECIMIENTO"
               End If
               var_nombre_establecimiento = ""
               If var_nombre_establecimiento = "" Then
                  var_nombre_establecimiento = "SIN ESTABLECIMIENTO"
               End If
               var_cp = Trim(rs1!VCHA_CLI_CP)
               If var_cp = "" Then
                  var_cp = "SIN CP"
               End If
               var_estado = Trim(rs1!VCHA_EST_ESTADO_ID)
               If var_estado = "" Then
                  var_estado = "SIN ESTADO"
               End If
               var_nombre_estado = Trim(rs1!VCHA_EST_NOMBRE)
               If var_nombre_estado = "" Then
                  var_nombre_estado = "SIN ESTADO"
               End If
               var_ciudad = Trim(rs1!VCHA_CIU_CIUDAD_ID)
               If var_ciudad = "" Then
                  var_ciudad = "SIN CIUDAD"
               End If
               var_nombre_ciudad = Trim(rs1!VCHA_CIU_NOMBRE)
               If var_nombre_ciudad = "" Then
                  var_nombre_ciudad = "SIN CIUDAD"
               End If
               var_municipio = Trim(rs1!VCHA_MUN_MUNICIPIO_ID)
               If var_municipio = "" Then
                  var_municipio = "SIN MUNICIPIO"
               End If
               var_nombre_municipio = Trim(rs1!VCHA_MUN_NOMBRE)
               If var_nombre_municipio = "" Then
                  var_nombre_municipio = "SIN MUNICIPIO"
               End If
               var_colonia = Trim(rs1!vcha_col_colonia_id)
               If var_colonia = "" Then
                  var_colonia = "SIN COLONIA"
               End If
               var_nombre_colonia = Trim(rs1!VCHA_COL_NOMBRE)
               If var_nombre_colonia = "" Then
                  var_nombre_colonia = "SIN COLONIA"
               End If
               var_articulo = Trim(rs1!VCHA_ART_ARTICULO_ID)
               If var_articulo = "" Then
                  var_articulo = "SIN ARTICULO"
               End If
               var_nombre_articulo = Trim(rs1!VCHA_aRT_NOMBRE_ESPAÑOL)
               If var_nombre_articulo = "" Then
                  var_nombre_articulo = "SIN ARTICULO"
               End If
               var_catalogo = Trim(rs1!VCHA_CAT_CATALOGO_ID)
               If var_catalogo = "" Then
                  var_catalogo = "SIN CATALOGO"
               End If
               var_nombre_catalogo = Trim(rs1!VCHA_CAT_NOMBRE)
               If var_nombre_catalogo = "" Then
                  var_nombre_catalogo = "SIN CATALOGO"
               End If
               var_diseño = Trim(rs1!VCHA_DIS_DISEÑO_ID)
               If var_diseño = "" Then
                  var_diseño = "SIN DISEÑO"
               End If
               var_nombre_diseño = Trim(rs1!VCHA_DIS_NOMBRE)
               If var_nombre_diseño = "" Then
                  var_nombre_diseño = "SIN DISEÑO"
               End If
               var_linea = Trim(rs1!VCHA_LIN_LINEA_ID)
               If var_linea = "" Then
                  var_linea = "SIN LINEA"
               End If
               var_nombre_linea = Trim(rs1!VCHA_LIN_NOMBRE)
               If var_nombre_linea = "" Then
                  var_nombre_linea = "SIN LINEA"
               End If
               var_talla = Trim(rs1!vcha_Tal_talla_id)
               If var_talla = "" Then
                  var_talla = "SIN TALLA"
               End If
               var_nombre_talla = Trim(rs1!VCHA_TAL_NOMBRE)
               If var_nombre_talla = "" Then
                  var_nombre_talla = "SIN TALLA"
               End If
               var_licencia = Trim(rs1!vcha_lic_licencia_id)
               If var_licencia = "" Then
                  var_licencia = "SIN LICENCIA"
               End If
               var_nombre_licencia = Trim(rs1!vcha_lic_nombre)
               If var_nombre_licencia = "" Then
                  var_nombre_licencia = "SIN LICENCIA"
               End If
               var_numero_licencia = Trim(rs1!VCHA_ART_NUMERO_LIC)
               If var_numero_licencia = "" Then
                  var_numero_licencia = "SIN NUMERO DE LICENCIA"
               End If
               var_pais = Trim(rs1!VCHA_PAI_PAIS_ID)
               If var_pais = "" Then
                  var_pais = "SIN PAIS"
               End If
               var_nombre_pais = Trim(rs1!vcha_pai_nombre)
               If var_nombre_pais = "" Then
                  var_nombre_pais = "SIN PAIS"
               End If
        
           
               var_descuento_1 = IIf(IsNull(rs1!FLOA_SAL_DESCUENTO_1), 0, rs1!FLOA_SAL_DESCUENTO_1)
               var_descuento_2 = IIf(IsNull(rs1!FLOA_SAL_DESCUENTO_2), 0, rs1!FLOA_SAL_DESCUENTO_2)
               VAR_PRECIO = rs1!FLOA_SAL_PRECIO - (rs1!FLOA_SAL_PRECIO * (var_descuento_1 / 100))
               VAR_PRECIO = VAR_PRECIO - (VAR_PRECIO * (var_descuento_2 / 100))
               VAR_PORCENTAJE_IVA = IIf(IsNull(rs1!FLOA_CAR_PORCENTAJE_IVA), 0, rs1!FLOA_CAR_PORCENTAJE_IVA)
               var_precio_2 = VAR_PRECIO * ((1 + (VAR_PORCENTAJE_IVA / 100)))
               var_importe_iva = var_precio_2 - VAR_PRECIO
               VAR_PRECIO = var_precio_2
               var_precio_base = IIf(IsNull(rs1!MONE_ART_PRECIO_BASE), 0, rs1!MONE_ART_PRECIO_BASE) * ((1 + (VAR_PORCENTAJE_IVA / 100)))
                      
               var_cadena = "insert into  VT_TB_VENTAS  (VTA_INDICE_ID, VTA_ORIGEN_ID, VTA_EMPRESA_ID, VTA_EMPRESA, VTA_UNIDAD_ORGANIZACIONAL_ID, VTA_UNIDAD_ORGANIZACIONAL,VTA_TIENDA_ID,VTA_TIENDA,VTA_CANAL_ID, VTA_CANAL, VTA_REGION,VTA_AGENTE_ID,VTA_AGENTE,VTA_RUTA_ID,VTA_RUTA,VTA_ZONA_ID,VTA_ZONA,VTA_CLIENTE_ID,VTA_CLIENTE_ID_UNFO,VTA_CLIENTE,VTA_RFC_CLIENTE,VTA_TITULAR_ID,VTA_TITULAR,VTA_GRUPO_ID,VTA_GRUPO, "
               var_cadena = var_cadena + " VTA_ESTABLECIMIENTO_ID,VTA_ESTABLECIMIENTO,VTA_CODIGO_POSTAL,VTA_ESTADO_ID,VTA_ESTADO,VTA_CIUDAD_ID,VTA_CIUDAD,VTA_MUNICIPIO_ID,VTA_MUNICIPIO, VTA_COLONIA_ID,VTA_COLONIA,VTA_FECHA, VTA_SEMANA,VTA_ID_ENCABEZADO_SID,VTA_ID_DETALLE_SID,VTA_MOVIMIENTO_ID,VTA_TIPO_MOVIMIENTO_ID, VTA_DOCUMENTO, VTA_DESCRIPCION_DOCUMENTO, VTA_NUMERO_DOCUMENTO,VTA_SERIE,VTA_PLAZO,VTA_ARTICULO_ID,                 VTA_ARTICULO, "
               var_cadena = var_cadena + "VTA_CATALOGO_ID,VTA_CATALOGO,VTA_DISENIO_ID,VTA_DISENIO, VTA_FAMILIA_ID, VTA_FAMILIA, VTA_LINEA_ID,VTA_LINEA, VTA_TALLA_ID,VTA_TALLA,VTA_LICENCIA_ID,VTA_LICENCIA,VTA_NUMERO_LICENCIA, "
               var_cadena = var_cadena + "VTA_PRECIO_BASE, VTA_IMPUESTO_PORCIENTO ,VTA_FECHA_ALTA_CODIGO, VTA_CANTIDAD, VTA_COSTO, VTA_PRECIO_LISTA_A, VTA_PRECIO_LISTA_B, VTA_PRECIO_LISTA_C, VTA_DESCUENTO1 ,VTA_DESCUENTO2, VTA_DESCUENTO3, VTA_TIPO_CAMBIO, VTA_MONEDA ,VTA_IMPORTE,VTA_FLETE_ID,VTA_IMPORTE_FLETE,VTA_IMPUESTO_FLETE,VTA_MOVIMIENTO_INDEX,VTA_TOTAL_MOVIMIENTOS,VTA_CLAVE_REGISTRO, VTA_PAIS_ID, VTA_PAIS, VTA_IMPUESTO, precio) VALUES"
               var_cadena = var_cadena + "(sq_ventas_id.nextval," + CStr(var_origen) + ",'" + Trim(rs1!VCHA_EMP_EMPRESA_ID) + "', '" + Trim(rs1!VCHA_EMP_NOMBRE) + "','" + Trim(rs1!VCHA_UOR_UNIDAD_ID) + "','" + Trim(rs1!VCHA_UOR_NOMBRE) + "','" + Trim(rs1!VCHA_ALM_ALMACEN_ID) + "', '" + Trim(rs1!VCHA_ALM_NOMBRE) + "','" + Trim(var_canal_venta) + "','" + Trim(var_nombre_canal) + "','','" + Trim(var_agente) + "', '" + Trim(var_nombre_agente) + "','" + Trim(var_ruta) + "','" + Trim(var_nombre_ruta) + "', '" + Trim(var_zona) + "', '" + Trim(var_nombre_zona) + "', '" + Trim(var_cliente) + "','" + Trim(VCHA_CLI_CLAVE_UNIFICADA_ID) + "','" + Trim(var_nombre_cliente) + "', '" + Trim(var_rfc) + "','" + Trim(var_titular) + "', '" + Trim(var_nombre_titular) + "', '" + Trim(var_grupo) + "', '" + Trim(var_nombre_grupo) + "',"
               var_cadena = var_cadena + "'" + Trim(var_establecimiento) + "', '" + Trim(var_nombre_establecimiento) + "', '" + Trim(var_cp) + "','" + Trim(var_estado) + "', '" + Trim(var_nombre_estado) + "', '" + Trim(var_ciudad) + "','" + Trim(var_nombre_ciudad) + "', '" + Trim(var_municipio) + "','" + Trim(var_nombre_municipio) + "','" + Trim(var_colonia) + "', '" + Trim(var_nombre_colonia) + "', TO_DATE('" + var_fecha_factura + "','DD/MM/YYYY')," + CStr(rs1!SEMANA) + ", '" + CStr(rs1!inte_Car_numero) + Trim("sqlquezada2") + Trim("sicantia") + "', '" + CStr(rs1!inte_sal_consecutivo_tabla) + Trim("sqlquezada2") + Trim("sidcantia") + "', '" + Trim(rs1!VCHA_MOV_MOVIMIENTO_ID) + "',  'FA',          'FA',         'FA'                 , " + CStr(rs1!inte_Car_numero) + ",'TC',0,'" + Trim(rs1!VCHA_ART_ARTICULO_ID) + "', '" + Trim(rs1!VCHA_aRT_NOMBRE_ESPAÑOL) + "',"
               var_cadena = var_cadena + "'" + Trim(var_catalogo) + "','" + Trim(var_nombre_catalogo) + "', '" + Trim(var_diseño) + "', '" + Trim(var_nombre_diseño) + "','','','" + Trim(var_linea) + "','" + Trim(var_nombre_linea) + "','" + Trim(var_talla) + "','" + Trim(var_nombre_talla) + "','" + Trim(var_licencia) + "', '" + Trim(var_nombre_licencia) + "','" + Trim(var_numero_licencia) + "',"
               var_cadena = var_cadena + CStr(var_precio_base) + "," + CStr(rs1!FLOA_CAR_PORCENTAJE_IVA) + ",TO_DATE('" + var_fecha_alta + "','DD/MM/YYYY')," + CStr(rs1!floa_sal_cantidad) + "," + CStr(IIf(IsNull(rs1!FLOA_SAL_COSTO), 0, rs1!FLOA_SAL_COSTO)) + "," + CStr(VAR_PRECIO) + ",0,0," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(1) + ",'" + Trim(rs1!vcha_mon_divisa) + "'," + CStr(rs1!floa_sal_cantidad * VAR_PRECIO) + ",0,0,0,'','','" + Trim(rs1!VCHA_EMP_EMPRESA_ID) + Trim(rs1!VCHA_UOR_UNIDAD_ID) + "FA" + Trim("TC") + CStr(rs1!inte_Car_numero) + "_" + CStr(rs1!inte_sal_consecutivo_tabla) + "','" + Trim(var_pais) + "','" + Trim(var_nombre_pais) + "', " + CStr(rs1!floa_sal_cantidad * var_importe_iva) + "," + CStr(rs1!FLOA_SAL_PRECIO) + ")"
               
            
               rs2.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
            
               rs3.Open "SELECT * FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS WHERE  VCHA_EMP_EMPRESA_ID = '" + rs1!VCHA_EMP_EMPRESA_ID + "' AND VCHA_CAR_TIPO_DOCUMENTO = 'FA' AND VCHA_CAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = 'TC' AND VCHA_CLI_CLAVE_ID = '" + rs1!VCHA_CLI_CLAVE_ID + "' AND INTE_CAR_NUMERO = " + CStr(rs1!inte_Car_numero), cnn_cdindustrial, adOpenDynamic, adLockOptimistic
               If rs3.EOF Then
                  rs4.Open "INSERT INTO TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS (VCHA_EMP_EMPRESA_ID,  VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO) VALUES ('" + rs1!VCHA_EMP_EMPRESA_ID + "','FA','FA','TC','" + rs1!VCHA_CLI_CLAVE_ID + "'," + CStr(rs1!inte_Car_numero) + ")", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
               End If
               rs3.Close
               
               rs1.MoveNext
               var_i = var_i + 1
               Text1 = var_i
               Me.Refresh
               Me.Text1.Refresh
               'Me.lbl_accion.Refresh
         Wend
         rs1.Close
      
         rs2.Open "select distinct VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
         var_i = 0
         While Not rs2.EOF
               rs4.Open "UPDATE TB_ENCABEZADO_MOVIMIENTOS SET INTE_EMO_NUMERO_ORIGEN = 1 WHeRE vcha_emp_empresa_id = '" + rs2!VCHA_EMP_EMPRESA_ID + "' and vcha_mov_movimiento_id = 'CC_2' and inte_emo_numero = " + CStr(rs2!inte_Car_numero), cnn_cdindustrial, adOpenDynamic, adLockOptimistic
               rs2.MoveNext
               var_i = var_i + 1
               Text1.Text = var_id
               Me.Refresh
               Me.Text1.Refresh
               Me.lbl_accion.Refresh
         Wend
         rs2.Close
         If rs4.State = 1 Then
            rs4.Close
         End If
         rs4.Open "SELECT VCHA_EMP_EMPRESA_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, vcha_cli_clave_id from TB_ENCABEZADO_movimientos WHERE INTE_EMO_NUMERO_ORIGEN = 2 and VCHA_MOV_MOVIMIENTO_id = 'CC_2'", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
         While Not rs4.EOF
               rs1.Open "DELETE FROM vt_tb_ventas where VTA_ORIGEN_ID = " + CStr(var_origen) + " and VTA_EMPRESA_ID = '" + rs4!VCHA_EMP_EMPRESA_ID + "' And vta_documento = 'FA' and vta_serie = 'TC' and vta_numero_documento = " + CStr(rs4!INTE_EMO_NUMERO) + " and vta_cliente_id = '" + rs4!VCHA_CLI_CLAVE_ID + "'", cnnoracle, adOpenDynamic, adLockOptimistic
               rs4.MoveNext
         Wend
         rs4.Close
       
         var_origen = 6
      
      
         rs1.Open "DELETE FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
         var_origen = 6
         If rs1.State = 1 Then
            rs1.Close
         End If
         var_cadena = " SELECT  16 as FLOA_CAR_PORCENTAJE_IVA, dtim_emo_Fecha, ISNULL(dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID, 'SIN EMPRESA') AS VCHA_EMP_EMPRESA_ID, ISNULL(dbo.TB_EMPRESAS.VCHA_EMP_NOMBRE, 'SIN EMPRESA') AS VCHA_EMP_NOMBRE, ISNULL(dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID, 'SIN UNIDAD') AS VCHA_UOR_UNIDAD_ID, ISNULL(dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_NOMBRE, 'SIN UNIDAD') AS VCHA_UOR_NOMBRE, REPLACE(REPLACE(ISNULL(dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_NOMBRE, 'SIN ESTABLECIMIENTO'), '''', '´'), '´', '') AS VCHA_ESB_NOMBRE, { fn WEEK(TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA) } AS SEMANA, 'FACTURA' AS VTA_DESCRIPCION_DOCUMENTO, ISNULL(dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID, '') AS VCHA_ART_ARTICULO_ID, ISNULL(dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, '') AS VCHA_ART_NOMBRE_ESPAÑOL, ISNULL(dbo.TB_CATALOGOS.VCHA_CAT_CATALOGO_ID, 'SIN CATALOGO') AS VCHA_CAT_CATALOGO_ID, ISNULL(dbo.TB_CATALOGOS.VCHA_CAT_NOMBRE, ' SIN CATALOGO ') AS VCHA_CAT_NOMBRE, ISNULL(dbo.TB_DISEÑOS.VCHA_DIS_DISEÑO_ID, "
         var_cadena = var_cadena + "'SIN DISEÑO') AS VCHA_DIS_DISEÑO_ID, ISNULL(dbo.TB_DISEÑOS.VCHA_DIS_NOMBRE, ' SIN DISEÑO ') AS VCHA_DIS_NOMBRE, ISNULL(dbo.TB_LINEAS.VCHA_LIN_LINEA_ID, 'SIN LINEA') AS VCHA_LIN_LINEA_ID, ISNULL(dbo.TB_LINEAS.VCHA_LIN_NOMBRE, 'SIN LINEA') AS VCHA_LIN_NOMBRE, ISNULL(dbo.TB_TALLAS.VCHA_TAL_TALLA_ID, 'SIN TALLA') AS VCHA_TAL_TALLA_ID, ISNULL(dbo.TB_TALLAS.VCHA_TAL_NOMBRE, 'SIN TALLA') AS VCHA_TAL_NOMBRE, ISNULL(dbo.TB_LICENCIAS.VCHA_LIC_LICENCIA_ID, 'SIN LICENCIA') AS VCHA_LIC_LICENCIA_ID, ISNULL(dbo.TB_LICENCIAS.VCHA_LIC_NOMBRE, 'SIN LICENCIA') AS VCHA_LIC_NOMBRE, ISNULL(dbo.TB_ARTICULOS.VCHA_ART_NUMERO_LIC, 'SIN NUMERO DE LICENCIA') AS VCHA_ART_NUMERO_LIC, ISNULL(dbo.TB_ARTICULOS.MONE_ART_PRECIO_BASE, 0) AS MONE_ART_PRECIO_BASE, ISNULL(dbo.TB_ARTICULOS.DTIM_ART_FECHA_ALTA, GETDATE()) AS DTIM_ART_FECHA_ALTA, ISNULL(dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, 0) AS FLOA_SAL_CANTIDAD, inte_emo_numero as inte_car_numero, "
         var_cadena = var_cadena + "dbo.TB_SALIDAS.FLOA_SAL_COSTO, dbo.TB_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2, dbo.TB_SALIDAS.FLOA_SAL_PROMOCION_1, dbo.TB_SALIDAS.INTE_INT_INTERFACE AS INTERFACE_DETALLE, dbo.TB_SALIDAS.INTE_SAL_CONSECUTIVO_TABLA, dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_ZON_DESCRIPCION, 'SIN ZONA') AS VCHA_ZON_NOMBRE, dbo.TB_SALIDAS.N_SERVIDOR AS SERVIDOR_DETALLE, dbo.TB_SALIDAS.N_BASEDATOS AS BASE_DATOS_DETALLE, dbo.TB_SALIDAS.INTE_SAL_AÑO, dbo.TB_SALIDAS.INTE_SAL_NUMERO, "
         var_cadena = var_cadena + "dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, ISNULL(dbo.VW_CLIENTES.VCHA_CAN_CANAL_VENTA_ID, 'SIN CANAL') AS VCHA_CAN_CANAL_VENTA_ID,  ISNULL(dbo.VW_CLIENTES.VCHA_CAN_NOMBRE, 'SIN CANAL') AS VCHA_CAN_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_AGE_AGENTE_ID, 'SIN AGENTE') AS VCHA_AGE_AGENTE_ID, ISNULL(dbo.VW_CLIENTES.VCHA_AGE_NOMBRE, 'SIN AGENTE') AS VCHA_AGE_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_RUT_RUTA_ID, 'SIN RUTA') AS VCHA_RUT_RUTA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_RUT_NOMBRE, 'SIN RUTA') AS VCHA_RUT_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_ZON_ZONA_ID, 'SIN ZONA') AS VCHA_ZON_ZONA_ID, "
         var_cadena = var_cadena + "ISNULL(dbo.VW_CLIENTES.VCHA_ZON_DESCRIPCION, 'SIN ZONA') AS VCHA_ZON_DESCRIPCION, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID, 'SIN CLIENTE') AS VCHA_CLI_CLAVE_ID, replace(replace(ISNULL(dbo.VW_CLIENTES.VCHA_CLI_NOMBRE, 'SIN CLIENTE'),'''','´'),'´','') AS VCHA_CLI_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_CLAVE_UNIFICADA_ID, '0') AS VCHA_CLI_CLAVE_UNIFICADA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_RFC, 'SIN RFC') AS VCHA_CLI_RFC, ISNULL(dbo.VW_CLIENTES.VCHA_TIT_TITULAR_ID, 'SIN TITULAR') AS VCHA_TIT_TITULAR_ID,     replace(replace(ISNULL(dbo.VW_CLIENTES.VCHA_TIT_NOMBRE, 'SIN TITULAR'),'''','´'),'´','') AS VCHA_TIT_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_GAC_GRUPO_ACTUAL_ID, 'SIN GRUPO ') AS VCHA_GAC_GRUPO_ACTUAL_ID, replace(replace(ISNULL(dbo.VW_CLIENTES.VCHA_GAC_NOMBRE, 'SIN GRUPO'),'''','´'),'´','')  AS VCHA_GAC_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_CLI_CP, 'SIN CP') AS VCHA_CLI_CP, ISNULL(dbo.VW_CLIENTES.VCHA_EST_ESTADO_ID, 'SIN ESTADO') "
         var_cadena = var_cadena + "AS VCHA_EST_ESTADO_ID, ISNULL(dbo.VW_CLIENTES.VCHA_EST_NOMBRE, 'SIN ESTADO') AS VCHA_EST_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_CIU_CIUDAD_ID, 'SIN CIUDAD') AS VCHA_CIU_CIUDAD_ID, ISNULL(dbo.VW_CLIENTES.VCHA_CIU_NOMBRE, 'SIN CIUDAD') AS VCHA_CIU_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_MUN_MUNICIPIO_ID, 'SIN MUNICIPIO') AS VCHA_MUN_MUNICIPIO_ID, ISNULL(dbo.VW_CLIENTES.VCHA_MUN_NOMBRE, 'SIN MUNICIPIO') AS VCHA_MUN_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_COL_COLONIA_ID, 'SIN COLONIA') AS VCHA_COL_COLONIA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_COL_NOMBRE, 'SIN COLONIA') AS VCHA_COL_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_MON_MONEDA_ID, 'SIN MONEDA') AS VCHA_MON_MONEDA_ID, ISNULL(dbo.VW_CLIENTES.VCHA_MON_DIVISA, 'SIN MONEDA') AS VCHA_MON_DIVISA, ISNULL(dbo.VW_CLIENTES.VCHA_ZON_DESCRIPCION, 'SIN ZONA') AS VCHA_ZON_NOMBRE, ISNULL(dbo.VW_CLIENTES.VCHA_PAI_PAIS_ID, '') AS VCHA_PAI_PAIS_ID, ISNULL(dbo.VW_CLIENTES.VCHA_PAI_NOMBRE, '') AS VCHA_PAI_NOMBRE FROM dbo.TB_LINEAS RIGHT OUTER JOIN "
         var_cadena = var_cadena + "dbo.TB_ALMACENES INNER JOIN dbo.TB_UNIDADESORGANIZACIONALES INNER JOIN dbo.TB_SALIDAS ON dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID ON dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID INNER JOIN dbo.TB_ESTABLECIMIENTOS INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND "
         var_cadena = var_cadena + " dbo.TB_SALIDAS.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN dbo.TB_EMPRESAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID INNER JOIN dbo.VW_CLIENTES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_CLI_CLAVE_ID = dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID LEFT OUTER JOIN dbo.TB_LICENCIAS ON dbo.TB_ARTICULOS.VCHA_LIC_LICENCIA_ID = dbo.TB_LICENCIAS.VCHA_LIC_LICENCIA_ID LEFT OUTER JOIN dbo.TB_TALLAS ON dbo.TB_ARTICULOS.VCHA_TAL_TALLA_ID = dbo.TB_TALLAS.VCHA_TAL_TALLA_ID ON dbo.TB_LINEAS.VCHA_LIN_LINEA_ID = dbo.TB_ARTICULOS.VCHA_LIN_LINEA_ID LEFT OUTER JOIN dbo.TB_DISEÑOS ON dbo.TB_ARTICULOS.VCHA_DIS_DISEÑO_ID = dbo.TB_DISEÑOS.VCHA_DIS_DISEÑO_ID LEFT OUTER JOIN dbo.TB_CATALOGOS ON dbo.TB_ARTICULOS.VCHA_ART_CATALOGO_VIGENTE = dbo.TB_CATALOGOS.VCHA_CAT_CATALOGO_ID WHERE (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'CC_2') AND DTIM_EMO_FECHA >= CONVERT(DATETIME, '" + VAR_FECHA_INICIO + "', 102) AND "
         var_cadena = var_cadena + " DTIM_EMO_FECHA <  CONVERT(DATETIME, '" + VAR_FECHA_FIN + "', 102) and inte_emo_numero_origen = 2"
         rs1.Open var_cadena, cnn_cdindustrial, adOpenDynamic, adLockOptimistic
         While Not rs1.EOF
               var_dia = CStr(Day(CDate(rs1!DTIM_emo_FECHA)))
               var_mes = CStr(Month(CDate(rs1!DTIM_emo_FECHA)))
               var_año = CStr(Year(CDate(rs1!DTIM_emo_FECHA)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_factura = var_dia + "/" + var_mes + "/" + var_año
               var_dia = CStr(Day(CDate(rs1!DTIM_ART_FECHA_ALTA)))
               var_mes = CStr(Month(CDate(rs1!DTIM_ART_FECHA_ALTA)))
               var_año = CStr(Year(CDate(rs1!DTIM_ART_FECHA_ALTA)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_alta = var_dia + "/" + var_mes + "/" + var_año
               var_canal_venta = Trim(rs1!VCHA_CAN_CANAL_VENTA_ID)
               If var_canal_venta = "" Then
                  var_canal_venta = "SIN CANAL"
               End If
               var_nombre_canal = Trim(rs1!VCHA_CAN_NOMBRE)
               If var_nombre_canal = "" Then
                  var_nombre_canal = "SIN CANAL"
               End If
               var_agente = Trim(rs1!VCHA_AGE_AGENTE_ID)
               If var_agente = "" Then
                  var_agente = "SIN AGENTE"
               End If
               var_nombre_agente = Trim(rs1!VCHA_AGE_NOMBRE)
               If var_nombre_agente = "" Then
                  var_nombre_agente = "SIN AGENTE"
               End If
               var_ruta = Trim(rs1!VCHA_RUT_RUTA_ID)
               If var_ruta = "" Then
                  var_ruta = "SIN RUTA"
               End If
               var_nombre_ruta = Trim(rs1!VCHA_RUT_NOMBRE)
               If var_nombre_ruta = "" Then
                  var_nombre_ruta = "SIN RUTA"
               End If
               var_zona = Trim(rs1!VCHA_ZON_ZONA_ID)
               If var_zona = "" Then
                  var_zona = "SIN ZONA"
               End If
               var_nombre_zona = Trim(rs1!VCHA_ZON_NOMBRE)
               If var_nombre_zona = "" Then
                  var_nombre_zona = "SIN ZONA"
               End If
               var_cliente = Trim(rs1!VCHA_CLI_CLAVE_ID)
               If var_cliente = "" Then
                  var_cliente = "SIN CLIENTE"
               End If
               var_cliente_unfo = Trim(VCHA_CLI_CLAVE_UNIFICADA_ID)
               If var_cliente_unfo = "" Then
                  var_cliente_unfo = "0"
               End If
               var_nombre_cliente = Trim(rs1!VCHA_CLI_NOMBRE)
               If var_nombre_cliente = "" Then
                  var_nombre_cliente = "SIN CLIENTE"
               End If
               var_rfc = Trim(rs1!VCHA_CLI_RFC)
               var_titular = Trim(rs1!VCHA_TIT_TITULAR_ID)
               If var_titular = "" Then
                  var_titular = "SIN TITULAR"
               End If
               var_nombre_titular = Trim(rs1!VCHA_TIT_NOMBRE)
               If var_nombre_titular = "" Then
                  var_nombre_titular = "SIN TITULAR"
               End If
               var_grupo = Trim(rs1!VCHA_GAC_GRUPO_ACTUAL_ID)
               If var_grupo = "" Then
                  var_grupo = "SIN GRUPO"
               End If
               var_nombre_grupo = Trim(rs1!VCHA_GAC_NOMBRE)
               If var_nombre_grupo = "" Then
                  var_nombre_grupo = "SIN GRUPO"
               End If
               var_establecimiento = ""
               If var_establecimiento = "" Then
                  var_establecimiento = "SIN ESTABLECIMIENTO"
               End If
               var_nombre_establecimiento = ""
               If var_nombre_establecimiento = "" Then
                  var_nombre_establecimiento = "SIN ESTABLECIMIENTO"
               End If
               var_cp = Trim(rs1!VCHA_CLI_CP)
               If var_cp = "" Then
                  var_cp = "SIN CP"
               End If
               var_estado = Trim(rs1!VCHA_EST_ESTADO_ID)
               If var_estado = "" Then
                  var_estado = "SIN ESTADO"
               End If
               var_nombre_estado = Trim(rs1!VCHA_EST_NOMBRE)
               If var_nombre_estado = "" Then
                  var_nombre_estado = "SIN ESTADO"
               End If
               var_ciudad = Trim(rs1!VCHA_CIU_CIUDAD_ID)
               If var_ciudad = "" Then
                  var_ciudad = "SIN CIUDAD"
               End If
               var_nombre_ciudad = Trim(rs1!VCHA_CIU_NOMBRE)
               If var_nombre_ciudad = "" Then
                  var_nombre_ciudad = "SIN CIUDAD"
               End If
               var_municipio = Trim(rs1!VCHA_MUN_MUNICIPIO_ID)
               If var_municipio = "" Then
                  var_municipio = "SIN MUNICIPIO"
               End If
               var_nombre_municipio = Trim(rs1!VCHA_MUN_NOMBRE)
               If var_nombre_municipio = "" Then
                  var_nombre_municipio = "SIN MUNICIPIO"
               End If
               var_colonia = Trim(rs1!vcha_col_colonia_id)
               If var_colonia = "" Then
                  var_colonia = "SIN COLONIA"
               End If
               var_nombre_colonia = Trim(rs1!VCHA_COL_NOMBRE)
               If var_nombre_colonia = "" Then
                  var_nombre_colonia = "SIN COLONIA"
               End If
               var_articulo = Trim(rs1!VCHA_ART_ARTICULO_ID)
               If var_articulo = "" Then
                  var_articulo = "SIN ARTICULO"
               End If
               var_nombre_articulo = Trim(rs1!VCHA_aRT_NOMBRE_ESPAÑOL)
               If var_nombre_articulo = "" Then
                  var_nombre_articulo = "SIN ARTICULO"
               End If
               var_catalogo = Trim(rs1!VCHA_CAT_CATALOGO_ID)
               If var_catalogo = "" Then
                  var_catalogo = "SIN CATALOGO"
               End If
               var_nombre_catalogo = Trim(rs1!VCHA_CAT_NOMBRE)
               If var_nombre_catalogo = "" Then
                  var_nombre_catalogo = "SIN CATALOGO"
               End If
               var_diseño = Trim(rs1!VCHA_DIS_DISEÑO_ID)
               If var_diseño = "" Then
                  var_diseño = "SIN DISEÑO"
               End If
               var_nombre_diseño = Trim(rs1!VCHA_DIS_NOMBRE)
               If var_nombre_diseño = "" Then
                  var_nombre_diseño = "SIN DISEÑO"
               End If
               var_linea = Trim(rs1!VCHA_LIN_LINEA_ID)
               If var_linea = "" Then
                  var_linea = "SIN LINEA"
               End If
               var_nombre_linea = Trim(rs1!VCHA_LIN_NOMBRE)
               If var_nombre_linea = "" Then
                  var_nombre_linea = "SIN LINEA"
               End If
               var_talla = Trim(rs1!vcha_Tal_talla_id)
               If var_talla = "" Then
                  var_talla = "SIN TALLA"
               End If
               var_nombre_talla = Trim(rs1!VCHA_TAL_NOMBRE)
               If var_nombre_talla = "" Then
                  var_nombre_talla = "SIN TALLA"
               End If
               var_licencia = Trim(rs1!vcha_lic_licencia_id)
               If var_licencia = "" Then
                  var_licencia = "SIN LICENCIA"
               End If
               var_nombre_licencia = Trim(rs1!vcha_lic_nombre)
               If var_nombre_licencia = "" Then
                  var_nombre_licencia = "SIN LICENCIA"
               End If
               var_numero_licencia = Trim(rs1!VCHA_ART_NUMERO_LIC)
               If var_numero_licencia = "" Then
                  var_numero_licencia = "SIN NUMERO DE LICENCIA"
               End If
               var_pais = Trim(rs1!VCHA_PAI_PAIS_ID)
               If var_pais = "" Then
                  var_pais = "SIN PAIS"
               End If
               var_nombre_pais = Trim(rs1!vcha_pai_nombre)
               If var_nombre_pais = "" Then
                  var_nombre_pais = "SIN PAIS"
               End If
        
         
               var_descuento_1 = IIf(IsNull(rs1!FLOA_SAL_DESCUENTO_1), 0, rs1!FLOA_SAL_DESCUENTO_1)
               var_descuento_2 = IIf(IsNull(rs1!FLOA_SAL_DESCUENTO_2), 0, rs1!FLOA_SAL_DESCUENTO_2)
               VAR_PRECIO = rs1!FLOA_SAL_PRECIO - (rs1!FLOA_SAL_PRECIO * (var_descuento_1 / 100))
               VAR_PRECIO = VAR_PRECIO - (VAR_PRECIO * (var_descuento_2 / 100))
               VAR_PORCENTAJE_IVA = IIf(IsNull(rs1!FLOA_CAR_PORCENTAJE_IVA), 0, rs1!FLOA_CAR_PORCENTAJE_IVA)
               var_precio_2 = VAR_PRECIO * ((1 + (VAR_PORCENTAJE_IVA / 100)))
               var_importe_iva = var_precio_2 - VAR_PRECIO
               VAR_PRECIO = var_precio_2
               var_precio_base = IIf(IsNull(rs1!MONE_ART_PRECIO_BASE), 0, rs1!MONE_ART_PRECIO_BASE) * ((1 + (VAR_PORCENTAJE_IVA / 100)))
                         
               var_cadena = "insert into  VT_TB_VENTAS  (VTA_INDICE_ID, VTA_ORIGEN_ID, VTA_EMPRESA_ID, VTA_EMPRESA, VTA_UNIDAD_ORGANIZACIONAL_ID, VTA_UNIDAD_ORGANIZACIONAL,VTA_TIENDA_ID,VTA_TIENDA,VTA_CANAL_ID, VTA_CANAL, VTA_REGION,VTA_AGENTE_ID,VTA_AGENTE,VTA_RUTA_ID,VTA_RUTA,VTA_ZONA_ID,VTA_ZONA,VTA_CLIENTE_ID,VTA_CLIENTE_ID_UNFO,VTA_CLIENTE,VTA_RFC_CLIENTE,VTA_TITULAR_ID,VTA_TITULAR,VTA_GRUPO_ID,VTA_GRUPO, "
               var_cadena = var_cadena + " VTA_ESTABLECIMIENTO_ID,VTA_ESTABLECIMIENTO,VTA_CODIGO_POSTAL,VTA_ESTADO_ID,VTA_ESTADO,VTA_CIUDAD_ID,VTA_CIUDAD,VTA_MUNICIPIO_ID,VTA_MUNICIPIO, VTA_COLONIA_ID,VTA_COLONIA,VTA_FECHA, VTA_SEMANA,VTA_ID_ENCABEZADO_SID,VTA_ID_DETALLE_SID,VTA_MOVIMIENTO_ID,VTA_TIPO_MOVIMIENTO_ID, VTA_DOCUMENTO, VTA_DESCRIPCION_DOCUMENTO, VTA_NUMERO_DOCUMENTO,VTA_SERIE,VTA_PLAZO,VTA_ARTICULO_ID,                 VTA_ARTICULO, "
               var_cadena = var_cadena + "VTA_CATALOGO_ID,VTA_CATALOGO,VTA_DISENIO_ID,VTA_DISENIO, VTA_FAMILIA_ID, VTA_FAMILIA, VTA_LINEA_ID,VTA_LINEA, VTA_TALLA_ID,VTA_TALLA,VTA_LICENCIA_ID,VTA_LICENCIA,VTA_NUMERO_LICENCIA, "
               var_cadena = var_cadena + "VTA_PRECIO_BASE, VTA_IMPUESTO_PORCIENTO ,VTA_FECHA_ALTA_CODIGO, VTA_CANTIDAD, VTA_COSTO, VTA_PRECIO_LISTA_A, VTA_PRECIO_LISTA_B, VTA_PRECIO_LISTA_C, VTA_DESCUENTO1 ,VTA_DESCUENTO2, VTA_DESCUENTO3, VTA_TIPO_CAMBIO, VTA_MONEDA ,VTA_IMPORTE,VTA_FLETE_ID,VTA_IMPORTE_FLETE,VTA_IMPUESTO_FLETE,VTA_MOVIMIENTO_INDEX,VTA_TOTAL_MOVIMIENTOS,VTA_CLAVE_REGISTRO, VTA_PAIS_ID, VTA_PAIS, VTA_IMPUESTO, precio) VALUES"
               var_cadena = var_cadena + "(sq_ventas_id.nextval," + CStr(var_origen) + ",'" + Trim(rs1!VCHA_EMP_EMPRESA_ID) + "', '" + Trim(rs1!VCHA_EMP_NOMBRE) + "','" + Trim(rs1!VCHA_UOR_UNIDAD_ID) + "','" + Trim(rs1!VCHA_UOR_NOMBRE) + "','" + Trim(rs1!VCHA_ALM_ALMACEN_ID) + "', '" + Trim(rs1!VCHA_ALM_NOMBRE) + "','" + Trim(var_canal_venta) + "','" + Trim(var_nombre_canal) + "','','" + Trim(var_agente) + "', '" + Trim(var_nombre_agente) + "','" + Trim(var_ruta) + "','" + Trim(var_nombre_ruta) + "', '" + Trim(var_zona) + "', '" + Trim(var_nombre_zona) + "', '" + Trim(var_cliente) + "','" + Trim(VCHA_CLI_CLAVE_UNIFICADA_ID) + "','" + Trim(var_nombre_cliente) + "', '" + Trim(var_rfc) + "','" + Trim(var_titular) + "', '" + Trim(var_nombre_titular) + "', '" + Trim(var_grupo) + "', '" + Trim(var_nombre_grupo) + "',"
               var_cadena = var_cadena + "'" + Trim(var_establecimiento) + "', '" + Trim(var_nombre_establecimiento) + "', '" + Trim(var_cp) + "','" + Trim(var_estado) + "', '" + Trim(var_nombre_estado) + "', '" + Trim(var_ciudad) + "','" + Trim(var_nombre_ciudad) + "', '" + Trim(var_municipio) + "','" + Trim(var_nombre_municipio) + "','" + Trim(var_colonia) + "', '" + Trim(var_nombre_colonia) + "', TO_DATE('" + var_fecha_factura + "','DD/MM/YYYY')," + CStr(rs1!SEMANA) + ", '" + CStr(rs1!inte_Car_numero) + Trim("sqlquezada2") + Trim("sicantia") + "', '" + CStr(rs1!inte_sal_consecutivo_tabla) + Trim("sqlquezada2") + Trim("sidcantia") + "', '" + Trim(rs1!VCHA_MOV_MOVIMIENTO_ID) + "',  'FA',          'FA',         'FA'                 , " + CStr(rs1!inte_Car_numero) + ",'TC',0,'" + Trim(rs1!VCHA_ART_ARTICULO_ID) + "', '" + Trim(rs1!VCHA_aRT_NOMBRE_ESPAÑOL) + "',"
               var_cadena = var_cadena + "'" + Trim(var_catalogo) + "','" + Trim(var_nombre_catalogo) + "', '" + Trim(var_diseño) + "', '" + Trim(var_nombre_diseño) + "','','','" + Trim(var_linea) + "','" + Trim(var_nombre_linea) + "','" + Trim(var_talla) + "','" + Trim(var_nombre_talla) + "','" + Trim(var_licencia) + "', '" + Trim(var_nombre_licencia) + "','" + Trim(var_numero_licencia) + "',"
               var_cadena = var_cadena + CStr(var_precio_base) + "," + CStr(rs1!FLOA_CAR_PORCENTAJE_IVA) + ",TO_DATE('" + var_fecha_alta + "','DD/MM/YYYY')," + CStr(rs1!floa_sal_cantidad) + "," + CStr(IIf(IsNull(rs1!FLOA_SAL_COSTO), 0, rs1!FLOA_SAL_COSTO)) + "," + CStr(VAR_PRECIO) + ",0,0," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(1) + ",'" + Trim(rs1!vcha_mon_divisa) + "'," + CStr(rs1!floa_sal_cantidad * VAR_PRECIO) + ",0,0,0,'','','" + Trim(rs1!VCHA_EMP_EMPRESA_ID) + Trim(rs1!VCHA_UOR_UNIDAD_ID) + "FA" + Trim("TC") + CStr(rs1!inte_Car_numero) + "_" + CStr(rs1!inte_sal_consecutivo_tabla) + "','" + Trim(var_pais) + "','" + Trim(var_nombre_pais) + "', " + CStr(rs1!floa_sal_cantidad * var_importe_iva) + "," + CStr(rs1!FLOA_SAL_PRECIO) + ")"
               
            
               rs2.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
               
               rs3.Open "SELECT * FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS WHERE  VCHA_EMP_EMPRESA_ID = '" + rs1!VCHA_EMP_EMPRESA_ID + "' AND VCHA_CAR_TIPO_DOCUMENTO = 'FA' AND VCHA_CAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = 'TC' AND VCHA_CLI_CLAVE_ID = '" + rs1!VCHA_CLI_CLAVE_ID + "' AND INTE_CAR_NUMERO = " + CStr(rs1!inte_Car_numero), cnn_cdindustrial, adOpenDynamic, adLockOptimistic
               If rs3.EOF Then
                  rs4.Open "INSERT INTO TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS (VCHA_EMP_EMPRESA_ID,  VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO) VALUES ('" + rs1!VCHA_EMP_EMPRESA_ID + "','FA','FA','TC','" + rs1!VCHA_CLI_CLAVE_ID + "'," + CStr(rs1!inte_Car_numero) + ")", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
               End If
               rs3.Close
            
               rs1.MoveNext
               var_i = var_i + 1
               Text1 = var_i
               Me.Refresh
               Me.Text1.Refresh
               'Me.lbl_accion.Refresh
         Wend
         rs1.Close
      
         rs2.Open "select distinct VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO FROM TB_TEMP_SUBIR_INFORMACION_VENTAS_FACTURAS", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
         var_i = 0
         While Not rs2.EOF
               rs4.Open "UPDATE TB_ENCABEZADO_MOVIMIENTOS SET INTE_EMO_NUMERO_ORIGEN = 1 WHeRE vcha_emp_empresa_id = '" + rs2!VCHA_EMP_EMPRESA_ID + "' and vcha_mov_movimiento_id = 'CC_2' and inte_emo_numero = " + CStr(rs2!inte_Car_numero), cnn_cdindustrial, adOpenDynamic, adLockOptimistic
               rs2.MoveNext
               var_i = var_i + 1
               Text1.Text = var_id
               Me.Refresh
               Me.Text1.Refresh
               Me.lbl_accion.Refresh
         Wend
         rs2.Close
          'devoluciones
          Call cantia_devoluciones
          var_origen = 6
      End If
      Next var_j
   MsgBox "Termino la carga de los movimientos", vbOKOnly, "ATENCION"
   'End
   
End Sub

Private Sub cmd_subir_notas_credito_Click()
   Dim var_canal_venta As String
   Dim var_nombre_canal As String
   Dim var_agente As String
   Dim var_nombre_agente As String
   Dim var_ruta As String
   Dim var_nombre_ruta As String
   Dim var_zona As String
   Dim var_nombre_zona As String
   Dim var_cliente As String
   Dim var_cliente_unfo As String
   Dim var_nombre_cliente As String
   Dim var_rfc As String
   Dim var_titular As String
   Dim var_nombre_titular As String
   Dim var_grupo As String
   Dim var_nombre_grupo As String
   Dim var_establecimiento As String
   Dim var_nombre_establecimiento As String
   Dim var_cp As String
   Dim var_estado As String
   Dim var_nombre_estado As String
   Dim var_ciudad As String
   Dim var_nombre_ciudad As String
   Dim var_municipio As String
   Dim var_nombre_municipio As String
   Dim var_colonia As String
   Dim var_nombre_colonia As String
   Dim var_articulo As String
   Dim var_nombre_articulo As String
   Dim var_catalogo As String
   Dim var_nombre_catalogo As String
   Dim var_diseño As String
   Dim var_nombre_diseño As String
   Dim var_linea As String
   Dim var_nombre_linea As String
   Dim var_talla As String
   Dim var_nombre_talla As String
   Dim var_licencia As String
   Dim var_nombre_licencia As String
   Dim var_numero_licencia As String
   Dim var_pais As String
   Dim var_nombre_pais As String
   
   Dim var_origen As Integer
   Dim cnn_cdindustrial As ADODB.Connection
   Dim cnn_distribucion As ADODB.Connection
   Dim cnn_recuperacion As ADODB.Connection
   Dim cnn_cantia As ADODB.Connection
   Dim cnnoracle As ADODB.Connection
   Dim rs1 As ADODB.Recordset
   Dim rs2 As ADODB.Recordset
   Dim rs3 As ADODB.Recordset
   Dim rs4 As ADODB.Recordset
   Dim rs5 As ADODB.Recordset
   Dim rs6 As ADODB.Recordset
   Dim rs7 As ADODB.Recordset
   
   Set cnn_cdindustrial = CreateObject("ADODB.connection")
   Set cnn_distribucion = CreateObject("ADODB.connection")
   Set cnn_recuperacion = CreateObject("ADODB.connection")
   Set cnn_cantia = CreateObject("ADODB.connection")
   Set cnnoracle = CreateObject("ADODB.connection")
   Set rs1 = CreateObject("ADODB.recordset")
   Set rs2 = CreateObject("ADODB.recordset")
   Set rs3 = CreateObject("ADODB.recordset")
   Set rs4 = CreateObject("ADODB.recordset")
   Set rs5 = CreateObject("ADODB.recordset")
   Set rs6 = CreateObject("ADODB.recordset")
   Set rs7 = CreateObject("ADODB.recordset")
   
   If IsDate(txt_fecha_inicio) Then
      If IsDate(Me.txt_fecha_fin) Then
         cnn_cdindustrial.Open "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=vianney;Data Source=DISTRIBUCION"
         var_origen = 3
         cnnoracle.Open "Provider=OraOLEDB.Oracle.1;User ID=distribucion;Data Source=AP;Extended Properties=;Persist Security Info=True;Password=distribucion"
         cnn_cdindustrial.CommandTimeout = 360
         
         
         var_dia = CStr(Day(CDate(txt_fecha_inicio)))
         var_mes = CStr(Month(CDate(txt_fecha_inicio)))
         var_año = CStr(Year(CDate(txt_fecha_inicio)))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         VAR_FECHA_INICIO = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
         
         var_dia = CStr(Day(CDate(txt_fecha_fin)))
         var_mes = CStr(Month(CDate(txt_fecha_fin)))
         var_año = CStr(Year(CDate(txt_fecha_fin)))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         VAR_FECHA_FIN = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"

         
         'MsgBox cnn_cdindustrial
         rs1.Open "delete from tb_temp_subir_informacion_devoluciones", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
         rs1.Open "exec SP_SUBIR_INFORMACION_VENTAS_NOTAS_CREDITO " + VAR_FECHA_INICIO + "," + VAR_FECHA_FIN, cnn_cdindustrial, adOpenDynamic, adLockOptimistic
         rs1.Open "SELECT * FROM VW_SUBIR_INFORMACION_NOTAS_CREDITO", cnn_cdindustrial, adOpenDynamic, adLockOptimistic
         var_i = 0
         While Not rs1.EOF
               var_dia = CStr(Day(CDate(rs1!DTIM_TEM_fECHA)))
               var_mes = CStr(Month(CDate(rs1!DTIM_TEM_fECHA)))
               var_año = CStr(Year(CDate(rs1!DTIM_TEM_fECHA)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_factura = var_dia + "/" + var_mes + "/" + var_año
         
               var_dia = CStr(Day(CDate(rs1!DTIM_TEM_fECHA)))
               var_mes = CStr(Month(CDate(rs1!DTIM_TEM_fECHA)))
               var_año = CStr(Year(CDate(rs1!DTIM_TEM_fECHA)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               
               var_fecha_alta = var_dia + "/" + var_mes + "/" + var_año
               
         
               var_canal_venta = Trim(rs1!VCHA_CAN_CANAL_VENTA_ID)
               If var_canal_venta = "" Then
                  var_canal_venta = "SIN CANAL"
               End If
               var_nombre_canal = Trim(rs1!VCHA_CAN_NOMBRE)
               If var_nombre_canal = "" Then
                  var_nombre_canal = "SIN CANAL"
               End If
               var_agente = Trim(rs1!VCHA_AGE_AGENTE_ID)
               If var_agente = "" Then
                  var_agente = "SIN AGENTE"
               End If
               var_nombre_agente = Trim(rs1!VCHA_AGE_NOMBRE)
               If var_nombre_agente = "" Then
                  var_nombre_agente = "SIN AGENTE"
               End If
               var_ruta = Trim(rs1!VCHA_RUT_RUTA_ID)
               If var_ruta = "" Then
                  var_ruta = "SIN RUTA"
               End If
               var_nombre_ruta = Trim(rs1!VCHA_RUT_NOMBRE)
               If var_nombre_ruta = "" Then
                  var_nombre_ruta = "SIN RUTA"
               End If
               var_zona = IIf(IsNull(rs1!VCHA_ZON_ZONA_ID), "", rs1!VCHA_ZON_ZONA_ID)
               If var_zona = "" Then
                  var_zona = "SIN ZONA"
               End If
               var_nombre_zona = IIf(IsNull(rs1!VCHA_ZON_DESCRIPCION), "", rs1!VCHA_ZON_DESCRIPCION)
               If var_nombre_zona = "" Then
                  var_nombre_zona = "SIN ZONA"
               End If
               var_cliente = Trim(rs1!VCHA_CLI_CLAVE_ID)
               If var_cliente = "" Then
                  var_cliente = "SIN CLIENTE"
               End If
               var_cliente_unfo = Trim(VCHA_CLI_CLAVE_UNIFICADA_ID)
               If var_cliente_unfo = "" Then
                  var_cliente_unfo = "0"
               End If
               var_nombre_cliente = Trim(rs1!NOMBRE_CLIENTE)
               If var_nombre_cliente = "" Then
                  var_nombre_cliente = "SIN CLIENTE"
               End If
               var_rfc = Trim(IIf(IsNull(rs1!VCHA_CLI_RFC), "SIN RFC", rs1!VCHA_CLI_RFC))
               var_titular = Trim(rs1!VCHA_TIT_TITULAR_ID)
               If var_titular = "" Then
                  var_titular = "SIN TITULAR"
               End If
               var_nombre_titular = Trim(IIf(IsNull(rs1!NOMBRE_TITULAR), "", rs1!NOMBRE_TITULAR))
               If var_nombre_titular = "" Then
                  var_nombre_titular = "SIN TITULAR"
               End If
               var_grupo = Trim(IIf(IsNull(rs1!VCHA_GAC_GRUPO_ACTUAL_ID), "", rs1!VCHA_GAC_GRUPO_ACTUAL_ID))
               If var_grupo = "" Then
                  var_grupo = "SIN GRUPO"
               End If
               var_nombre_grupo = Trim(IIf(IsNull(rs1!NOMBRE_GRUPO), "", rs1!NOMBRE_GRUPO))
               If var_nombre_grupo = "" Then
                  var_nombre_grupo = "SIN GRUPO"
               End If
               var_establecimiento = "SIN ESTABLECIMIENTO"
               var_nombre_establecimiento = "SIN ESTABLECIMIENTO"
               var_cp = Trim(rs1!VCHA_CLI_CP)
               If var_cp = "" Then
                  var_cp = "SIN CP"
               End If
               var_estado = Trim(rs1!VCHA_EST_ESTADO_ID)
               If var_estado = "" Then
                  var_estado = "SIN ESTADO"
               End If
               var_nombre_estado = Trim(IIf(IsNull(rs1!VCHA_EST_NOMBRE), "", rs1!VCHA_EST_NOMBRE))
               If var_nombre_estado = "" Then
                  var_nombre_estado = "SIN ESTADO"
               End If
               var_ciudad = Trim(rs1!VCHA_CIU_CIUDAD_ID)
               If var_ciudad = "" Then
                  var_ciudad = "SIN CIUDAD"
               End If
               var_nombre_ciudad = Trim(IIf(IsNull(rs1!VCHA_CIU_NOMBRE), "", rs1!VCHA_CIU_NOMBRE))
               If var_nombre_ciudad = "" Then
                  var_nombre_ciudad = "SIN CIUDAD"
               End If
               var_municipio = Trim(rs1!VCHA_MUN_MUNICIPIO_ID)
               If var_municipio = "" Then
                  var_municipio = "SIN MUNICIPIO"
               End If
               var_nombre_municipio = Trim(IIf(IsNull(rs1!VCHA_MUN_NOMBRE), "", rs1!VCHA_MUN_NOMBRE))
               If var_nombre_municipio = "" Then
                  var_nombre_municipio = "SIN MUNICIPIO"
               End If
               var_colonia = IIf(IsNull(rs1!vcha_col_colonia_id), "", rs1!vcha_col_colonia_id)
               If var_colonia = "" Then
                  var_colonia = "SIN COLONIA"
               End If
               var_nombre_colonia = Trim(IIf(IsNull(rs1!VCHA_COL_NOMBRE), "", rs1!VCHA_COL_NOMBRE))
               If var_nombre_colonia = "" Then
                  var_nombre_colonia = "SIN COLONIA"
               End If
               var_articulo = Trim(rs1!VCHA_ART_ARTICULO_ID)
               If var_articulo = "" Then
                  var_articulo = "SIN ARTICULO"
               End If
               var_nombre_articulo = Trim(rs1!VCHA_aRT_NOMBRE_ESPAÑOL)
               If var_nombre_articulo = "" Then
                  var_nombre_articulo = "SIN ARTICULO"
               End If
               var_catalogo = Trim(IIf(IsNull(rs1!vcha_art_catalogo_id), "", rs1!vcha_art_catalogo_id))
               If var_catalogo = "" Then
                  var_catalogo = "SIN CATALOGO"
               End If
               var_nombre_catalogo = Trim(IIf(IsNull(rs1!VCHA_CAT_NOMBRE), "", rs1!VCHA_CAT_NOMBRE))
               If var_nombre_catalogo = "" Then
                  var_nombre_catalogo = "SIN CATALOGO"
               End If
               var_diseño = Trim(IIf(IsNull(rs1!VCHA_DIS_DISEÑO_ID), "", rs1!VCHA_DIS_DISEÑO_ID))
               If var_diseño = "" Then
                  var_diseño = "SIN DISEÑO"
               End If
               var_nombre_diseño = Trim(IIf(IsNull(rs1!VCHA_DIS_NOMBRE), "", rs1!VCHA_DIS_NOMBRE))
               If var_nombre_diseño = "" Then
                  var_nombre_diseño = "SIN DISEÑO"
               End If
               var_linea = Trim(IIf(IsNull(rs1!VCHA_LIN_LINEA_ID), "", rs1!VCHA_LIN_LINEA_ID))
               If var_linea = "" Then
                  var_linea = "SIN LINEA"
               End If
               var_nombre_linea = Trim(IIf(IsNull(rs1!VCHA_LIN_NOMBRE), "", rs1!VCHA_LIN_NOMBRE))
               If var_nombre_linea = "" Then
                  var_nombre_linea = "SIN LINEA"
               End If
               var_talla = Trim(IIf(IsNull(rs1!vcha_Tal_talla_id), "", rs1!vcha_Tal_talla_id))
               If var_talla = "" Then
                  var_talla = "SIN TALLA"
               End If
               var_nombre_talla = Trim(IIf(IsNull(rs1!VCHA_TAL_NOMBRE), "", rs1!VCHA_TAL_NOMBRE))
               If var_nombre_talla = "" Then
                  var_nombre_talla = "SIN TALLA"
               End If
               var_licencia = Trim(IIf(IsNull(rs1!vcha_lic_licencia_id), "", rs1!vcha_lic_licencia_id))
               If var_licencia = "" Then
                  var_licencia = "SIN LICENCIA"
               End If
               var_nombre_licencia = Trim(IIf(IsNull(rs1!vcha_lic_nombre), "", rs1!vcha_lic_nombre))
               If var_nombre_licencia = "" Then
                  var_nombre_licencia = "SIN LICENCIA"
               End If
               var_numero_licencia = Trim(IIf(IsNull(rs1!vcha_lic_numero), "", rs1!vcha_lic_numero))
               If var_numero_licencia = "" Then
                  var_numero_licencia = "SIN NUMERO DE LICENCIA"
               End If
               var_pais = Trim(rs1!VCHA_PAI_PAIS_ID)
               If var_pais = "" Then
                  var_pais = "SIN PAIS"
               End If
               var_nombre_pais = Trim(IIf(IsNull(rs1!vcha_pai_nombre), "", rs1!vcha_pai_nombre))
               If var_nombre_pais = "" Then
                  var_nombre_pais = "SIN PAIS"
               End If
               
         
               var_descuento_1 = 0
               var_descuento_2 = 0
               VAR_PRECIO = rs1!floa_TEM_precio
               If VAR_EMPRESA = "03" Or VAR_EMPRESA = "28" Then
                  VAR_PORCENTAJE_IVA = 0
               Else
                  If Year(rs1!DTIM_TEM_fECHA) < 2010 Then
                     VAR_PORCENTAJE_IVA = 15
                  Else
                     VAR_PORCENTAJE_IVA = 16
                  End If
               End If
               var_precio_2 = VAR_PRECIO / ((1 + (VAR_PORCENTAJE_IVA / 100)))
               var_importe_iva = VAR_PRECIO - var_precio_2
               var_precio_base = VAR_PRECIO
               var_tipo_cambio = 1
               rs2.Open "select sq_notascredito_id.nextval from TB_sECUENCIA", cnnoracle, adOpenDynamic, adLockOptimistic
               If Not rs2.EOF Then
                  VAR_SECUENCIA = CStr(rs2(0).Value)
               End If
               rs2.Close
               'MsgBox VAR_SECUENCIA
               var_cadena = "insert into  VT_TB_VENTAS  (VTA_INDICE_ID, VTA_ORIGEN_ID, VTA_EMPRESA_ID, VTA_EMPRESA, VTA_UNIDAD_ORGANIZACIONAL_ID, VTA_UNIDAD_ORGANIZACIONAL,VTA_TIENDA_ID,VTA_TIENDA,VTA_CANAL_ID, VTA_CANAL, VTA_REGION,VTA_AGENTE_ID,VTA_AGENTE,VTA_RUTA_ID,VTA_RUTA,VTA_ZONA_ID,VTA_ZONA,VTA_CLIENTE_ID,VTA_CLIENTE_ID_UNFO,VTA_CLIENTE,VTA_RFC_CLIENTE,VTA_TITULAR_ID,VTA_TITULAR,VTA_GRUPO_ID,VTA_GRUPO, "
               var_cadena = var_cadena + " VTA_ESTABLECIMIENTO_ID,VTA_ESTABLECIMIENTO,VTA_CODIGO_POSTAL,VTA_ESTADO_ID,VTA_ESTADO,VTA_CIUDAD_ID,VTA_CIUDAD,VTA_MUNICIPIO_ID,VTA_MUNICIPIO, VTA_COLONIA_ID,VTA_COLONIA,VTA_FECHA, VTA_SEMANA,VTA_ID_ENCABEZADO_SID,VTA_ID_DETALLE_SID,VTA_MOVIMIENTO_ID,VTA_TIPO_MOVIMIENTO_ID, VTA_DOCUMENTO, VTA_DESCRIPCION_DOCUMENTO, VTA_NUMERO_DOCUMENTO,VTA_SERIE,VTA_PLAZO,VTA_ARTICULO_ID,                 VTA_ARTICULO, "
               var_cadena = var_cadena + "VTA_CATALOGO_ID,VTA_CATALOGO,VTA_DISENIO_ID,VTA_DISENIO, VTA_FAMILIA_ID, VTA_FAMILIA, VTA_LINEA_ID,VTA_LINEA, VTA_TALLA_ID,VTA_TALLA,VTA_LICENCIA_ID,VTA_LICENCIA,VTA_NUMERO_LICENCIA, "
               var_cadena = var_cadena + "VTA_PRECIO_BASE, VTA_IMPUESTO_PORCIENTO ,VTA_FECHA_ALTA_CODIGO, VTA_CANTIDAD, VTA_COSTO, VTA_PRECIO_LISTA_A, VTA_PRECIO_LISTA_B, VTA_PRECIO_LISTA_C, VTA_DESCUENTO1 ,VTA_DESCUENTO2, VTA_DESCUENTO3, VTA_TIPO_CAMBIO, VTA_MONEDA ,VTA_IMPORTE,VTA_FLETE_ID,VTA_IMPORTE_FLETE,VTA_IMPUESTO_FLETE,VTA_MOVIMIENTO_INDEX,VTA_TOTAL_MOVIMIENTOS,VTA_CLAVE_REGISTRO, VTA_PAIS_ID, VTA_PAIS, VTA_IMPUESTO, precio) VALUES"
               var_cadena = var_cadena + "(sq_ventas_id.nextval," + CStr(var_origen) + ",'" + Trim(rs1!VCHA_EMP_EMPRESA_ID) + "', '" + Trim(rs1!VCHA_EMP_NOMBRE) + "','SIN_UNIDADORG','SIN UNIDADORG','SA', 'SA','" + Trim(var_canal_venta) + "','" + Trim(var_nombre_canal) + "','','" + Trim(var_agente) + "', '" + Trim(var_nombre_agente) + "','" + Trim(var_ruta) + "','" + Trim(var_nombre_ruta) + "', '" + Trim(var_zona) + "', '" + Trim(var_nombre_zona) + "', '" + Trim(var_cliente) + "','" + Trim(VCHA_CLI_CLAVE_UNIFICADA_ID) + "','" + Trim(var_nombre_cliente) + "', '" + Trim(var_rfc) + "','" + Trim(var_titular) + "', '" + Trim(var_nombre_titular) + "', '" + Trim(var_grupo) + "', '" + Trim(Mid(var_nombre_grupo, 1, 50)) + "',"
               var_cadena = var_cadena + "'" + Trim(var_establecimiento) + "', '" + Trim(var_nombre_establecimiento) + "', '" + Trim(var_cp) + "','" + Trim(var_estado) + "', '" + Trim(var_nombre_estado) + "', '" + Trim(var_ciudad) + "','" + Trim(var_nombre_ciudad) + "', '" + Trim(var_municipio) + "','" + Trim(var_nombre_municipio) + "','" + Trim(var_colonia) + "', '" + Trim(var_nombre_colonia) + "', TO_DATE('" + var_fecha_factura + "','DD/MM/YYYY')," + CStr(rs1!SEMANA) + ", '" + VAR_EMPRESA + "NC" + "', '" + VAR_EMPRESA + "NC" + "', 'NC',  'NC',          'NC',         'NC'                 , " + CStr(VAR_SECUENCIA) + ",'',0,'" + Trim(rs1!VCHA_ART_ARTICULO_ID) + "', '" + Trim(rs1!VCHA_aRT_NOMBRE_ESPAÑOL) + "',"
               var_cadena = var_cadena + "'" + Trim(var_catalogo) + "','" + Trim(var_nombre_catalogo) + "', '" + Trim(var_diseño) + "', '" + Trim(var_nombre_diseño) + "','','','" + Trim(var_linea) + "','" + Trim(var_nombre_linea) + "','" + Trim(var_talla) + "','" + Trim(var_nombre_talla) + "','" + Trim(var_licencia) + "', '" + Trim(var_nombre_licencia) + "','" + Trim(var_numero_licencia) + "',"
               var_cadena = var_cadena + CStr(var_precio_base) + "," + CStr(VAR_PORCENTAJE_IVA) + ",TO_DATE('" + var_fecha_alta + "','DD/MM/YYYY')," + CStr(rs1!floa_TEM_cantidad) + "," + CStr(VAR_PRECIO) + "," + CStr(VAR_PRECIO) + ",0,0,0,0,0," + CStr(var_tipo_cambio) + ",'" + Trim(rs1!vcha_mon_divisa) + "'," + CStr(rs1!floa_TEM_cantidad * VAR_PRECIO) + ",0,0,0,'','','" + Trim(rs1!VCHA_EMP_EMPRESA_ID) + "NC" + CStr(VAR_SECUENCIA) + "','" + Trim(var_pais) + "','" + Trim(var_nombre_pais) + "', " + CStr(rs1!floa_TEM_cantidad * var_importe_iva) + "," + CStr(rs1!floa_TEM_precio) + ")"
         
               rs2.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
               rs1.MoveNext
               var_i = var_i + 1
               Text1 = var_i
               Me.Refresh
               Me.Text1.Refresh
         Wend
         rs1.Close
         MsgBox "Se a terminado de cargar las devoluciones", vbOKOnly, "ATENCION"
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha inicial incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   Set cnn_cdindustrial = CreateObject("ADODB.connection")
   Set cnn_distribucion = CreateObject("ADODB.connection")
   Set cnn_recuperacion = CreateObject("ADODB.connection")
   Set cnn_cantia = CreateObject("ADODB.connection")
   Set cnnoracle = CreateObject("ADODB.connection")
   Set rs1 = CreateObject("ADODB.recordset")
   Set rs2 = CreateObject("ADODB.recordset")
   Set rs3 = CreateObject("ADODB.recordset")
   Set rs4 = CreateObject("ADODB.recordset")
   Set rs5 = CreateObject("ADODB.recordset")
   Set rs6 = CreateObject("ADODB.recordset")
   Set rs7 = CreateObject("ADODB.recordset")
   
   'Me.opt_cdindustrial.Value = True
   'Me.opt_distribucion.Value = True
   
   Me.txt_fecha_inicio = Date - 1
   Me.txt_fecha_fin = Date + 1
   'Call cmd_subir_informacion_Click
   
End Sub
