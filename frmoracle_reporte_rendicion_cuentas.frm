VERSION 5.00
Begin VB.Form frmoracle_reporte_rendicion_cuentas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rendición de cuentas"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   " Tipo transporte "
      Height          =   750
      Left            =   45
      TabIndex        =   9
      Top             =   1275
      Width           =   3840
      Begin VB.OptionButton opt_ambos 
         Caption         =   "Ambos"
         Height          =   255
         Left            =   165
         TabIndex        =   12
         Top             =   315
         Width           =   900
      End
      Begin VB.OptionButton opt_externo 
         Caption         =   "Externo"
         Height          =   255
         Left            =   2835
         TabIndex        =   11
         Top             =   315
         Width           =   900
      End
      Begin VB.OptionButton opt_propio 
         Caption         =   "Propio"
         Height          =   255
         Left            =   1500
         TabIndex        =   10
         Top             =   315
         Width           =   900
      End
   End
   Begin VB.TextBox txt_embarque 
      Height          =   330
      Left            =   1110
      TabIndex        =   6
      Top             =   30
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3525
      Picture         =   "frmoracle_reporte_rendicion_cuentas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmoracle_reporte_rendicion_cuentas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   30
      TabIndex        =   0
      Top             =   405
      Width           =   3885
      Begin VB.TextBox txt_fin 
         Alignment       =   2  'Center
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
         Left            =   2595
         TabIndex        =   7
         Top             =   270
         Width           =   1215
      End
      Begin VB.TextBox txt_inicio 
         Alignment       =   2  'Center
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
         Left            =   615
         TabIndex        =   1
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2280
         TabIndex        =   8
         Top             =   353
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   353
         Width           =   420
      End
   End
   Begin VB.Frame x 
      Height          =   120
      Left            =   0
      TabIndex        =   5
      Top             =   285
      Width           =   3885
   End
End
Attribute VB_Name = "frmoracle_reporte_rendicion_cuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub cmd_imprimir_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
      var_dia_s = CStr(Day(CDate(Me.txt_inicio)))
      If Len(var_dia_s) = 1 Then
         var_dia_s = "0" + CStr(var_dia_s)
      End If
      var_mes_s = CStr(Month(CDate(Me.txt_inicio)))
      If Len(var_mes_s) = 1 Then
         var_mes_s = "0" + CStr(var_mes_s)
      End If
      var_año_s = CStr(Year(CDate(Me.txt_inicio)))
      If Len(var_año_s) = 1 Then
         var_año_s = "200" + CStr(var_año_s)
      End If
      If Len(var_año_s) = 2 Then
         var_año_s = "20" + CStr(var_año_s)
      End If
      If Len(var_año_s) = 3 Then
         var_año_s = "2" + CStr(var_año_s)
      End If
      var_fecha_inicio = var_dia_s + "/" + var_mes_s + "/" + var_año_s
      var_fecha_inicio_2 = "{d '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "'}"
      var_dia_s = CStr(Day(CDate(Me.txt_fin) + 1))
      If Len(var_dia_s) = 1 Then
         var_dia_s = "0" + CStr(var_dia_s)
      End If
      var_mes_s = CStr(Month(CDate(Me.txt_fin) + 1))
      If Len(var_mes_s) = 1 Then
         var_mes_s = "0" + CStr(var_mes_s)
      End If
      var_año_s = CStr(Year(CDate(Me.txt_fin) + 1))
      If Len(var_año_s) = 1 Then
         var_año_s = "200" + CStr(var_año_s)
      End If
      If Len(var_año_s) = 2 Then
         var_año_s = "20" + CStr(var_año_s)
      End If
      If Len(var_año_s) = 3 Then
         var_año_s = "2" + CStr(var_año_s)
      End If
      var_fecha_fin = var_dia_s + "/" + var_mes_s + "/" + var_año_s
      var_dia_s = CStr(Day(CDate(Me.txt_fin)))
      If Len(var_dia_s) = 1 Then
         var_dia_s = "0" + CStr(var_dia_s)
      End If
      var_mes_s = CStr(Month(CDate(Me.txt_fin)))
      If Len(var_mes_s) = 1 Then
         var_mes_s = "0" + CStr(var_mes_s)
      End If
      var_año_s = CStr(Year(CDate(Me.txt_fin)))
      If Len(var_año_s) = 1 Then
         var_año_s = "200" + CStr(var_año_s)
      End If
      If Len(var_año_s) = 2 Then
         var_año_s = "20" + CStr(var_año_s)
      End If
      If Len(var_año_s) = 3 Then
         var_año_s = "2" + CStr(var_año_s)
      End If
      
      var_fecha_finaL_2 = "{d '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "'}"
      strconsulta = "select * from xxvia_Tb_encabezado_embarques  where fecha_fin >= to_date(?,'DD/MM/YYYY') and fecha_fin < to_date(?,'DD/MM/YYYY') order by embarque desc"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 10, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 10, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux15 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing

      If Not rsaux15.EOF Then
         cnn.BeginTrans
         rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_consecutivo = IIf(IsNull(rs!numero), 0, rs!numero)
         Else
            var_consecutivo = 0
         End If
         var_consecutivo = var_consecutivo + 1
         rs.Close
         rs.Open "insert into TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
         While Not rsaux15.EOF
               Me.txt_embarque = rsaux15!Embarque
'--------

               strconsulta = "SELECT INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER, MAX(CUSTOMER_NAME) AS CLIENTE, MAX(ENTREGA) AS ESTABLECIMIENTO, SUM(FLOA_SAL_CANTIDAD_LEIDA) AS CANTIDAD FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = ?  GROUP BY INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER ORDER BY INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               strconsulta = "SELECT TRANSPORTE, TO_CHAR(FECHA_FIN, 'DD/MM/YYYY HH24:MI:SS') AS FECHA_FIN, TO_CHAR(FECHA_INICIO, 'DD/MM/YYYY HH24:MI:SS') FECHA_INICIO, chofer FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
                    .Parameters.Append parametro
               End With
               Set rsaux11 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               
               rsaux6.Open "select * from tb_oracle_transportes where clave = '" + IIf(IsNull(rsaux11!transporte), "", rsaux11!transporte) + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux6.EOF Then
                  var_numero_unidad = IIf(IsNull(rsaux6!nombre), "", rsaux6!nombre)
               Else
                  var_numero_unidad = ""
               End If
               rsaux6.Close
               
               
               var_chofer = IIf(IsNull(rsaux11!chofer), "", rsaux11!chofer)
               
               If var_chofer = "" Then
                  VAR_NOMBRE_CHOFER = "                                                     "
               Else
                  rsaux6.Open "select * from tb_choferes where vcha_cho_chofer_id = '" + var_chofer + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux6.EOF Then
                     VAR_NOMBRE_CHOFER = IIf(IsNull(rsaux6!vcha_cho_nombre), "                                              ", rsaux6!vcha_cho_nombre)
                     var_tipo_transporte = IIf(IsNull(rsaux6!inte_cho_tipo), 0, rsaux6!inte_cho_tipo)
                  Else
                     VAR_NOMBRE_CHOFER = "                                             "
                     var_tipo_transporte = 0
                  End If
                  rsaux6.Close
               End If
               rsaux6.Open "select * from tb_usuarios where vcha_usu_usuario_id = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux6.EOF Then
                  var_usuario_facturacion = IIf(IsNull(rsaux6!VCHA_USU_NOMBRE), "", rsaux6!VCHA_USU_NOMBRE) + " " + IIf(IsNull(rsaux6!VCHA_USU_APELLIDOS), "", rsaux6!VCHA_USU_APELLIDOS)
               Else
                  var_usuario_facturacion = ""
               End If
               rsaux6.Close
               If Not rsaux9.EOF Then
                  var_fecha_embarque = IIf(IsNull(rsaux11!FECHA_FIN), rsaux11!FECHA_INiCIO, rsaux11!FECHA_FIN)
                  rsaux5.Open "SELECT * FROM TB_ORACLE_TRANSPORTES where clave = '" + IIf(IsNull(rsaux11!transporte), "", rsaux11!transporte) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux5.EOF Then
                     var_transporte = IIf(IsNull(rsaux5!nombre), "", rsaux5!nombre)
                  Else
                     var_transporte = ""
                  End If
                  rsaux5.Close
                  var_cadena_sellos = ""
                  rsaux5.Open "select * from tb_sellos where inte_emb_embarque = " + CStr(CDbl(Me.txt_embarque)), cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux5.EOF
                        If var_cadena_sellos = "" Then
                           var_cadena_sellos = IIf(IsNull(rsaux5!vcha_sel_Sello), "", rsaux5!vcha_sel_Sello)
                        Else
                           var_cadena_sellos = var_cadena_sellos + ", " + IIf(IsNull(rsaux5!vcha_sel_Sello), "", rsaux5!vcha_sel_Sello)
                        End If
                        rsaux5.MoveNext
                  Wend
                  rsaux5.Close
               
                
                  strconsulta = "SELECT DISTINCT  J.SALESREP_ID, J.NAME  FROM OE_ORDER_HEADERS_ALL OHA, XXVIA_TB_SALIDAS_CAJAS, XXVIA_VENDEDORES J WHERE  INTE_EMB_EMBARQUE = ? AND OHA.ORDER_NUMBER = SOURCE_HEADER_NUMBER AND OHA.SALESREP_ID = J.SALESREP_ID AND J.SALESREP_ID <> -3"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
                       .Parameters.Append parametro
                  End With
                  Set rsaux5 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
               
                  VAR_CADENA_RUTAS = ""
                  While Not rsaux5.EOF
                        If VAR_CADENA_RUTAS = "" Then
                           VAR_CADENA_RUTAS = IIf(IsNull(rsaux5!Name), "", rsaux5!Name)
                        Else
                           VAR_CADENA_RUTAS = VAR_CADENA_RUTAS + ", " + IIf(IsNull(rsaux5!Name), "", rsaux5!Name)
                        End If
                        rsaux5.MoveNext
                  Wend
                  rsaux5.Close
                  While Not rsaux9.EOF
                        rsaux7.Open "select * from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES WHERE PEDIDO = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux7.EOF Then
                           var_orden = IIf(IsNull(rsaux7!orden_pedido), 0, rsaux7!orden_pedido)
                        Else
                           var_orden = 0
                        End If
                        rsaux7.Close
               
                        strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                              Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rsaux9!source_header_number)
                             .Parameters.Append parametro
                        End With
                        Set rsaux8 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If rsaux8!order_type_id = 1002 Then
                           var_nombre_cliente = IIf(IsNull(rsaux9!Cliente), "", rsaux9!Cliente)
                           If var_nombre_cliente = "VIANNEY TEXTIL HOGAR SA DE CV" Then
                              var_nombre_cliente = IIf(IsNull(rsaux9!ESTABLECIMIENTO), var_nombre_cliente, rsaux9!ESTABLECIMIENTO)
                           End If
                           var_tipo = "Nota de envío"
                           var_folio = rsaux9!source_header_number
                           
                           strconsulta = "SELECT DISTINCT INTE_PAQ_CAJA FROM XXVIA_TB_sALIDAS_CAJAS WHERE SOURCE_HEADER_NUMBER = ? AND FLOA_SAL_cANTIDAD_LEIDA > 0"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux9!source_header_number)
                                .Parameters.Append parametro
                           End With
                           Set rsaux10 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           VAR_CANTIDAD_BULTOS = 0
                           While Not rsaux10.EOF
                                 VAR_CANTIDAD_BULTOS = VAR_CANTIDAD_BULTOS + 1
                                 rsaux10.MoveNext
                           Wend
                           rsaux10.Close
                          
                           strconsulta = "select sum(floa_sal_Cantidad_leida) from xxvia_tb_salidas_cajas where source_header_number = ?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux9!source_header_number)
                                .Parameters.Append parametro
                           End With
                           Set rsaux10 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                        
                           If Not rsaux10.EOF Then
                              'var_cantidad_leida = var_cantidad_leida + IIf(IsNull(rsaux10(0).Value), 0, rsaux10(0).Value)
                              var_cantidad_leida = IIf(IsNull(rsaux10(0).Value), 0, rsaux10(0).Value)
                           Else
                              var_cantidad_leida = var_cantidad_leida + 0
                           End If
                           rsaux10.Close
'-----------   -
                           strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                                .Parameters.Append parametro
                           End With
                           Set rsaux7 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           
                           If Not rsaux7.EOF Then
                              var_clave_cliente = rsaux7!attribute1
                              'var_nombre_cliente = rsaux7!Description
                           Else
                              var_clave_cliente = ""
                              'var_nombre_cliente = ""
                           End If
                           rsaux7.Close
   
   
   
   
   
'--   ----------
                           
                           var_tipo_transporte = CStr(IIf(IsNull(var_tipo_transporte), "", var_tipo_transporte))
                           'var_cadena = "INSERT INTO TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS (INTE_TEM_CONSECUTIVO, EMBARQUE, CLIENTE, CANTIDAD, PEDIDO, unidad, sellos, FECHA_EMBARQUE, RUTA, direccion_entrega, BULTOS, TIPO, folio, clave_cliente, ORDEN_ENTREGA, chofer, encargado_facturacion, tipo_transporte, NUMERO_ECONOMICO) VALUES (" + CStr(var_consecutivo) + "," + CStr(rsaux9!inte_emb_embarque) + ",'" + var_nombre_cliente + "'," + CStr(var_cantidad_leida) + "," + CStr(rsaux9!source_header_number) + ",'" + var_transporte + "','" + var_cadena_sellos + "', '" + var_fecha_embarque + "','" + VAR_CADENA_RUTAS + "',''," + CStr(VAR_CANTIDAD_BULTOS) + ",'" + var_tipo + "'," + CStr(var_folio) + ",'" + var_clave_cliente + "'," + CStr(var_orden) + ",'" + VAR_NOMBRE_CHOFER + "','" + var_usuario_facturacion + "'," + CStr(var_tipo_transporte) + ",'" + var_numero_unidad + "')"
                           'MsgBox var_cadena
                           rs.Open "INSERT INTO TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS (INTE_TEM_CONSECUTIVO, EMBARQUE, CLIENTE, CANTIDAD, PEDIDO, unidad, sellos, FECHA_EMBARQUE, RUTA, direccion_entrega, BULTOS, TIPO, folio, clave_cliente, ORDEN_ENTREGA, chofer, encargado_facturacion, tipo_transporte, NUMERO_ECONOMICO) VALUES (" + CStr(var_consecutivo) + "," + CStr(rsaux9!inte_emb_embarque) + ",'" + var_nombre_cliente + "'," + CStr(var_cantidad_leida) + "," + CStr(rsaux9!source_header_number) + ",'" + var_transporte + "','" + var_cadena_sellos + "', '" + var_fecha_embarque + "','" + VAR_CADENA_RUTAS + "',''," + CStr(VAR_CANTIDAD_BULTOS) + ",'" + var_tipo + "'," + CStr(var_folio) + ",'" + var_clave_cliente + "'," + CStr(var_orden) + ",'" + VAR_NOMBRE_CHOFER + "','" + var_usuario_facturacion + "','" + CStr(var_tipo_transporte) + "','" + var_numero_unidad + "')", cnn, adOpenDynamic, adLockOptimistic
                           
                        Else
                           var_tipo = "Factura"
                           strconsulta = "SELECT A.NAME FROM OE_TRANSACTION_TYPES_TL A, OE_ORDER_HEADERS_ALL B WHERE A.TRANSACTION_TYPE_ID = B.ORDER_TYPE_ID AND B.ORDER_NUMBER = ? AND ROWNUM = 1"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rsaux9!source_header_number)
                                .Parameters.Append parametro
                           End With
                           Set rsaux7 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
'-----   ---
                           strconsulta = "SELECT  hps.pArty_site_number as clave_cliente , HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.invoice_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux9!source_header_number)
                                .Parameters.Append parametro
                           End With
                           Set rsaux6 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                               
                           If Not rsaux6.EOF Then
                              strconsulta = "SELECT  hps.pArty_site_number as clave_cliente ,HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = '" + CStr(CDbl(var_pedido)) + "' AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux9!source_header_number)
                                   .Parameters.Append parametro
                              End With
                              Set rsaux5 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                                    
                              If Not rsaux5.EOF Then
                                 var_clave_cliente = IIf(IsNull(rsaux5!clave_cliente), "", rsaux5!clave_cliente)
                                 VAR_DIRECCION = Mid(IIf(IsNull(rsaux5!calle), "", rsaux5!calle) + " " + IIf(IsNull(rsaux5!numero), "", rsaux5!numero), 1, 50)
                                 VAR_COLONIA = IIf(IsNull(rsaux5!colonia), "", rsaux5!colonia)
                                 var_ciudad = IIf(IsNull(rsaux5!ciudad), "", rsaux5!ciudad)
                                 VAR_MUNICIPIO = IIf(IsNull(rsaux5!municipio), "", rsaux5!municipio)
                                 var_estado = IIf(IsNull(rsaux5!estado), "", rsaux5!estado)
                                 var_pais = IIf(IsNull(rsaux5!pais), "", rsaux5!pais)
                                 VAR_CP = IIf(IsNull(rsaux5!cp), "", rsaux5!cp)
                                 rsaux5.Close
                              Else
                                 rsaux5.Close
                                 var_clave_cliente = IIf(IsNull(rsaux6!clave_cliente), "", rsaux6!clave_cliente)
                                 VAR_DIRECCION = Mid(IIf(IsNull(rsaux6!calle), "", rsaux6!calle) + " " + IIf(IsNull(rsaux6!numero), "", rsaux6!numero), 1, 50)
                                 VAR_COLONIA = IIf(IsNull(rsaux6!colonia), "", rsaux6!colonia)
                                 var_ciudad = IIf(IsNull(rsaux6!ciudad), "", rsaux6!ciudad)
                                 VAR_MUNICIPIO = IIf(IsNull(rsaux6!municipio), "", rsaux6!municipio)
                                 var_estado = IIf(IsNull(rsaux6!estado), "", rsaux6!estado)
                                 var_pais = IIf(IsNull(rsaux6!pais), "", rsaux6!pais)
                                 VAR_CP = IIf(IsNull(rsaux6!cp), "", rsaux6!cp)
                              End If
                           Else
                              VAR_DIRECCION = ""
                              VAR_COLONIA = ""
                              var_ciudad = ""
                              VAR_MUNICIPIO = ""
                              var_estado = ""
                              var_pais = ""
                              VAR_CP = ""
                           End If
                           rsaux6.Close
   
   
   
   
   
'--   ------
                        
   
   
                           strconsulta = "SELECT INTERFACE_HEADER_ATTRIBUTE1, RCT.customer_trx_id, HCAS.CUST_ACCOUNT_ID, APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME,HCSU.site_use_id,  sum(quantity_invoiced) as CANTIDAD  From hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, RA_CUSTOMER_TRX_ALL RCT,  ra_customer_trx_lines_all rctl, xxvia_importe_facturas APS Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND RCT.BILL_TO_SITE_USE_ID = HCSU.SITE_USE_ID AND INTERFACE_HEADER_ATTRIBUTE1 = ? AND INTERFACE_HEADER_ATTRIBUTE2 = ? and rctl.customer_trx_id = rct.customer_trx_id and extended_amount >0 AND APS.CUSTOMER_TRX_ID = RCT.CUSTOMER_TRX_ID GROUP BY INTERFACE_HEADER_ATTRIBUTE1, RCT.customer_trx_id, HCAS.CUST_ACCOUNT_ID, APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1, HCSU.site_use_id"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux9!source_header_number))
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux7!Name))
                                .Parameters.Append parametro
                           End With
                           Set rsaux6 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux6.EOF Then
                              var_i = 0
                              While Not rsaux6.EOF
                                    var_i = var_i + 1
                                    rsaux6.MoveNext
                              Wend
                              rsaux6.MoveFirst
                              var_j = 0
                              While Not rsaux6.EOF
                                    var_j = var_j + 1
                                    var_folio = rsaux6!TRX_NUMBER
                                    If var_j = var_i Then
                                    
                                       strconsulta = "SELECT DISTINCT INTE_PAQ_CAJA FROM XXVIA_TB_sALIDAS_CAJAS WHERE SOURCE_HEADER_NUMBER = ? AND FLOA_SAL_cANTIDAD_LEIDA > 0"
                                       With comandoORA
                                            .ActiveConnection = cnnoracle_4
                                            .CommandType = adCmdText
                                            .CommandText = strconsulta
                                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux9!source_header_number)
                                            .Parameters.Append parametro
                                       End With
                                       Set rsaux10 = comandoORA.execute
                                       Set comandoORA = Nothing
                                       Set parametro = Nothing
                                       VAR_CANTIDAD_BULTOS = 0
                                       While Not rsaux10.EOF
                                             VAR_CANTIDAD_BULTOS = VAR_CANTIDAD_BULTOS + 1
                                             rsaux10.MoveNext
                                       Wend
                                       rsaux10.Close
                                    Else
                                       VAR_CANTIDAD_BULTOS = 0
                                    End If
                                    rs.Open "INSERT INTO TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS (INTE_TEM_CONSECUTIVO, EMBARQUE, CLIENTE, CANTIDAD, PEDIDO, unidad, sellos, FECHA_EMBARQUE, RUTA, direccion_entrega, BULTOS, TIPO, folio, clave_cliente, ORDEN_ENTREGA, chofer, encargado_facturacion, NUMERO_ECONOMICO) VALUES (" + CStr(var_consecutivo) + "," + CStr(rsaux9!inte_emb_embarque) + ",'" + rsaux6!customer_name + "'," + CStr(IIf(IsNull(rsaux6!Cantidad), 0, rsaux6!Cantidad)) + "," + CStr(rsaux9!source_header_number) + ",'" + var_transporte + "','" + var_cadena_sellos + "', '" + var_fecha_embarque + "','" + VAR_CADENA_RUTAS + "',''," + CStr(VAR_CANTIDAD_BULTOS) + ",'" + var_tipo + "'," + CStr(var_folio) + ",'" + var_clave_cliente + "'," + CStr(var_orden) + ",'" + VAR_NOMBRE_CHOFER + "','" + var_usuario_facturacion + "','" + var_numero_unidad + "')", cnn, adOpenDynamic, adLockOptimistic
                                    rsaux6.MoveNext
                              Wend
                           End If
                           rsaux6.Close
                           rsaux7.Close
                        End If
                        rsaux8.Close
                        rsaux9.MoveNext
                  Wend
               End If
               rsaux15.MoveNext
         Wend
         rs.Open "DELETE FROM TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and ruta is null", cnn, adOpenDynamic, adLockOptimistic
         rs.Open "select distinct embarque as embarque from TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and tipo = 'Nota de envío'", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               var_ruta = ""
               rsaux.Open "select distinct cliente as cliente from TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and tipo = 'Nota de envío' and embarque = " + CStr(rs!Embarque), cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux.EOF
                     If var_ruta = "" Then
                        var_ruta = var_ruta + rsaux!Cliente
                     Else
                        var_ruta = var_ruta + ", " + rsaux!Cliente
                     End If
                     rsaux.MoveNext
               Wend
               rsaux.Close
               rsaux.Open "update TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS set ruta = ruta + '" + var_ruta + "' where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and embarque = " + CStr(rs!Embarque), cnn, adOpenDynamic, adLockOptimistic
               rs.MoveNext
         Wend
         rs.Close
         var_contador_embarques = 0
         rs.Open "select distinct embarque as embarque from TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux10.Open "select distinct tipo, folio  from TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS where inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and tipo is not null and embarque = " + CStr(rs!Embarque), cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux10.EOF Then
                  var_total_notas = 0
                  var_total_facturas = 0
                  While Not rsaux10.EOF
                        If rsaux10!tipo = "Nota de envío" Then
                           var_total_notas = var_total_notas + 1 'IIf(IsNull(rsaux10!Cantidad), 0, rsaux10!Cantidad)
                        End If
                        If rsaux10!tipo = "Factura" Then
                            var_total_facturas = var_total_facturas + 1 'IIf(IsNull(rsaux10!Cantidad), 0, rsaux10!Cantidad)
                        End If
                        rsaux10.MoveNext
                  Wend
               Else
                  var_total_notas = 0
                  var_total_facturas = 0
               End If
               rsaux10.Close
               var_contador_embarques = var_contador_embarques + 1
               rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS set TOTAL_NOTAS = " + CStr(var_total_notas) + ", TOTAL_FACTURAS = " + CStr(var_total_facturas) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and embarque = " + CStr(rs!Embarque), cnn, adOpenDynamic, adLockOptimistic
               rs.MoveNext
         Wend
         rs.Close
         rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS set TOTAL_embarques = " + CStr(var_contador_embarques) + ", fecha_inicio = " + var_fecha_inicio_2 + ", fecha_fin = " + var_fecha_finaL_2 + " where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Set reporte = appl.OpenReport(App.Path + "\repL_oracle_rendicion_cuentas.rpt")
         If Me.opt_ambos = True Then
            reporte.RecordSelectionFormula = "{VW_ORACLE_RENDICION_CUENTAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
         End If
         If Me.opt_externo = True Then
            reporte.RecordSelectionFormula = "{VW_ORACLE_RENDICION_CUENTAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_RENDICION_CUENTAS.TIPO_TRANSPORTE} = 0"
         End If
         If Me.opt_propio = True Then
            reporte.RecordSelectionFormula = "{VW_ORACLE_RENDICION_CUENTAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_RENDICION_CUENTAS.TIPO_TRANSPORTE} = 1"
         End If
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de rendición de cuentas"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            Set reporte = appl.OpenReport(App.Path + "\repL_oracle_rendicion_cuentas_excel.rpt")
            If Me.opt_ambos = True Then
               reporte.RecordSelectionFormula = "{VW_ORACLE_RENDICION_CUENTAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            End If
            If Me.opt_externo = True Then
               reporte.RecordSelectionFormula = "{VW_ORACLE_RENDICION_CUENTAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_RENDICION_CUENTAS.TIPO_TRANSPORTE} = 0"
            End If
            If Me.opt_propio = True Then
               reporte.RecordSelectionFormula = "{VW_ORACLE_RENDICION_CUENTAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_RENDICION_CUENTAS.TIPO_TRANSPORTE} = 1"
            End If
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\reporte_rendicion_cuentas_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
         End If
         
         rs.Open "DELETE FROM TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      Else
         MsgBox "No existe información para la fecha seleccionada", vbOKOnly, "ATENCION"
      End If
'--------
      rsaux15.Close
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha inicial incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 2300
   Left = 3800
   Me.txt_inicio = Date
   Me.txt_fin = Date
   Me.opt_ambos = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes = CDate(Me.txt_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      Me.txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      Me.txt_inicio = var_fecha_general
   End If
End Sub
