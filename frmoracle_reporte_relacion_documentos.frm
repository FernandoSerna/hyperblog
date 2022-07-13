VERSION 5.00
Begin VB.Form frmoracle_reporte_relacion_documentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relación de documentos"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   30
      TabIndex        =   4
      Top             =   405
      Width           =   3195
      Begin VB.TextBox txt_embarque 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1125
         TabIndex        =   0
         Top             =   180
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   345
         Width           =   765
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmoracle_reporte_relacion_documentos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2865
      Picture         =   "frmoracle_reporte_relacion_documentos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame x 
      Height          =   120
      Left            =   0
      TabIndex        =   3
      Top             =   285
      Width           =   3195
   End
End
Attribute VB_Name = "frmoracle_reporte_relacion_documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub cmd_imprimir_Click()
      If IsNumeric(Me.txt_embarque) Then
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
         If Not rsaux11.EOF Then
            var_chofer = IIf(IsNull(rsaux11!CHOFER), "", rsaux11!CHOFER)
         Else
            var_chofer = ""
         End If
         If var_chofer = "" Then
            VAR_NOMBRE_CHOFER = "                                                     "
         Else
            rsaux6.Open "select * from tb_choferes where vcha_cho_chofer_id = '" + var_chofer + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux6.EOF Then
               VAR_NOMBRE_CHOFER = IIf(IsNull(rsaux6!vcha_cho_nombre), "                                              ", rsaux6!vcha_cho_nombre)
            Else
               VAR_NOMBRE_CHOFER = "                                             "
            End If
            rsaux6.Close
         End If
         rsaux6.Open "select * from tb_usuarios where vcha_usu_usuario_id = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux6.EOF Then
            var_usuario_facturacion = IIf(IsNull(rsaux6!vcha_usu_nombre), "", rsaux6!vcha_usu_nombre) + " " + IIf(IsNull(rsaux6!vcha_usu_apellidos), "", rsaux6!vcha_usu_apellidos)
         Else
            var_usuario_facturacion = ""
         End If
         rsaux6.Close
         If Not rsaux9.EOF Then
            var_fecha_embarque = IIf(IsNull(rsaux11!fecha_fin), rsaux11!FECHA_INICIO, rsaux11!fecha_fin)
            rsaux5.Open "SELECT * FROM TB_ORACLE_TRANSPORTES where clave = '" + IIf(IsNull(rsaux11!transporte), "", rsaux11!transporte) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux5.EOF Then
               var_transporte = IIf(IsNull(rsaux5!NOMBRE), "", rsaux5!NOMBRE)
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
            'strconsulta = "SELECT DISTINCT  J.SALESREP_ID, J.NAME  FROM OE_ORDER_HEADERS_ALL OHA, XXVIA_TB_SALIDAS_CAJAS, XXVIA_VENDEDORES J WHERE  INTE_EMB_EMBARQUE = ? AND OHA.ORDER_NUMBER = SOURCE_HEADER_NUMBER AND OHA.SALESREP_ID = J.SALESREP_ID"
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
                  If rsaux8!ORDER_TYPE_ID = 1002 Or rsaux8!ORDER_TYPE_ID = 1023 Then
                     var_nombre_cliente = IIf(IsNull(rsaux9!Cliente), "", rsaux9!Cliente)
                     If var_nombre_cliente = "VIANNEY TEXTIL HOGAR SA DE CV" Or var_nombre_cliente = "LA TEXTILERA S.A DE C.V" Then
                        var_nombre_cliente = IIf(IsNull(rsaux9!ESTABLECIMIENTO), var_nombre_cliente, rsaux9!ESTABLECIMIENTO)
                     End If
                     var_tipo = "Nota de envío"
                     var_folio = rsaux9!source_header_number
                     
                     strconsulta = "SELECT DISTINCT INTE_PAQ_CAJA, TIPO_CAJA FROM XXVIA_TB_sALIDAS_CAJAS WHERE SOURCE_HEADER_NUMBER = ? AND FLOA_SAL_cANTIDAD_LEIDA > 0"
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
                     VAR_CANTIDAD_COSTALES = 0
                     While Not rsaux10.EOF
                           If Mid(rsaux10!tipo_caja, 1, 6) = "COSTAL" Or rsaux10!tipo_caja = "CAJA BIASI" Or rsaux10!tipo_caja = "CAJA BIASI" Then
                              VAR_CANTIDAD_COSTALES = VAR_CANTIDAD_COSTALES + 1
                           End If
                           VAR_CANTIDAD_BULTOS = VAR_CANTIDAD_BULTOS + 1
                           rsaux10.MoveNext
                     Wend
                     rsaux10.Close
                     rsaux10.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     'strconsulta = "select sum(floa_sal_Cantidad_leida) from xxvia_tb_salidas_cajas where source_header_number = ?"
                     rsaux17.Open "select * from tb_oracle_pedidos_Asignados_embarques where pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                     If rsaux17.RecordCount > 1 Then
                        strconsulta = "select SUM(floa_Sal_Cantidad_leida) from xxvia_tb_Salidas_cajas where  source_header_number = ? and inte_emb_embarque = ?"
                         With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux9!source_header_number)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                             .Parameters.Append parametro
                        End With
                        Set rsaux10 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                     
                     Else
                        strconsulta = "select SUM(SHIPPED_QUANTITY) from wsh_deliverables_v where RELEASED_STATUS = 'C' AND source_header_number = ?"
                     
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
                     
                     
                     End If
                     rsaux17.Close
                  
                     If Not rsaux10.EOF Then
                        var_cantidad_leida = IIf(IsNull(rsaux10(0).Value), 0, rsaux10(0).Value)
                     Else
                        var_cantidad_leida = var_cantidad_leida + 0
                     End If
                     rsaux10.Close
'------------
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
                        vvar_clave_cliente = ""
                        'var_nombre_cliente = ""
                     End If
                     rsaux7.Close





'------------
                     
                     
                     rs.Open "INSERT INTO TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS (INTE_TEM_CONSECUTIVO, EMBARQUE, CLIENTE, CANTIDAD, PEDIDO, unidad, sellos, FECHA_EMBARQUE, RUTA, direccion_entrega, BULTOS, TIPO, folio, clave_cliente, ORDEN_ENTREGA, chofer, encargado_facturacion, COSTALES) VALUES (" + CStr(var_consecutivo) + "," + CStr(rsaux9!inte_emb_embarque) + ",'" + var_nombre_cliente + "'," + CStr(var_cantidad_leida) + "," + CStr(rsaux9!source_header_number) + ",'" + var_transporte + "','" + var_cadena_sellos + "', '" + var_fecha_embarque + "','" + VAR_CADENA_RUTAS + "',''," + CStr(VAR_CANTIDAD_BULTOS) + ",'" + var_tipo + "'," + CStr(var_folio) + ",'" + var_clave_cliente + "'," + CStr(var_orden) + ",'" + VAR_NOMBRE_CHOFER + "','" + var_usuario_facturacion + "'," + CStr(VAR_CANTIDAD_COSTALES) + ")", cnn, adOpenDynamic, adLockOptimistic
                     
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
'--------
                     strconsulta = "SELECT  hps.pArty_site_number as clave_cliente , HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.SHIP_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
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
                        strconsulta = "SELECT  hps.pArty_site_number as clave_cliente ,HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
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
                           VAR_DIRECCION = Mid(IIf(IsNull(rsaux6!calle), "", rsaux6!calle) + " " + IIf(IsNull(rsaux6!numero), "", rsaux6!numero), 1, 50)
                           VAR_COLONIA = IIf(IsNull(rsaux6!colonia), "", rsaux6!colonia)
                           var_ciudad = IIf(IsNull(rsaux6!ciudad), "", rsaux6!ciudad)
                           VAR_MUNICIPIO = IIf(IsNull(rsaux6!municipio), "", rsaux6!municipio)
                           var_estado = IIf(IsNull(rsaux6!estado), "", rsaux6!estado)
                           var_pais = IIf(IsNull(rsaux6!pais), "", rsaux6!pais)
                           VAR_CP = IIf(IsNull(rsaux6!cp), "", rsaux6!cp)
                           VAR_DIRECCION = VAR_DIRECCION + " " + VAR_COLONIA + " " + var_ciudad + " " + var_estado + " CP:" + VAR_CP
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
                           VAR_DIRECCION = VAR_DIRECCION + " " + VAR_COLONIA + " " + var_ciudad + " " + var_estado + " CP:" + VAR_CP
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





'--------
                     


                     strconsulta = "SELECT INTERFACE_HEADER_ATTRIBUTE1, RCT.customer_trx_id, HCAS.CUST_ACCOUNT_ID, APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME,HCSU.site_use_id,  sum(quantity_invoiced) as CANTIDAD, RCT.SHIP_TO_SITE_USE_ID  From hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, RA_CUSTOMER_TRX_ALL RCT,  ra_customer_trx_lines_all rctl, xxvia_importe_facturas APS Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND RCT.BILL_TO_SITE_USE_ID = HCSU.SITE_USE_ID AND INTERFACE_HEADER_ATTRIBUTE1 = ? AND INTERFACE_HEADER_ATTRIBUTE2 = ? and rctl.customer_trx_id = rct.customer_trx_id and extended_amount >0 AND APS.CUSTOMER_TRX_ID = RCT.CUSTOMER_TRX_ID GROUP BY INTERFACE_HEADER_ATTRIBUTE1, RCT.customer_trx_id, HCAS.CUST_ACCOUNT_ID, APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1, "
                     strconsulta = strconsulta + " HCSU.site_use_id, RCT.SHIP_TO_SITE_USE_ID "
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
                        var_ship_to_site_use_id = rsaux6!SHIP_TO_SITE_USE_ID
                        
                        strconsulta = "select party_site_number  from xxvia_vw_clientes_bcp where site_use_id = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_ship_to_site_use_id)
                             .Parameters.Append parametro
                        End With
                        Set rsaux17 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If Not rsaux17.EOF Then
                           var_establecimiento = rsaux17!party_site_number
                        Else
                           var_establecimiento = ""
                        End If
                        rsaux17.Close
                        
                        
                        var_i = 0
                        While Not rsaux6.EOF
                              var_i = var_i + 1
                              rsaux6.MoveNext
                        Wend
                        rsaux6.MoveFirst
                        var_j = 0
                        While Not rsaux6.EOF
                              var_j = var_j + 1
                              var_folio = rsaux6!trx_number
                              VAR_CANTIDAD_COSTALES = 0
                              If var_j = var_i Then
                                 
                                 strconsulta = "SELECT DISTINCT INTE_PAQ_CAJA, TIPO_CAJA FROM XXVIA_TB_sALIDAS_CAJAS WHERE SOURCE_HEADER_NUMBER = ? AND FLOA_SAL_cANTIDAD_LEIDA > 0"
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
                                 VAR_CANTIDAD_COSTALES = 0
                                 While Not rsaux10.EOF
                                       If Mid(rsaux10!tipo_caja, 1, 6) = "COSTAL" Then
                                          VAR_CANTIDAD_COSTALES = VAR_CANTIDAD_COSTALES + 1
                                       End If
                                       VAR_CANTIDAD_BULTOS = VAR_CANTIDAD_BULTOS + 1
                                       rsaux10.MoveNext
                                 Wend
                                 rsaux10.Close
                              Else
                                 VAR_CANTIDAD_BULTOS = 0
                              End If
                              rs.Open "INSERT INTO TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS (INTE_TEM_CONSECUTIVO, EMBARQUE, CLIENTE, CANTIDAD, PEDIDO, unidad, sellos, FECHA_EMBARQUE, RUTA, direccion_entrega, BULTOS, TIPO, folio, clave_cliente, ORDEN_ENTREGA, chofer, encargado_facturacion, COSTALES, ESTABLECIMIENTO) VALUES (" + CStr(var_consecutivo) + "," + CStr(rsaux9!inte_emb_embarque) + ",'" + rsaux6!customer_name + "'," + CStr(IIf(IsNull(rsaux6!cantidad), 0, rsaux6!cantidad)) + "," + CStr(rsaux9!source_header_number) + ",'" + var_transporte + "','" + var_cadena_sellos + "', '" + var_fecha_embarque + "','" + VAR_CADENA_RUTAS + "','" + VAR_DIRECCION + "'," + CStr(VAR_CANTIDAD_BULTOS) + ",'" + var_tipo + "'," + CStr(var_folio) + ",'" + var_clave_cliente + "'," + CStr(var_orden) + ",'" + VAR_NOMBRE_CHOFER + "','" + var_usuario_facturacion + "'," + CStr(IIf(IsNull(VAR_CANTIDAD_COSTALES), 0, VAR_CANTIDAD_COSTALES)) + ",'" + var_establecimiento + "')", cnn, adOpenDynamic, adLockOptimistic
                              rsaux6.MoveNext
                        Wend
                     End If
                     rsaux6.Close
                     rsaux7.Close
                     ''FACTURA COSTALES
                     var_tipo = "Factura bultos"
                     
                     strconsulta = "SELECT * FROM OE_ORDER_HEADERS_ALL WHERE orig_sys_document_ref  = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "SIDCOSTALES_" + Trim(CStr(rsaux9!source_header_number)))
                          .Parameters.Append parametro
                     End With
                     Set rsaux12 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If Not rsaux12.EOF Then
                        VAR_PEDIDO_COSTALES = CDbl(rsaux12!order_number)
                                     
                        strconsulta = "SELECT A.NAME FROM OE_TRANSACTION_TYPES_TL A, OE_ORDER_HEADERS_ALL B WHERE A.TRANSACTION_TYPE_ID = B.ORDER_TYPE_ID AND B.ORDER_NUMBER = ? AND ROWNUM = 1"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, VAR_PEDIDO_COSTALES)
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
'--------
                        strconsulta = "SELECT  hps.pArty_site_number as clave_cliente , HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.SHIP_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, VAR_PEDIDO_COSTALES)
                             .Parameters.Append parametro
                        End With
                        Set rsaux6 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                             
                        If Not rsaux6.EOF Then
                           strconsulta = "SELECT  hps.pArty_site_number as clave_cliente ,HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, VAR_PEDIDO_COSTALES)
                                .Parameters.Append parametro
                           End With
                           Set rsaux5 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                                 
                           If Not rsaux5.EOF Then
                              var_clave_cliente = IIf(IsNull(rsaux5!clave_cliente), "", rsaux5!clave_cliente)
                              VAR_DIRECCION = Mid(IIf(IsNull(rsaux6!calle), "", rsaux6!calle) + " " + IIf(IsNull(rsaux6!numero), "", rsaux6!numero), 1, 50)
                              VAR_COLONIA = IIf(IsNull(rsaux6!colonia), "", rsaux6!colonia)
                              var_ciudad = IIf(IsNull(rsaux6!ciudad), "", rsaux6!ciudad)
                              VAR_MUNICIPIO = IIf(IsNull(rsaux6!municipio), "", rsaux6!municipio)
                              var_estado = IIf(IsNull(rsaux6!estado), "", rsaux6!estado)
                              var_pais = IIf(IsNull(rsaux6!pais), "", rsaux6!pais)
                              VAR_CP = IIf(IsNull(rsaux6!cp), "", rsaux6!cp)
                              VAR_DIRECCION = VAR_DIRECCION + " " + VAR_COLONIA + " " + var_ciudad + " " + var_estado + " CP:" + VAR_CP
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
                              VAR_DIRECCION = VAR_DIRECCION + " " + VAR_COLONIA + " " + var_ciudad + " " + var_estado + " CP:" + VAR_CP
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





'--------
                        
   
   
                        strconsulta = "SELECT INTERFACE_HEADER_ATTRIBUTE1, RCT.customer_trx_id, HCAS.CUST_ACCOUNT_ID, APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME,HCSU.site_use_id,  sum(quantity_invoiced) as CANTIDAD  From hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, RA_CUSTOMER_TRX_ALL RCT,  ra_customer_trx_lines_all rctl, xxvia_importe_facturas APS Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND RCT.BILL_TO_SITE_USE_ID = HCSU.SITE_USE_ID AND INTERFACE_HEADER_ATTRIBUTE1 = ? AND INTERFACE_HEADER_ATTRIBUTE2 = ? and rctl.customer_trx_id = rct.customer_trx_id and extended_amount >0 AND APS.CUSTOMER_TRX_ID = RCT.CUSTOMER_TRX_ID GROUP BY INTERFACE_HEADER_ATTRIBUTE1, RCT.customer_trx_id, HCAS.CUST_ACCOUNT_ID, APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1, HCSU.site_use_id"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(VAR_PEDIDO_COSTALES))
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux7!Name))
                             .Parameters.Append parametro
                        End With
                        Set rsaux6 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If Not rsaux6.EOF Then
                           While Not rsaux6.EOF
                                 var_cantidad = IIf(IsNull(rsaux6!cantidad), 0, rsaux6!cantidad)
                                 var_folio = rsaux6!trx_number
                                 VAR_CANTIDAD_BULTOS = 0
                                 rs.Open "INSERT INTO TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS (INTE_TEM_CONSECUTIVO, EMBARQUE, CLIENTE, CANTIDAD, PEDIDO, unidad, sellos, FECHA_EMBARQUE, RUTA, direccion_entrega, BULTOS, TIPO, folio, clave_cliente, ORDEN_ENTREGA, chofer, encargado_facturacion,BULTOS_FACTURADOS) VALUES (" + CStr(var_consecutivo) + "," + CStr(rsaux9!inte_emb_embarque) + ",'" + rsaux6!customer_name + "',0," + CStr(VAR_PEDIDO_COSTALES) + ",'" + var_transporte + "','" + var_cadena_sellos + "', '" + var_fecha_embarque + "','" + VAR_CADENA_RUTAS + "','" + VAR_DIRECCION + "',0,'Factura bultos'," + CStr(var_folio) + ",'" + var_clave_cliente + "'," + CStr(var_orden) + ",'" + VAR_NOMBRE_CHOFER + "','" + var_usuario_facturacion + "'," + CStr(var_cantidad) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 rsaux6.MoveNext
                           Wend
                        End If
                        rsaux6.Close
                        rsaux7.Close
                     End If
                     rsaux12.Close
                     
                     ''FIN FACTURA COSTALES
                  End If
                  rsaux8.Close
                  

                  rsaux9.MoveNext
            Wend
            var_cadena_pedidos_tiendas = ""
            var_cadena_pedidos_clientes = ""
            rsaux10.Open "select distinct tipo, folio, pedido  from TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS where inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and tipo is not null ", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux10.EOF Then
               var_total_notas = 0
               var_total_facturas = 0
               While Not rsaux10.EOF
                     If rsaux10!tipo = "Nota de envío" Then
                         If var_cadena_pedidos_tiendas = "" Then
                            var_cadena_pedidos_tiendas = CStr(rsaux10!pedido)
                         Else
                            var_cadena_pedidos_tiendas = var_cadena_pedidos_tiendas + ", " + CStr(rsaux10!pedido)
                         End If
                         var_total_notas = var_total_notas + 1 'IIf(IsNull(rsaux10!Cantidad), 0, rsaux10!Cantidad)
                     End If
                     If rsaux10!tipo = "Factura" Or rsaux10!tipo = "Factura bultos" Then
                         var_total_facturas = var_total_facturas + 1 'IIf(IsNull(rsaux10!Cantidad), 0, rsaux10!Cantidad)
                         If var_cadena_pedidos_clientes = "" Then
                            var_cadena_pedidos_clientes = CStr(rsaux10!pedido)
                         Else
                            var_cadena_pedidos_clientes = var_cadena_pedidos_clientes + ", " + CStr(rsaux10!pedido)
                         End If
                     End If
                     rsaux10.MoveNext
               Wend
               
            Else
               var_total_notas = 0
               var_total_facturas = 0
            End If
            rsaux10.Close
                                                        
            var_cantidad_leida = 0
            strconsulta = "SELECT SUM(FLOA_SAL_CANTIDAD_LEIDA) FROM XXVIA_TB_sALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = ? "
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                 .Parameters.Append parametro
            End With
            Set rsaux10 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If Not rsaux10.EOF Then
               var_cantidad_leida = var_cantidad_leida + IIf(IsNull(rsaux10(0).Value), 0, rsaux10(0).Value)
            End If
            rsaux10.Close
            
            strconsulta = "SELECT SUM(FLOA_SAL_CANTIDAD_LEIDA) FROM XXVIA_TB_SALIDAS WHERE INTE_EMB_EMBARQUE = ? "
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                 .Parameters.Append parametro
            End With
            Set rsaux10 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If Not rsaux10.EOF Then
               var_cantidad_leida = var_cantidad_leida + IIf(IsNull(rsaux10(0).Value), 0, rsaux10(0).Value)
            End If
            rsaux10.Close
            
            
            rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS SET TOTAL_NOTAS = " + CStr(var_total_notas) + ", TOTAL_FACTURAS = " + CStr(var_total_facturas) + ", CANTIDAD_LEIDA = " + CStr(var_cantidad_leida) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            
            rs.Open "DELETE FROM TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and ruta is null", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "DELETE FROM TB_tEMP_ORACLE_RELACION_COSTALES WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            
            rs.Open "SELECT CLAVE_CLIENTE, CLIENTE, TIPO, MAX(PEDIDO) AS PEDIDO FROM TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " GROUP BY CLAVE_CLIENTE, CLIENTE, TIPO", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  If rs!tipo = "Factura" Then
                     var_cadena = "select C.PARTY_SITE_NUMBER, SUM(INVOICED_QUANTITY) AS CANTIDAD, e.segment1 as CODIGO, e.Description from oe_order_lines_all A, oe_order_headers_all b, XXVIA_VW_CLIENTES_BCP c, xxvia_system_items_b e where A.inventory_item_id in (33578,1873944) and b.INVOICE_TO_ORG_ID = C.site_use_id and a.inventory_item_id = e.inventory_item_id and a.header_id = b.header_id and e.organization_id = 93 and nvl(invoiced_quantity,0) <> 0 AND c.PARTY_SITE_NUMBER = ?"
                     var_cadena = var_cadena + " GROUP BY C.PARTY_SITE_NUMBER,E.SEGMENT1,E.DESCRIPTION"
                     With comandoORA
                         .ActiveConnection = cnnoracle_4
                         .CommandType = adCmdText
                         .CommandText = var_cadena
                         Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rs!clave_cliente)
                         .Parameters.Append parametro
                     End With
                     Set rsaux10 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     While Not rsaux10.EOF
                           rsaux12.Open "INSERT INTO TB_TEMP_ORACLE_RELACION_COSTALES (INTE_TEM_CONSECUTIVO, CLAVE_CLIENTE, CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, EMBARQUE)  VALUES (" + CStr(var_consecutivo) + ",'" + rs!clave_cliente + "','" + rs!Cliente + "', '" + rsaux10!codigo + "','" + rsaux10!Description + "'," + CStr(IIf(IsNull(rsaux10!cantidad), 0, rsaux10!cantidad)) + "," + Me.txt_embarque + ")", cnn, adOpenDynamic, adLockOptimistic
                           rsaux10.MoveNext
                     Wend
                     rsaux10.Close
                  End If
                  If rs!tipo = "Nota de envío" Then
                     strconsulta = "SELECT x.ATTRIBUTE1 FROM po_requisition_headers_ALL x, MTL_SECONDARY_INVENTORIES Y, OE_ORDER_HEADERS_ALL A Where requisition_header_id = A.source_document_id   AND secondary_inventory_name = x.ATTRIBUTE1 and rownum =1 AND ORDER_NUMBER = ?"
                     With comandoORA
                         .ActiveConnection = cnnoracle_4
                         .CommandType = adCmdText
                         .CommandText = strconsulta
                         Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rs!pedido)
                         .Parameters.Append parametro
                     End With
                     Set rsaux10 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If Not rsaux10.EOF Then
                        var_clave_cliente = rsaux10!attribute1
                     Else
                        var_clave_cliente = ""
                     End If
                     rsaux10.Close
                     rsaux10.Open "select * from tb_oracle_empaques WHERE LEN(ISNULL(CODIGO,''))=8", cnn, adOpenDynamic, adLockOptimistic
                     While Not rsaux10.EOF
                           strconsulta = "select * from Xxvia_vw_existencias_inv where organization_id = ? and subinventory_code = ? and segment1 = ?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_unidad_organizacional))
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_clave_cliente)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rsaux10!codigo)
                                .Parameters.Append parametro
                           End With
                           Set rsaux9 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux9.EOF Then
                              rsaux12.Open "INSERT INTO TB_TEMP_ORACLE_RELACION_COSTALES (INTE_TEM_CONSECUTIVO, CLAVE_CLIENTE, CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, EMBARQUE)  VALUES (" + CStr(var_consecutivo) + ",'" + var_clave_cliente + "','" + rs!Cliente + "', '" + rsaux10!codigo + "','" + rsaux9!Description + "'," + CStr(IIf(IsNull(rsaux9!CANTMANO), 0, rsaux9!CANTMANO)) + "," + Me.txt_embarque + ")", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux9.Close
                           rsaux10.MoveNext
                     Wend
                     rsaux10.Close
                  End If
                  rs.MoveNext
            Wend
            rs.Close
            
            
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_relacion_documentos.rpt")
            reporte.RecordSelectionFormula = "{VW_ORACLE_REPORTE_RELACION_DOCUMENTOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Relación de documentos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            rs.Open "SELECT * FROM VW_ORACLE_RELACION_COSTALES_FACTURAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_relacion_costales_facturas.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_RELACION_COSTALES_FACTURAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Relación de existencias de bultos por cliente"
               frmvistasprevias.Show 1
               Set reporte = Nothing
            Else
               MsgBox "No existe relación de existencia de bultos.", vbOKOnly, "ATENCION"
            End If
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "DELETE FROM TB_TEMP_ORACLE_REPORTE_RELACION_DOCUMENTOS WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_ORACLE_RELACION_COSTALES where INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                    
        Else
           MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
         If rsaux9.State = 1 Then
            rsaux9.Close
         End If
         rsaux11.Close
      End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 3200
   Left = 4200
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

