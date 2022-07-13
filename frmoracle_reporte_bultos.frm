VERSION 5.00
Begin VB.Form frmoracle_reporte_bultos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de bultos"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2940
      Picture         =   "frmoracle_reporte_bultos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmoracle_reporte_bultos.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame x 
      Height          =   120
      Left            =   75
      TabIndex        =   4
      Top             =   285
      Width           =   3195
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   105
      TabIndex        =   3
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
         Top             =   255
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmoracle_reporte_bultos"
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
         strconsulta = "SELECT EMBARQUE, CHAR_EMB_ESTATUS, CASE CHAR_EMB_ESTATUS WHEN 'F' THEN 'CERRADO' WHEN 'I' THEN 'CERRADO' ELSE 'ABIERTO' END ESTATUS FROM XXVIA_TB_ENCABEZADO_eMBARQUES WHERE EMBARQUE = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
              .Parameters.Append parametro
         End With
         Set rsaux14 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rsaux14.EOF Then
            If IIf(IsNull(rsaux14!estatus), "ABIERTO", rsaux14!estatus) = "CERRADO" Then
            
               strconsulta = "SELECT INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER, MAX(CUSTOMER_NAME) AS CLIENTE, MAX(ENTREGA) AS ESTABLECIMIENTO, SUM(FLOA_SAL_CANTIDAD_LEIDA) AS CANTIDAD FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = ? and floa_Sal_Cantidad_leida > 0 GROUP BY INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER  ORDER BY INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER"
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
               rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_ORACLE_REPORTE_CAJAS", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
               Else
                  var_consecutivo = 0
               End If
               var_consecutivo = var_consecutivo + 1
               rs.Close
               rs.Open "insert into TB_TEMP_ORACLE_REPORTE_CAJAS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               rs.Open "delete from TB_TEMP_ORACLE_CAJAS_ADUANA_DIVIDIDAS_EN_3 where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               strconsulta = "SELECT TRANSPORTE, TO_CHAR(FECHA_FIN, 'DD/MM/YYYY HH24:MI:SS') AS FECHA_FIN, TO_CHAR(FECHA_INICIO, 'DD/MM/YYYY HH24:MI:SS') FECHA_INICIO, USUARIO_CERRO FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = ?"
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
               rsaux10.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + IIf(IsNull(rsaux11!USUARIO_CERRO), "", rsaux11!USUARIO_CERRO) + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux10.EOF Then
                  VAR_USUARIO_CERRO = IIf(IsNull(rsaux10!vcha_usu_nombre), "", rsaux10!vcha_usu_nombre) + " " + IIf(IsNull(rsaux10!vcha_usu_apellidos), "", rsaux10!vcha_usu_apellidos)
                  If VAR_USUARIO_CERRO = "" Then
                     rsaux6.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux6.EOF Then
                        VAR_USUARIO_CERRO = IIf(IsNull(rsaux6!vcha_usu_nombre), "", rsaux6!vcha_usu_nombre) + " " + IIf(IsNull(rsaux6!vcha_usu_apellidos), "", rsaux6!vcha_usu_apellidos)
                     Else
                        VAR_USUARIO_CERRO = ""
                     End If
                     rsaux6.Close
                  End If
               Else
                  rsaux6.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux6.EOF Then
                     VAR_USUARIO_CERRO = IIf(IsNull(rsaux6!vcha_usu_nombre), "", rsaux6!vcha_usu_nombre) + " " + IIf(IsNull(rsaux6!vcha_usu_apellidos), "", rsaux6!vcha_usu_apellidos)
                  Else
                     VAR_USUARIO_CERRO = ""
                  End If
                  rsaux6.Close
               End If
               rsaux10.Close
               If Not rsaux9.EOF Then
                  var_fecha_embarque = IIf(IsNull(rsaux11!fecha_fin), rsaux11!FECHA_INICIO, rsaux11!fecha_fin)
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
                        strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rsaux9!SOURCE_HEADER_NUMBER)
                             .Parameters.Append parametro
                        End With
                        Set rsaux8 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If rsaux8!order_type_id = 1002 Then
                           var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
'----
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
                              var_nombre_cliente = rsaux7!Description
                           Else
                              var_clave_cliente = ""
                              var_nombre_cliente = ""
                           End If
                           rsaux7.Close
                           If var_almacen_tienda <> "" Then
                        
                              strconsulta = "select * from mtl_secondary_inventories where secondary_inventory_name = ?"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_almacen_tienda)
                                   .Parameters.Append parametro
                              End With
                              Set rsaux3 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                           
                              If Not rsaux3.EOF Then
                                 var_location_id = IIf(IsNull(rsaux3!LOCATION_ID), 0, rsaux3!LOCATION_ID)
                                 If var_location_id > 0 Then
                                 
                                    strconsulta = "select ADDRESS_LINE_1, ADDRESS_LINE_2, TOWN_OR_CITY, REGION_1, COUNTRY, POSTAL_CODE  from hr_locations_all where location_id = ?"
                                    With comandoORA
                                         .ActiveConnection = cnnoracle_4
                                         .CommandType = adCmdText
                                         .CommandText = strconsulta
                                         Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_location_id)
                                         .Parameters.Append parametro
                                    End With
                                    Set rsaux4 = comandoORA.execute
                                    Set comandoORA = Nothing
                                    Set parametro = Nothing
                                    If Not rsaux4.EOF Then
                                       VAR_DIRECCION = Mid(IIf(IsNull(rsaux4!ADDRESS_LINE_1), "", rsaux4!ADDRESS_LINE_1), 1, 50)
                                       VAR_COLONIA = IIf(IsNull(rsaux4!ADDRESS_LINE_2), "", rsaux4!ADDRESS_LINE_2)
                                       var_ciudad = IIf(IsNull(rsaux4!TOWN_OR_CITY), "", rsaux4!TOWN_OR_CITY)
                                       var_estado = IIf(IsNull(rsaux4!REGION_1), "", rsaux4!REGION_1)
                                       var_pais = IIf(IsNull(rsaux4!COUNTRY), "", rsaux4!COUNTRY)
                                       VAR_CP = IIf(IsNull(rsaux4!POSTAL_CODE), "", rsaux4!POSTAL_CODE)
                                    End If
                                    rsaux4.Close
                                 End If
                              Else
                                 VAR_DIRECCION = ""
                                 VAR_COLONIA = ""
                                 var_ciudad = ""
                                 var_estado = ""
                                 var_pais = ""
                                 VAR_CP = ""
                              End If
                              rsaux3.Close
                           Else
                              VAR_DIRECCION = ""
                              VAR_COLONIA = ""
                              var_ciudad = ""
                              var_estado = ""
                              var_pais = ""
                              VAR_CP = ""
                           End If
                        Else
                     
                           strconsulta = "SELECT  hps.pArty_site_number as clave_cliente , HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.invoice_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux9!SOURCE_HEADER_NUMBER)
                                .Parameters.Append parametro
                           End With
                           Set rsaux6 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                                
                           If Not rsaux6.EOF Then
                              strconsulta = "SELECT  hps.pArty_site_number as clave_cliente ,HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux9!SOURCE_HEADER_NUMBER)
                                   .Parameters.Append parametro
                              End With
                              Set rsaux5 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                                    
                              If Not rsaux5.EOF Then
                                 var_clave_cliente = IIf(IsNull(rsaux5!clave_cliente), "", rsaux5!clave_cliente)
                                 VAR_DIRECCION = Mid(IIf(IsNull(rsaux5!calle), "", rsaux5!calle) + " " + IIf(IsNull(rsaux5!NUMERO), "", rsaux5!NUMERO), 1, 50)
                                 VAR_COLONIA = IIf(IsNull(rsaux5!colonia), "", rsaux5!colonia)
                                 var_ciudad = IIf(IsNull(rsaux5!ciudad), "", rsaux5!ciudad)
                                 VAR_MUNICIPIO = IIf(IsNull(rsaux5!municipio), "", rsaux5!municipio)
                                 var_estado = IIf(IsNull(rsaux5!estado), "", rsaux5!estado)
                                 var_pais = IIf(IsNull(rsaux5!pais), "", rsaux5!pais)
                                 VAR_CP = IIf(IsNull(rsaux5!cp), "", rsaux5!cp)
                                 VAR_DIRECCION = IIf(IsNull(rsaux5!customer_name), "", rsaux5!customer_name) + ", Direcci�n de entrega: " + VAR_DIRECCION
                                 rsaux5.Close
                              Else
                                 rsaux5.Close
                                 var_clave_cliente = IIf(IsNull(rsaux6!clave_cliente), "", rsaux6!clave_cliente)
                                 VAR_DIRECCION = Mid(IIf(IsNull(rsaux6!calle), "", rsaux6!calle) + " " + IIf(IsNull(rsaux6!NUMERO), "", rsaux6!NUMERO), 1, 50)
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
                        End If
                        rsaux8.Close
                     
                        var_direccion_str = VAR_DIRECCION + ", " + VAR_COLONIA + ", " + var_ciudad + ", " + VAR_MUNICIPIO + ", " + var_estado + ", " + var_pais + ", CP: " + VAR_CP
                     
                     
                     
                     
                        var_nombre_cliente = IIf(IsNull(rsaux9!Cliente), "", rsaux9!Cliente)
                        If var_nombre_cliente = "VIANNEY TEXTIL HOGAR SA DE CV" Then
                           var_nombre_cliente = IIf(IsNull(rsaux9!ESTABLECIMIENTO), var_nombre_cliente, rsaux9!ESTABLECIMIENTO)
                        End If
                     
                        strconsulta = "SELECT SOURCE_HEADER_NUMBER, TIPO_CAJA, COUNT(TIPO_CAJA) AS CANTIDAD FROM XXVIA_VW_CANTIDAD_BULTOS WHERE SOURCE_HEADER_NUMBER = ? and tipo_caja is not null GROUP BY SOURCE_HEADER_NUMBER, TIPO_CAJA"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rsaux9!SOURCE_HEADER_NUMBER)
                             .Parameters.Append parametro
                        End With
                        Set rsaux8 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        rsaux7.Open "select * from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES WHERE PEDIDO = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux7.EOF Then
                           var_orden = IIf(IsNull(rsaux7!orden_pedido), 0, rsaux7!orden_pedido)
                        Else
                           var_orden = 0
                        End If
                        rsaux7.Close
                        rs.Open "INSERT INTO TB_TEMP_ORACLE_REPORTE_CAJAS (INTE_TEM_CONSECUTIVO, EMBARQUE, CLIENTE, CANTIDAD, PEDIDO, unidad, sellos, FECHA_EMBARQUE, RUTA, direccion_entrega, CLAVE_CLIENTE, ORDEN_ENTREGA, USUARIO_CERRO) VALUES (" + CStr(var_consecutivo) + "," + CStr(rsaux9!inte_Emb_Embarque) + ",'" + Replace(var_nombre_cliente, "'", " ") + "'," + CStr(IIf(IsNull(rsaux9!cantidad), 0, rsaux9!cantidad)) + "," + CStr(rsaux9!SOURCE_HEADER_NUMBER) + ",'" + var_transporte + "','" + var_cadena_sellos + "', '" + var_fecha_embarque + "','" + VAR_CADENA_RUTAS + "','" + Replace(var_direccion_str, "'", " ") + "','" + var_clave_cliente + "'," + CStr(var_orden) + ",'" + VAR_USUARIO_CERRO + "')", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux8.EOF
                              If rsaux8!tipo_caja = "CAJA EXTRAGRANDE" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_EXTRAGRANDE = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "CAJA GRANDE" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_GRANDE = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "CAJA MEDIANA" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_MEDIANA = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "CAJA CHICA" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_CHICA = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "CAJA MINI/CATALOGO" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_MINI_CATALOGO = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "CAJA  SOBRE-CAJA" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_SOBRE_CAJA = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "CAJA CORTINERO" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_CORTINERO = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "COSTAL GRANDE" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set COSTAL_GRANDE = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "COSTAL CHICO" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set COSTAL_CHICO = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "EMPLAYE CORTINEROS" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set EMPLAYE_CORTINEROS = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "PAQUETE BOLSA" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set PAQUETE_BOLSA = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "PAQUETE PUBLICIDAD" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set PAQUETE_PUBLICIDAD = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "OTROS" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set OTROS = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "OTROS MUEBLES" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set OTROS_MUEBLES = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "CAJA CHICA C/CATALOGO" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_CHICA_CATALOGO = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "CAJA BIASI" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_GRIS = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "COSTAL EXTRAGRANDE" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set COSTAL_EXTRAGRANDE = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "COSTAL MEDIANO" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set COSTAL_MEDIANO = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!SOURCE_HEADER_NUMBER), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux8.MoveNext
                        Wend
                        
                        'If CDbl(Me.txt_embarque) <> 140832 Then
                        '   strConsulta = "SELECT DISTINCT INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER, INTE_PAQ_CAJA, CAJA_PEDIDO, TIPO_CAJA, SELLO, AUDITADA FROM XXVIA_TB_SALIDAS_CAJAS WHERE inte_emb_embarque = ? AND INTE_PAQ_CAJA > 0 and source_header_number = ? order by inte_paq_caja"
                        '   With comandoORA
                        '       .ActiveConnection = cnnoracle_4
                        '       .CommandType = adCmdText
                        '       .CommandText = strConsulta
                        '       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
                        '       .Parameters.Append parametro
                        '       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rsaux9!SOURCE_HEADER_NUMBER)
                        '        .Parameters.Append parametro
                        '   End With
                        '   Set rsaux7 = comandoORA.execute
                        '   Set comandoORA = Nothing
                        '   Set parametro = Nothing
                        'Else
                           rsaux7.Open "SELECT DISTINCT INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER, INTE_PAQ_CAJA, CAJA_PEDIDO, TIPO_CAJA, SELLO, AUDITADA FROM XXVIA_TB_SALIDAS_CAJAS WHERE inte_emb_embarque = " + Me.txt_embarque + " AND INTE_PAQ_CAJA > 0 and source_header_number = " + CStr(rsaux9!SOURCE_HEADER_NUMBER) + " and nvl(caja_pedido,0) <> 0 and nvl(inte_paq_caja_anterior,0) = 0 order by inte_paq_caja", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'End If
                      
                        var_i = 0
                        While Not rsaux7.EOF
                              var_i = var_i + 1
                              rsaux7.MoveNext
                        Wend
                
                        VAR_Z = Round(var_i / 3, 2)
                        VAR_Y = Round(var_i / 3, 0)
                        var_x = VAR_Z - VAR_Y
                        If var_x < 0.5 Then
                           If var_x = 0 Then
                              VAR_Z = Round(var_i / 3, 0)
                           Else
                              VAR_Z = Round(Round(var_i / 3, 2) + 0.5, 0)
                           End If
                        Else
                           VAR_Z = Round(var_i / 3, 0)
                        End If
               
                        rsaux7.MoveFirst
                        'MsgBox rsaux7.RecordCount
                        var_j = 0
                        var_m = 1
                        While Not rsaux7.EOF
                              If var_j = VAR_Z Then
                                 var_j = 0
                                 var_m = var_m + 1
                              End If
                              var_j = var_j + 1
                              If Not rsaux7.EOF Then
                                 var_numero_caja = rsaux7!INTE_PAQ_CAJA
                                 If Len(Trim(Str(var_numero_caja))) = 1 Then
                                    var_referencia_caja = "00" + Trim(Str(var_numero_caja))
                                 End If
                                 If Len(Trim(Str(var_numero_caja))) = 2 Then
                                    var_referencia_caja = "0" + Trim(Str(var_numero_caja))
                                 End If
                                 If Len(Trim(Str(var_numero_caja))) = 3 Then
                                    var_referencia_caja = Trim(Str(var_numero_caja))
                                 End If
                                 If Len(Trim(Str(txt_embarque))) = 1 Then
                                    var_referencia_embarque = "00000" + Trim(Str(txt_embarque))
                                 End If
                                 If Len(Trim(Str(txt_embarque))) = 2 Then
                                    var_referencia_embarque = "0000" + Trim(Str(txt_embarque))
                                 End If
                                 If Len(Trim(Str(txt_embarque))) = 3 Then
                                    var_referencia_embarque = "000" + Trim(Str(txt_embarque))
                                 End If
                                 If Len(Trim(Str(txt_embarque))) = 4 Then
                                    var_referencia_embarque = "00" + Trim(Str(txt_embarque))
                                 End If
                                 If Len(Trim(Str(txt_embarque))) = 5 Then
                                    var_referencia_embarque = "0" + Trim(Str(txt_embarque))
                                 End If
                                 If Len(Trim(Str(txt_embarque))) = 6 Then
                                    var_referencia_embarque = Trim(Str(txt_embarque))
                                 End If
                                 var_contingencia = 1
                                 If var_contingencia = 1 Then
                                    VAR_ESTATUS = "S"
                                 Else
                                 rsaux6.Open "select * from tb_oracle_cajas_aduana where embarque = " + Me.txt_embarque + " and numero_caja = " + CStr(rsaux7!INTE_PAQ_CAJA), cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux6.EOF Then
                                    VAR_ESTATUS = IIf(IsNull(rsaux6!estatus), "", rsaux6!estatus)
                                 Else
                                    VAR_ESTATUS = ""
                                 End If
                                 rsaux6.Close
                                 End If
                                 If var_m = 1 Then
                                    var_cadena = "insert into TB_TEMP_ORACLE_CAJAS_ADUANA_DIVIDIDAS_EN_3 (inte_tem_consecutivo, renglon, pedido, caja_" + CStr(var_m) + ",codigo_" + CStr(var_m) + ", tipo_" + CStr(var_m) + ", sello_" + CStr(var_m) + ", auditada_" + CStr(var_m) + ",estatus_" + CStr(var_m) + ")"
                                    var_cadena = var_cadena + " values (" + CStr(var_consecutivo) + "," + CStr(var_j) + "," + CStr(rsaux7!SOURCE_HEADER_NUMBER) + "," + CStr(rsaux7!caja_pedido) + ",'C" + var_referencia_embarque + var_referencia_caja + "','" + rsaux7!tipo_caja + "','" + IIf(IsNull(rsaux7!sello), "", rsaux7!sello) + "'," + CStr(IIf(IsNull(rsaux7!auditada), 0, rsaux7!auditada)) + ",'" + VAR_ESTATUS + "')"
                                 Else
                                    var_cadena = "update TB_TEMP_ORACLE_CAJAS_ADUANA_DIVIDIDAS_EN_3 set caja_" + CStr(var_m) + " = " + CStr(rsaux7!caja_pedido) + ", codigo_" + CStr(var_m) + " = 'C" + var_referencia_embarque + var_referencia_caja + "',tipo_" + CStr(var_m) + " = '" + rsaux7!tipo_caja + "', sello_" + CStr(var_m) + "= '" + IIf(IsNull(rsaux7!sello), "", rsaux7!sello) + "', auditada_" + CStr(var_m) + "  = " + CStr(IIf(IsNull(rsaux7!auditada), 0, rsaux7!auditada)) + ", estatus_" + CStr(var_m) + " = '" + VAR_ESTATUS + "'  where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and renglon = " + CStr(var_j) + " and pedido = " + CStr(rsaux7!SOURCE_HEADER_NUMBER)
                                 End If
                                 'MsgBox var_cadena
                                 rsaux6.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux7.MoveNext
                        Wend
                        rsaux7.Close
                        rsaux9.MoveNext
                  Wend
                  
                        strconsulta = "SELECT INTE_EMB_EMBARQUE, COUNT(*) AS VECES FROM XXVIA_VW_BULTOS_AUDITADOS WHERE INTE_EMB_EMBARQUE = ? GROUP BY INTE_eMB_EMBARQUE"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
                             .Parameters.Append parametro
                        End With
                        Set rsaux8 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If Not rsaux8.EOF Then
                           var_total_bultos_embarques = IIf(IsNull(rsaux8!VECES), 0, rsaux8!VECES)
                        Else
                           var_total_bultos_embarques = 0
                        End If
                  
                        rsaux8.Close
                        rsaux8.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set TOTAL_BULTOS_AUDITODOS = " + CStr(var_total_bultos_embarques) + " where embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
                        
                  rs.Open "DELETE FROM TB_TEMP_ORACLE_REPORTE_CAJAS WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and ruta is null", cnn, adOpenDynamic, adLockOptimistic
                  Set reporte = appl.OpenReport(App.Path + "\rep_oracle_cajas_en_embarque_exportaciones.rpt")
                  reporte.RecordSelectionFormula = "{VW_ORACLE_REPORTE_CAJAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Pedidos cargados"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  rs.Open "DELETE FROM TB_TEMP_ORACLE_REPORTE_CAJAS WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               Else
                 MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
               End If
               If rsaux9.State = 1 Then
                  rsaux9.Close
               End If
               If rsaux11.State = 1 Then
                  rsaux11.Close
               End If
            Else
               MsgBox "El embarque aun no a sido cerrado", vbOKOnly, "ATENCION"
            End If
         Else
            
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
         If rsaux14.State = 14 Then
            rsaux14.Close
         End If
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
