VERSION 5.00
Begin VB.Form frmoracle_subir_pedidos_contingencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subir pedidos por contingencia"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_cargar_pedidos 
      Caption         =   "Cargar pedidos"
      Height          =   690
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   4320
   End
End
Attribute VB_Name = "frmoracle_subir_pedidos_contingencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cargar_pedidos_Click()
   Dim var_inserta As Boolean
   Dim var_factura As Integer
   Dim var_posible_cliente As Boolean
'On Error GoTo salir:
   rsaux11.Open "DELETE FROM tb_temp_oracle_pedidos_subir", cnn, adOpenDynamic, adLockOptimistic
   strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls); DBQ=C:\REPORTESSID\pedido.XLS"
   rsaux11.Open "select codigo as segment1, remision, cliente, establecimiento, titular, cantidad  FROM [pedido$] ", strConnectionString
   var_primera_vez = True
   While Not rsaux11.EOF
         txt_codigo = rsaux11!SEGMENT1
         var_cantidad_leida = rsaux11!Cantidad
         var_cadena = "insert into tb_temp_oracle_pedidos_subir (remision, codigo, cantidad, cliente, establecimiento, titular, estatus)"
         var_cadena = var_cadena + " values ('" + rsaux11!REMISION + "','000" + CStr(rsaux11!SEGMENT1) + "'," + CStr(rsaux11!Cantidad) + ",'" + UCase(rsaux11!Cliente) + "','" + UCase(rsaux11!ESTABLECIMIENTO) + "','" + UCase(rsaux11!TITULAR) + "',0)"
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         txt_codigo = ""
         rsaux11.MoveNext
   Wend
   rsaux11.Close
   rsaux11.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   rsaux11.Open "select distinct CODIGO AS SEGMENT1 from tb_temp_oracle_pedidos_subir", cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux11.EOF
         txt_codigo = rsaux11!SEGMENT1
         If Trim(txt_codigo) <> "" Then
            rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
               var_unidad_medida = rsaux8!PRIMARY_UOM_CODE
               var_descripcion_articulo = rsaux8!Description
               var_inventory_item_id = rsaux8!INVENTORY_ITEM_ID
            End If
            rsaux9.Open "update tb_temp_oracle_pedidos_subir set UNIDAD_MEDIDA = '" + var_unidad_medida + "', DESCRIPCION = '" + var_descripcion_articulo + "', INVENTORY_ITEM_ID = " + CStr(var_inventory_item_id) + " WHERE CODIGO = '" + rsaux11!SEGMENT1 + "'", cnn, adOpenDynamic, adLockOptimistic
            rsaux8.Close
         End If
         rsaux11.MoveNext
   Wend
   rsaux11.Close
   var_clave_movimiento = "VDI"
   If var_clave_movimiento = "VDI" Then
      rsaux11.Open "SELECT DISTINCT REMISION, TITULAR, CLIENTE, ESTABLECIMIENTO FROM tb_temp_oracle_pedidos_subir", cnn, adOpenDynamic, adLockOptimistic
      While Not rsaux11.EOF
            rsaux12.Open "SELECT  DISTINCT hp.party_name as nombre_titular,  account_number as vcha_tit_titular_id, hcas.cust_account_id AS VCHA_CLI_CLAVE_ID FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors Arc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id  AND account_number = '" + rsaux11!TITULAR + "' ORDER BY hp.party_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
            VAR_TITULAR = rsaux12!VCHA_CLI_CLAVE_ID
            rsaux12.Close
            rsaux12.Open "SELECT  hcp.site_use_id AS VCHA_CLI_CLAVE_ID FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO'  and hcas.cust_account_id = " + CStr(VAR_TITULAR) + " AND hps.party_site_number = '" + rsaux11!Cliente + "' ORDER BY hl.address1", cnnoracle_4, adOpenDynamic, adLockOptimistic
            VAR_CLIENTE = rsaux12!VCHA_CLI_CLAVE_ID
            rsaux12.Close
            rsaux12.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id  = ps.party_site_id AND ps.location_id  = lo.location_id  AND csu.site_use_code = 'SHIP_TO' AND cas.cust_account_id     = " + CStr(VAR_TITULAR) + " AND ps.party_site_number = '" + rsaux11!ESTABLECIMIENTO + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            VAR_ESTABLECIMIENTO = rsaux12!VCHA_ESB_ESTABLECIMIENTO_ID
            rsaux12.Close
            var_cadena = "SELECT  hcsu.price_list_id, hcsu.order_type_id,hca.cust_account_id, hcp.site_use_id AS VCHA_CLI_CLAVE_ID,hl.address1 VCHA_CLI_NOMBRE FROM hz_parties hp,hz_party_sites hps,hz_cust_accounts hca,hz_cust_acct_sites_all hcas,hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND  hcp.site_use_id = " + CStr(VAR_CLIENTE)
            rsaux10.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux10.EOF Then
               var_clave_lista_precios = rsaux10!price_list_id
            Else
               var_clave_lista_precios = 0
            End If
            rsaux10.Close
            var_clave_tipo_pedido = 1042
            If var_clave_tipo_pedido > 0 Then
               VAR_LISTA_PRECIOS = var_clave_lista_precios
               If VAR_LISTA_PRECIOS <> "" Then
                  If rs.State = 1 Then
                     rs.Close
                  End If
                  rs.Open "SELECT * FROM tb_temp_oracle_pedidos_subir WHERE REMISION = '" + rsaux11!REMISION + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     rsaux7.Open "select name from qp_secu_list_headers_v where list_header_id = " + CStr(var_clave_lista_precios), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     VAR_LISTA_PRECIOS = rsaux7(0).Value
                     rsaux7.Close
                              
                     var_cadena = "INSERT INTO oe_headers_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref, creation_date, created_by, last_update_date, last_updated_by, operation_code , sold_to_org_id        , SHIP_TO_ORG_id                   ,INVOICE_TO_ORG_ID     , Order_type_ID, PRICE_LIST, org_id, ship_from_org_id)"
                     var_cadena = var_cadena + "  VALUES (1001,'SIDREM_" + rs!REMISION + "',SYSDATE,-1,SYSDATE, -1,'INSERT', " + CStr(VAR_TITULAR) + "," + CStr(VAR_ESTABLECIMIENTO) + "," + CStr(VAR_CLIENTE) + "," + CStr(var_clave_tipo_pedido) + ",'" + VAR_LISTA_PRECIOS + "',92,93)"
                     If rsaux.State = 1 Then
                        rsaux.Close
                     End If
                     rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_i = 0
                     While Not rs.EOF
                           var_i = var_i + 1
                           rsaux10.Open "SELECT PRIMARY_UOM_CODE FROM xxvia_system_items_b WHERE INVENTORY_ITEM_ID = " + CStr(rs!INVENTORY_ITEM_ID) + " AND ORGANIZATION_ID = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux10.EOF Then
                                  VAR_MEDIDA = rsaux10(0).Value
                           End If
                           rsaux10.Close
                           
                           var_cadena = "INSERT INTO oe_lines_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref,orig_sys_line_ref,inventory_item_id,ordered_quantity, operation_code, created_by, creation_date, last_updated_by, last_update_date, unit_selling_price, unit_list_price, calculate_price_flag, PRICING_QUANTITY, PRICING_QUANTITY_UOM, ATTRIBUTE1, subinventory, org_id, ship_from_org_id)"
                           var_cadena = var_cadena + " VALUES (1001,'SIDREM_" + rs!REMISION + "','" + CStr(var_i) + "', " + CStr(rs!INVENTORY_ITEM_ID) + ", " + CStr(rs!Cantidad) + ",'INSERT', -1,SYSDATE, -1,SYSDATE,0,0,'Y', " + CStr(rs!Cantidad) + ", '" + VAR_MEDIDA + "','','CDI_ALMPT',92,93)"
                           rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           rs.MoveNext
                     Wend
                     On Error GoTo salir
                     rs.MoveFirst
                     rsaux.Open "INSERT INTO oe_actions_iface_all (order_source_ID, orig_sys_document_ref, operation_code) VALUES (1001, 'SIDREM_" + rs!REMISION + "','BOOK_ORDER')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     rsaux.Open "CALL XXVIA_PK_INTERFACES_OM.importar_pedido('SIDREM_" + rs!REMISION + "'," + var_empresa + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If rsaux.State = 1 Then
                         rsaux.Close
                     End If
                     rsaux.Open "select order_number from oe_order_headers_all where orig_sys_document_ref = 'SIDREM_" + rs!REMISION + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_pedido = rsaux(0).Value
                     rsaux.Close
                     MsgBox var_pedido
                  End If
                  rs.Close
               Else
                  MsgBox "No se a indicado una lista de precios", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No se a indicado un tipo de pedido", vbOKOnly, "ATENCION"
            End If
            rsaux11.MoveNext
      Wend
      rsaux11.Close
   End If
Exit Sub
salir:
   If Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
       rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      Resume
   Else
      MsgBox Err.Description
     ' Resume
      If rs.State = 1 Then
         rs.Close
      End If
      If rsaux.State = 1 Then
         rsaux.Close
      End If
      If rsaux1.State = 1 Then
         rsaux1.Close
      End If
      If rsaux2.State = 1 Then
         rsaux2.Close
      End If
      If rsaux3.State = 1 Then
         rsaux3.Close
      End If
      If rsaux4.State = 1 Then
         rsaux4.Close
      End If
      If rsaux5.State = 1 Then
         rsaux5.Close
      End If
      If rsaux6.State = 1 Then
         rsaux6.Close
      End If
      If rsaux7.State = 1 Then
         rsaux7.Close
      End If
      If rsaux8.State = 1 Then
         rsaux8.Close
      End If
      If rsaux9.State = 1 Then
         rsaux9.Close
      End If
   End If
   Exit Sub
salir_factura:
   MsgBox "Surgio un error al generar los documentos electrónicos", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux1.State = 1 Then
      rsaux1.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   If objConn.State = 1 Then
      objConn.RollbackTrans
      objConn.Close
   End If
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub
