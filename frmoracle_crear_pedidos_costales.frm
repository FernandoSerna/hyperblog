VERSION 5.00
Begin VB.Form frmoracle_crear_pedidos_costales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creación de pedidos de costales"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1245
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   3315
      Begin VB.TextBox txt_embarque 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1500
         TabIndex        =   2
         Top             =   330
         Width           =   1710
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Presione Enter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   1635
         TabIndex        =   3
         Top             =   870
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   135
         TabIndex        =   1
         Top             =   390
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmoracle_crear_pedidos_costales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub Form_Load()
   'Top = 3000
   'Left = 3800
   If var_embarque_costales > 0 Then
      Me.txt_embarque = var_embarque_costales
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_embarque) Then
         rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         strconsulta = "select * from xxvia_tb_encabezado_embarques where embarque = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
              .Parameters.Append parametro
         End With
         Set rsaux4 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rsaux4.EOF Then
            VAR_eSTATUS_ = IIf(IsNull(rsaux4!char_emb_estatus), "", rsaux4!char_emb_estatus)
            If VAR_eSTATUS_ <> "" Then
               strconsulta = "select distinct source_header_number from XXVIA_TB_SALIDAS_cAJAS where inte_emb_embarque = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                    .Parameters.Append parametro
               End With
               Set rsaux5 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               While Not rsaux5.EOF
                     strconsulta = "SELECT TIPO_CAJA, COUNT(*) AS CANTIDAD FROM XXVIA_VW_CAJAS_POR_PEDIDO WHERE SOURCE_HEADER_NUMBER = ? AND (TIPO_CAJA LIKE '%COSTAL%' OR TIPO_CAJA LIKE 'CAJA BIASI') GROUP BY TIPO_CAJA"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux5!source_header_number)
                          .Parameters.Append parametro
                     End With
                     Set rsaux6 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     
                     
                     If Not rsaux6.EOF Then
                        strconsulta = "select * from oe_order_headers_all where order_number = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rsaux5!source_header_number)
                             .Parameters.Append parametro
                        End With
                        Set rsaux9 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        var_posible_pedido = 1
                        If rsaux9!order_type_id = 1002 Then
                           var_posible_pedido = 0
                           var_pedido_tienda = IIf(IsNull(rsaux9!order_number), "", rsaux9!order_number)
                        End If
                        rsaux9.Close
                        If var_posible_pedido = 1 Then
                           strconsulta = "SELECT * FROM OE_ORDER_HEADERS_ALL WHERE ORIG_SYS_DOCUMENT_REF = ?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "SIDCOSTALES_" + Trim(CStr((rsaux5!source_header_number))))
                                .Parameters.Append parametro
                           End With
                           Set rsaux11 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If rsaux11.EOF Then
                              strconsulta = "SELECT  distinct hp.party_name as nombre_titular,  account_number as vcha_tit_titular_id, hcas.cust_account_id AS VCHA_CLI_CLAVE_ID FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, OE_ORDER_HEADERS_ALL OHA Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND hcas.cust_account_id = OHA.SOLD_TO_ORG_ID AND ORDER_NUMBER = ? ORDER BY hp.party_name"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Trim(CStr((rsaux5!source_header_number))))
                                   .Parameters.Append parametro
                              End With
                              Set rsaux12 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                              If Not rsaux12.EOF Then
                                 rsaux13.Open "SELECT * FROM TB_ORACLE_TITULARES_FACTURA_COSTALES WHERE TITULAR = '" + CStr(rsaux12!vcha_tit_titular_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux13.EOF Then
                                    var_posible_pedido = 1
                                 Else
                                    var_posible_pedido = 0
                                 End If
                                 rsaux13.Close
                              Else
                                 var_posible_pedido = 0
                              End If
                              rsaux12.Close
                              If var_posible_pedido = 1 Then
                              strconsulta = "SELECT SOLD_TO_ORG_ID AS TITULAR, SHIP_TO_ORG_ID AS ESTABLECIMIENTO, INVOICE_TO_ORG_ID AS CLIENTE, PRICE_LIST_ID AS LISTA_PRECIOS FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ?"
                              With comandoORA
                                  .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux5!source_header_number))
                                   .Parameters.Append parametro
                              End With
                              Set rsaux7 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                              
                              strconsulta = "select name from qp_secu_list_headers_v where list_header_id = ?"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rsaux7!LISTA_PRECIOS)
                                   .Parameters.Append parametro
                              End With
                              Set rsaux8 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                              var_lista_precios = rsaux8!Name
                              rsaux8.Close
                              var_clave_tipo_pedido = 1681
                              strconsulta = "INSERT INTO oe_headers_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref, creation_date, created_by, last_update_date, last_updated_by, operation_code , sold_to_org_id        , SHIP_TO_ORG_id                   ,INVOICE_TO_ORG_ID     , Order_type_ID, PRICE_LIST, SHIP_FROM_ORG_ID, attribute7)"
                              strconsulta = strconsulta + "  VALUES (1001,?,SYSDATE,-1,SYSDATE,-1,'INSERT', ?,?,?,?,?,?,?)"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "SIDCOSTALES_" + Trim(CStr((rsaux5!source_header_number))))
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux7!TITULAR)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux7!ESTABLECIMIENTO)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux7!Cliente)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_clave_tipo_pedido)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_lista_precios)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_unidad_organizacional)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "FACT. DE COSTALES")
                                   .Parameters.Append parametro
                              End With
                              Set rsaux8 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                           
                           
                           
                              var_i = 0
                              While Not rsaux6.EOF
                                    var_i = var_i + 1
                                    rs.Open "select * from tb_oracle_empaques where empaque = '" + rsaux6!tipo_caja + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rs.EOF Then
                                       strconsulta = "select PRIMARY_UOM_CODE, INVENTORY_ITEM_ID from xxvia_system_items_b where SEGMENT1 = ? AND ORGANIZATION_ID = ?"
                                       With comandoORA
                                            .ActiveConnection = cnnoracle_4
                                            .CommandType = adCmdText
                                            .CommandText = strconsulta
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rs!codigo)
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_unidad_organizacional)
                                            .Parameters.Append parametro
                                       End With
                                       Set rsaux8 = comandoORA.execute
                                       Set comandoORA = Nothing
                                       Set parametro = Nothing
                                       var_inventory_item_id = rsaux8!inventory_item_id
                                       VAR_MEDIDA = rsaux8!PRIMARY_UOM_CODE
                                       rsaux8.Close
                                    
                                    
                                       
                                       strconsulta = "INSERT INTO oe_lines_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref,orig_sys_line_ref,inventory_item_id,ordered_quantity, operation_code, created_by, creation_date, last_updated_by, last_update_date, unit_selling_price, unit_list_price, calculate_price_flag, PRICING_QUANTITY, PRICING_QUANTITY_UOM, ATTRIBUTE1, subinventory, org_id, ship_from_org_id)"
                                       strconsulta = strconsulta + " VALUES (1001,?,?,?, ?,'INSERT', -1,SYSDATE, -1,SYSDATE,0,0,'Y', ?, ?,'0','CDI_ALMPT',?,?)"
                                       With comandoORA
                                            .ActiveConnection = cnnoracle_4
                                            .CommandType = adCmdText
                                            .CommandText = strconsulta
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "SIDCOSTALES_" + Trim(CStr((rsaux5!source_header_number))))
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_i)
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_inventory_item_id)
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux6!cantidad)
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux6!cantidad)
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, VAR_MEDIDA)
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_empresa)
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_unidad_organizacional)
                                            .Parameters.Append parametro
                                       End With
                                       Set rsaux8 = comandoORA.execute
                                       Set comandoORA = Nothing
                                       Set parametro = Nothing
                                    End If
                                    rs.Close
                                    rsaux6.MoveNext
                              Wend
                              On Error GoTo salir2
                              rsaux8.Open "INSERT INTO oe_actions_iface_all (order_source_ID, orig_sys_document_ref, operation_code) VALUES (1001, 'SIDCOSTALES_" + Trim(CStr((rsaux5!source_header_number))) + "','BOOK_ORDER')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              rsaux8.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              rsaux8.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              rsaux8.Open "CALL XXVIA_PK_INTERFACES_OM.importar_pedido('SIDCOSTALES_" + Trim(CStr((rsaux5!source_header_number))) + "'," + var_empresa + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              End If
                           End If
                           rsaux11.Close
                           
                        Else
                           ' traspaso para tiendas
                           var_posible_pedido = 1
                           rsaux8.Open "SELECT * FROM TB_ORACLE_PEDIDOS_TIENDAS_COSTALES WHERE PEDIDO = " + CStr(var_pedido_tienda), cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux8.EOF Then
                              var_posible_pedido = 0
                           End If
                           rsaux8.Close
                           If var_posible_pedido = 1 Then
                              strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B,  OE_ORDER_HEADERS_ALL OHA Where requisition_header_id = OHA.SOURCE_DOCUMENT_ID AND secondary_inventory_name = A.ATTRIBUTE1 AND OHA.ORDER_NUMBER = ?"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_pedido_tienda)
                                   .Parameters.Append parametro
                              End With
                              Set rsaux8 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                              var_almacen_tienda = IIf(IsNull(rsaux8!attribute1), "", rsaux8!attribute1)
                              p_almacendestinofinal = var_almacen_tienda
                              rsaux8.Close
                              If var_almacen_tienda <> "" Then
                                 var_i = 0
                                 rsaux8.Open "SELECT XXVIA_SQ_LINEA_TM.nextval FROM dual", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux8.EOF Then
                                    p_origenencabezadoid = rsaux8(0).Value
                                 End If
                                 rsaux8.Close
                                 rsaux8.Open "select XXVIA_SQ_ENCABEZADO_MT_ID.nextval from dual", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux8.EOF Then
                                    P_ENCABEZADO_MT_ID = rsaux8(0).Value
                                 End If
                                 rsaux8.Close
                                 While Not rsaux6.EOF
                                       var_i = var_i + 1
                                       rs.Open "select * from tb_oracle_empaques where empaque = '" + rsaux6!tipo_caja + "'", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rs.EOF Then
                                          strconsulta = "select PRIMARY_UOM_CODE, INVENTORY_ITEM_ID from xxvia_system_items_b where SEGMENT1 = ? AND ORGANIZATION_ID = ?"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rs!codigo)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_unidad_organizacional)
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux8 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          var_inventory_item_id = rsaux8!inventory_item_id
                                          p_um = rsaux8!PRIMARY_UOM_CODE
                                          rsaux8.Close
                                          p_organizacion_id = var_unidad_organizacional
                                          p_organizacion_destino = var_unidad_organizacional
                                          If var_empresa = 92 Then
                                             p_subinventario = "CDI_ALMPT"
                                          End If
                                          If var_empresa = 83 Then
                                             p_subinventario = "CDISTEX_PT"
                                          End If
                                          p_subinventario_destino = "TRANS"
                                          p_codigoarticulo = rs!codigo
                                          p_cantidadorigen = rsaux6!cantidad
                                          p_Cantidadrecibida = 0
                                          p_origentransaccion = "SID_COSTALES_" + CStr(rsaux5!source_header_number)
                                          p_referencia_transaccion = rsaux5!source_header_number
                                          p_mensajeerror = ""
                                          strconsulta = "call xxvia_pk_inventarios.xxvia_sp_inventarios4 (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adInteger, adParamInput, 100, p_organizacion_id)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adInteger, adParamInput, 100, p_organizacion_destino)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_subinventario)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_subinventario_destino)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adInteger, adParamInput, 100, 2)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_codigoarticulo)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adInteger, adParamInput, 100, p_cantidadorigen)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adInteger, adParamInput, 100, p_cantidadorigen)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_origentransaccion)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adInteger, adParamInput, 100, p_origenencabezadoid)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adDouble, adParamInput, 100, Null)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adDouble, adParamInput, 100, Null)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_referencia_transaccion)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adDouble, adParamInput, 100, Null)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_almacendestinofinal)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adDate, adParamInput, 100, Date)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Null)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_um)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adDouble, adParamInput, 100, P_ENCABEZADO_MT_ID)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_clave_usuario_global)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, fun_NombrePc)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Null)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "")
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "")
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamOutput, 100, p_mensajeerror)
                                               .Parameters.Append parametro
                                          End With
                                          'MsgBox strconsulta
                                          Set rsaux9 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          
                                       End If
                                       rs.Close
                                       rsaux6.MoveNext
                                 Wend
                                 strconsulta = "call xxvia_pk_inventarios.xxvia_valida_interface (1,?,?)"
                                 With comandoORA
                                      .ActiveConnection = cnnoracle_4
                                      .CommandType = adCmdText
                                      .CommandText = strconsulta
                                      Set parametro = .CreateParameter(, adDouble, adParamInput, 100, p_origenencabezadoid)
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adDouble, adParamInput, 200, 0)
                                      .Parameters.Append parametro
                                 End With
                                 rsaux9.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 On Error GoTo salir2
                                 Set rsaux9 = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing
                                 rsaux8.Open "insert into tb_oracle_pedidos_tiendas_costales (pedido) values (" + CStr(rsaux5!source_header_number) + ")", cnn, adOpenDynamic, adLockOptimistic
                              End If
                           End If
                        End If
                     End If
                     rsaux6.Close
                     
                     rsaux5.MoveNext
               Wend
               rsaux5.Close
               MsgBox "Se a terminado el proceso"
               Unload Me
            End If
         Else
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
         rsaux4.Close
      End If
   Else
      'KeyAscii = 0
   End If
   Exit Sub
salir2:
   If Err.Number = -2147217900 Then
      rsaux10.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux10.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      'MsgBox Err.Description
      Resume
   Else
      MsgBox Err.Description
      Resume
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
   MsgBox "No se pudo generar el documento electrónico", vbOKOnly, "ATENCION"
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
End Sub

Private Sub txt_embarque_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Shift = 6 And KeyCode = 111 Then
       
   End If
End Sub
