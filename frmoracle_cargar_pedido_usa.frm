VERSION 5.00
Begin VB.Form frmoracle_cargar_pedido_usa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cargar pedido USA"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_embarque 
      Height          =   555
      Left            =   120
      TabIndex        =   2
      Top             =   75
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox txt_pedido 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1200
      TabIndex        =   1
      Top             =   90
      Width           =   2385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pedido:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "frmoracle_cargar_pedido_usa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub Form_Load()
   Top = 3200
   Left = 4200
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_pedido) Then
         If rsaux1.State = 1 Then
            rsaux1.Close
         End If
         rsaux1.Open "select * from tb_pedidos_usa_cargados where pedido = " + Me.txt_pedido, cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux1.EOF Then
            var_posible = 0
         Else
            var_posible = 1
         End If
         rsaux1.Close
         var_posible = 1
         If var_posible = 1 Then
            If cnn_icg_usa.State = 1 Then
               cnn_icg_usa.Close
            End If
            cnn_icg_usa.Open "Provider=SQLOLEDB.1;Password=ICGUsa2014;Persist Security Info=True;User ID=ICGAdmin;Initial Catalog=bd1;Data Source=sqlcedishou.VIANNEYcatalog.COM"
            var_cadena = "SELECT distinct inte_emb_embarque, source_header_number from xxvia_tb_salidas_cajas where source_header_number = ?"
            'MsgBox var_cadena
            strconsulta = var_cadena
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, Me.txt_pedido)
                 .Parameters.Append parametro
            End With
            Set rsaux12 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            While Not rsaux12.EOF
                  Me.txt_embarque = IIf(IsNull(rsaux12!inte_Emb_Embarque), "", rsaux12!inte_Emb_Embarque)
                  If rsaux1.State = 1 Then
                     rsaux1.Close
                  End If
                  rsaux1.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                   var_cadena = "SELECT b.vcha_tit_titular_id, D.vcha_esb_establecimient_id FROM XXVIA_VW_CLIENTES_PEDIDOS B, OE_ORDER_HEADERS_ALL C, XXVIA_VW_ESTABLECIMIENTOS_PED D, XXVIA_VW_ESTABLECIMIENTOS_PED E Where c.order_number = ? AND c.SOLD_TO_ORG_ID = B.CUST_ACCOUNT_ID AND D.SITE_USE_ID    = C.SHIP_TO_ORG_ID AND E.SITE_USE_ID    = C.INVOICE_TO_ORG_ID"
                  'MsgBox var_cadena
                   strconsulta = var_cadena
                   With comandoORA
                        .ActiveConnection = cnnoracle_4
                        .CommandType = adCmdText
                        .CommandText = strconsulta
                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_pedido)
                        .Parameters.Append parametro
                   End With
                   Set rsaux9 = comandoORA.execute
                   Set comandoORA = Nothing
                   Set parametro = Nothing
                   var_Titular_vianney_catalog = rsaux9!vcha_tit_titular_id
                   var_establecimiento = rsaux9!vcha_esb_establecimient_id
                   var_si_vianney_Catalog = 0
                   var_almacen_Destino = rsaux9!vcha_esb_establecimient_id
                   var_almacen_Destino = "E000005087"
                   var_establecimiento = "E000005087"
                   If var_Titular_vianney_catalog = "T000000343" Or var_Titular_vianney_catalog = "000012364" Then
                      var_si_vianney_Catalog = 1
                   End If
                   
                   var_dia = CStr(Day(CDate(Date)))
                   var_mes = CStr(Month(CDate(Date)))
                   var_año = CStr(Year(CDate(Date)))
                   If Len(Trim(var_dia)) = 1 Then
                      var_dia = "0" + var_dia
                   End If
                   If Len(Trim(var_mes)) = 1 Then
                      var_mes = "0" + var_mes
                   End If
                   If Len(Trim(var_hora)) = 1 Then
                      var_hora = "0" + var_hora
                   End If
                   var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                   If var_si_vianney_Catalog = 1 Then
                      'var_cadena = "SELECT sum(floa_sal_Cantidad_leida) as cantidad, organizacion,  inte_emb_embarque as embarque, inte_paq_caja as caja, source_header_number  as pedido, a.segment1 as codigo, collector_id as agente,name as nombre_agente, customer_id as cliente, customer_name as nombre_cliente, a.inventory_item_id, caja_pedido, sello, UNIT_WEIGHT as peso,  item_description as descripcion    FROM XXVIA_TB_salidas_cajas a, xxvia_tb_encabezado_embarques, xxvia_system_items_b b, oe_order_headers_all oh where inte_emb_embarque = embarque and organizacion = b.organization_id and a.inventory_item_id = b.inventory_item_id and order_number = a.source_header_number and oh.ship_from_org_id = organizacion and inte_emb_embarque = " + Me.txt_embarque + " and floa_sal_Cantidad_leida >0 AND a.source_header_number = " + CStr(rsaux12(0).Value) + " group by "
                      'var_cadena = var_cadena + " organizacion,  inte_emb_embarque as embarque, inte_paq_caja as caja, source_header_number  as pedido, a.segment1 as codigo, collector_id as agente,name as nombre_agente, customer_id as cliente, customer_name as nombre_cliente, a.inventory_item_id, caja_pedido, sello, UNIT_WEIGHT as peso,  item_description as descripcion"
                        
                      var_cadena = "SELECT SUM(floa_sal_Cantidad_leida) AS cantidad, organizacion, inte_emb_embarque    AS embarque, inte_paq_caja AS caja, source_header_number AS pedido, a.segment1 AS codigo, collector_id         AS agente, name                 AS nombre_agente, customer_id          AS cliente, customer_name        AS nombre_cliente, a.inventory_item_id, caja_pedido, sello, UNIT_WEIGHT      AS peso, item_description As descripcion FROM XXVIA_TB_salidas_cajas a, xxvia_tb_encabezado_embarques, xxvia_system_items_b b, oe_order_headers_all oh Where inte_emb_embarque = Embarque AND organizacion            = b.organization_id AND a.inventory_item_id     = b.inventory_item_id AND order_number            = a.source_header_number AND oh.ship_from_org_id     = organizacion AND source_header_number       = " + Me.txt_pedido + " "
                      var_cadena = var_cadena + " AND floa_sal_Cantidad_leida >0 AND a.source_header_number  = " + CStr(rsaux12!source_header_number) + " GROUP BY organizacion, inte_emb_embarque    , inte_paq_caja        , source_header_number , a.segment1           , collector_id, name, customer_id, customer_name, a.inventory_item_id, caja_pedido, sello, UNIT_WEIGHT, item_description  order by caja_pedido, codigo "
                      If rs.State = 1 Then
                         rs.Close
                      End If
                      rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                      If Not rs.EOF Then
                         strconsulta = "select SECONDARY_INVENTORY_NAME from mtl_secondary_inventories where ATTRIBUTE8 = ?"
                         With comandoORA
                              .ActiveConnection = cnnoracle_4
                              .CommandType = adCmdText
                              .CommandText = strconsulta
                              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen_Destino)
                              .Parameters.Append parametro
                         End With
                         Set rsaux = comandoORA.execute
                         Set comandoORA = Nothing
                         Set parametro = Nothing
                         If Not rsaux.EOF Then
                            var_almacen_icg = rsaux(0).Value
                         Else
                            var_almacen_icg = ""
                         End If
                         'VAR_ESTABLECIMIENTO = "E000005087"
                         If var_establecimiento = "E000005087" Or var_establecimiento = "" Then
                            var_almacen_icg = "VCA_TD6013"
                         End If
                         
                         If var_establecimiento = "000014448" Then
                            var_almacen_icg = "VCA_TD6819"
                         End If
                         
                         If var_establecimiento = "000017009" Then
                            var_almacen_icg = "CR_PAVAS"
                         End If
                         If var_establecimiento = "000016989" Then
                            var_almacen_icg = "CR_COLON"
                         End If
                         'GoTo fer:
                         rsaux.Close
                         var_almacen_icg = "VCA_TD6013"
                         rsaux.Open "alter session set nls_language= 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                         While Not rs.EOF
                               var_cadena = "SELECT * FROM OPENQUERY(ICGCENTRAL, 'SELECT * FROM BDICGDESA_USA.DBO.IT_PEDIDOCOMPRA where no_embarque = ''" + Me.txt_embarque + "'' and no_pedido = ''" + CStr(rs!pedido) + "'' and no_caja =  ''" + CStr(rs!Caja) + "'' and codigo = ''" + rs!codigo + "''')A"
                               If rsaux10.State = 1 Then
                                  rsaux10.Close
                               End If
                              'MsgBox cnn_icg_usa.ConnectionString
                               '|rsaux10.Open "SELECT * FROM BDICGDESA_USA.DBO.IT_PEDIDOCOMPRA", cnn_icg_usa, adOpenDynamic, adLockOptimistic
                               'If cnn_icg_usa.State = 1 Then
                               '   cnn_icg_usa.Close
                               'End If
                               'cnn_icg_usa.Open "Provider=SQLOLEDB.1;Password=ICGUsa2014;Persist Security Info=True;User ID=ICGAdmin;Initial Catalog=bd1;Data Source=sqlcedishou.VIANNEYcatalog.COM"
                               'MsgBox cnn_icg_usa
                               rsaux10.Open var_cadena, cnn_icg_usa, adOpenDynamic, adLockOptimistic
                               If rsaux10.EOF Then
                                  var_pedido = rs!pedido
                                  strconsulta = "select unit_selling_price from oe_order_headers_all oh, oe_order_lines_all ol where order_number = ? and oh.header_id = ol.header_id and oh.ship_from_org_id = ? and ol.inventory_item_id = ?"
                                  With comandoORA
                                       .ActiveConnection = cnnoracle_4
                                       .CommandType = adCmdText
                                       .CommandText = strconsulta
                                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_pedido))
                                       .Parameters.Append parametro
                                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
                                       .Parameters.Append parametro
                                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, IIf(IsNull(rs!inventory_item_id), 0, rs!inventory_item_id))
                                       .Parameters.Append parametro
                                  End With
                                  Set rsaux = comandoORA.execute
                                  Set comandoORA = Nothing
                                  Set parametro = Nothing
                            
                                  If Not rsaux.EOF Then
                                     var_precio = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                                  Else
                                     var_precio = 0
                                  End If
                                  rsaux.Close
                                  x = 0
                                  If x = 0 Then
                                  rsaux11.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                  If var_Titular_vianney_catalog = "000012364" Then
                                     rsaux11.Open "insert into IT_PEDIDOCOMPRA (OU, SUBINVENTORY_CODE, TRANSFER_SUBINVENTORY, FECHA, NO_EMBARQUE, NO_PEDIDO, NO_CAJA, CODIGO, CANTIDAD, PRECIO, DESCRIPCION) values ('10','" + var_almacen_icg + "','CDI_ALMPT', " + var_fecha + ", " + Me.txt_embarque + ",'" + CStr(rs!pedido) + "','" + CStr(rs!Caja) + "','" + rs!codigo + "'," + CStr(rs!cantidad) + "," + CStr(var_precio) + ",'" + rs!Descripcion + "')", cnn, adOpenDynamic, adLockOptimistic
                                  Else
                                     rsaux11.Open "insert into IT_PEDIDOCOMPRA (OU, SUBINVENTORY_CODE, TRANSFER_SUBINVENTORY, FECHA, NO_EMBARQUE, NO_PEDIDO, NO_CAJA, CODIGO, CANTIDAD, PRECIO, DESCRIPCION) values ('361','" + var_almacen_icg + "','CDI_ALMPT', " + var_fecha + ", " + Me.txt_embarque + ",'" + CStr(rs!pedido) + "','" + CStr(rs!Caja) + "','" + rs!codigo + "'," + CStr(rs!cantidad) + "," + CStr(var_precio) + ",'" + rs!Descripcion + "')", cnn, adOpenDynamic, adLockOptimistic
                                  End If
                                  End If
                               End If
                               rsaux10.Close
                               rs.MoveNext
                         Wend
                         rs.Close
                         
fer:
                         cnn_icg_usa.CommandTimeout = 360
                         If rsaux1.State = 1 Then
                            rsaux1.Close
                         End If
                         x = 1
                         If x = 1 Then
                            rsaux1.Open "select OU, SUBINVENTORY_CODE, TRANSFER_SUBINVENTORY, FECHA, NO_EMBARQUE, NO_PEDIDO, NO_CAJA, CODIGO, CANTIDAD, PRECIO, DESCRIPCION,3 from IT_PEDIDOCOMPRA where NO_EMBARQUE = " + Me.txt_embarque + " AND NO_PEDIDO = " + CStr(Me.txt_pedido), cnn, adOpenDynamic, adLockOptimistic
                            While Not rsaux1.EOF
                                  var_dia = CStr(Day(CDate(rsaux1!Fecha)))
                                  var_mes = CStr(Month(CDate(rsaux1!Fecha)))
                                  var_año = CStr(Year(CDate(rsaux1!Fecha)))
                                  If Len(Trim(var_dia)) = 1 Then
                                     var_dia = "0" + var_dia
                                  End If
                                  If Len(Trim(var_mes)) = 1 Then
                                     var_mes = "0" + var_mes
                                  End If
                                  If Len(Trim(var_año)) = 2 Then
                                     var_año = "20" + var_año
                                  End If
                                  var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                  If rs.State = 1 Then
                                     rs.Close
                                  End If
                                  'If VAR_ESTABLECIMIENTO = "E000005087" Or VAR_ESTABLECIMIENTO = "000014448" Or VAR_ESTABLECIMIENTO = "000017009" Or VAR_ESTABLECIMIENTO = "000016989" Then
                                     var_cadena = "(" + CStr(rsaux1!OU) + ", 'VCA_TD6013', '" + rsaux1!TRANSFER_SUBINVENTORY + "', " + var_fecha + ", " + CStr(rsaux1!NO_EMBARQUE) + ", '" + CStr(rsaux1!NO_PEDIDO) + "', '" + CStr(rsaux1!NO_CAJA) + "', '" + Trim(CStr(rsaux1!codigo)) + "', " + CStr(rsaux1!cantidad) + ", " + CStr(rsaux1!Precio) + ", '" + CStr(rsaux1!Descripcion) + "',3)"
                                     rs.Open "INSERT INTO ICGCENTRAL.BDICGDESA_USA.DBO.IT_PEDIDOCOMPRA (OU, SUBINVENTORY_CODE, TRANSFER_SUBINVENTORY, FECHA, NO_EMBARQUE, NO_PEDIDO, NO_CAJA, CODIGO, CANTIDAD, PRECIO, DESCRIPCION,status) values  " + var_cadena, cnn_icg_usa, adOpenDynamic, adLockOptimistic
                                  'Else
                                     'var_cadena = "(" + CStr(rsaux1!OU) + ", '" + rsaux1!SUBINVENTORY_CODE + "', '" + rsaux1!TRANSFER_SUBINVENTORY + "', " + var_fecha + ", " + CStr(rsaux1!NO_EMBARQUE) + ", '" + CStr(rsaux1!NO_PEDIDO) + "', '" + CStr(rsaux1!NO_CAJA) + "', '" + Trim(CStr(rsaux1!codigo)) + "', " + CStr(rsaux1!cantidad) + ", " + CStr(rsaux1!Precio) + ", 3)"
                                     'rs.Open "INSERT INTO ICGCENTRAL.GENERAL.DBO.IT_PEDIDOCOMPRADIRECTA (numb_organization_id, vcha_SUBINVENTORY_CODE, vcha_TRANSFER_SUBINVENTORY, date_FECHA, vcha_doc_interno_icg, vcha_NOta_envio, vcha_numero_CAJA, vcha_CODIGO, numb_CANTIDAD, PRECIO, numb_status) values  " + var_cadena, cnn_icg_usa, adOpenDynamic, adLockOptimistic
                                  
                                  'end If
                                  
                                  'MsgBox var_cadena
                                  If rs.State = 1 Then
                                     rs.Close
                                  End If
                                  rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                  
                                  If var_establecimiento = "E000005087" Or var_establecimiento = "000014448" Or var_establecimiento = "000017009" Or var_establecimiento = "000016989" Then
                                     
                                  Else
                                     
                                  End If
                                  rsaux1.MoveNext
                            Wend
                            'If VAR_ESTABLECIMIENTO = "E000005087" Or VAR_ESTABLECIMIENTO = "000014448" Then
                               var_pedido = Me.txt_pedido
                               'If cnn_icg_usa.State = 1 Then
                               '   cnn_icg_usa.Close
                               'End If
                               'cnn_icg_usa.Open "Provider=SQLOLEDB.1;Password=ICGUsa2014;Persist Security Info=True;User ID=ICGAdmin;Initial Catalog=bd1;Data Source=sqlcedishou.VIANNEYcatalog.COM"
                               cnn_icg_usa.CommandTimeout = 600
                               
                               rs.Open "update ICGCENTRAL.BDICGDESA_USA.DBO.IT_PEDIDOCOMPRA set status = 0,log= ''  where NO_EMBARQUE = " + Me.txt_embarque + " AND NO_PEDIDO = " + CStr(var_pedido), cnn_icg_usa, adOpenDynamic, adLockOptimistic
                               
                               rs.Open "EXEC [GENERAL].DBO.[vyt_caller_OC] 361", cnn_icg_usa, adOpenDynamic, adLockOptimistic
                               rs.Open "EXEC [GENERAL].DBO.[vyt_caller_OC] 10", cnn_icg_usa, adOpenDynamic, adLockOptimistic
                            'Else
                               'rs.Open "update ICGCENTRAL.GENERAL.DBO.IT_PEDIDOCOMPRADIRECTA set numb_status = 0 where vcha_NOta_envio = " + Me.txt_pedido, cnn_icg_usa, adOpenDynamic, adLockOptimistic
                               'rs.Open "EXEC ICGCENTRAL.general.DBO.[vyt_crea_pedido_cedis] 361," + Me.txt_pedido, cnn_icg_usa, adOpenDynamic, adLockOptimistic
                            'End If
                         Else
                            'rs.Open "INSERT OPENQUERY (icgcentral, 'SELECT OU, SUBINVENTORY_CODE, TRANSFER_SUBINVENTORY, FECHA, NO_EMBARQUE, NO_PEDIDO, NO_CAJA, CODIGO, CANTIDAD, PRECIO, DESCRIPCION FROM SQLQUEZADA2.SIDAlmacenBkp.DBO.it_pedidocompra') select OU, SUBINVENTORY_CODE, TRANSFER_SUBINVENTORY, FECHA, NO_EMBARQUE, NO_PEDIDO, NO_CAJA, CODIGO, CANTIDAD, PRECIO, DESCRIPCION from IT_PEDIDOCOMPRA where NO_EMBARQUE = " + Me.txt_embarque + " AND NO_PEDIDO = " + CStr(var_pedido), cnnicg_sql, adOpenDynamic, adLockOptimistic
                         End If
                      End If
                   End If
                   
                   var_conexion_string = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & parametros(1) & ";Data Source=" & parametros(0)
                   var_bd_reportes = "SIDAlmacenbkp"
                   If cnn.State = 1 Then
                      cnn.Close
                   End If
                   cnn.Open var_conexion_string
                   
                   rsaux13.Open "insert into tb_pedidos_usa_cargados (pedido) values (" + CStr(Me.txt_pedido) + ")", cnn, adOpenDynamic, adLockOptimistic
                   rs.Open "EXEC [GENERAL].DBO.[vyt_caller_OC] 361", cnn_icg_usa, adOpenDynamic, adLockOptimistic
                   rsaux12.MoveNext
            Wend
            rsaux12.Close
            MsgBox "Se a terminado el proceso de carga", vbOKOnly, "ATENCION"
         Else
            MsgBox "El pedido ya fue cargado con anterioridad", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub
