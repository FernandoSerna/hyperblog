VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmoracle_packing_list 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Packing list"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "Compreso"
      Height          =   330
      Left            =   1710
      Picture         =   "frmoracle_packing_list.frx":0000
      TabIndex        =   10
      ToolTipText     =   "Compreso"
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmd_facturas_por_embarque 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1380
      Picture         =   "frmoracle_packing_list.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Archivo plano para exportaciones"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   315
      Left            =   2100
      TabIndex        =   8
      Top             =   15
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CommandButton cmd_resumen 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1050
      Picture         =   "frmoracle_packing_list.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Packing list de resumen"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      Picture         =   "frmoracle_packing_list.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Packing list de exportaciones"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   0
      TabIndex        =   5
      Top             =   345
      Width           =   2715
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   375
      Picture         =   "frmoracle_packing_list.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmoracle_packing_list.frx":050A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2340
      Picture         =   "frmoracle_packing_list.frx":060C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   " Embarque "
      Height          =   750
      Left            =   45
      TabIndex        =   0
      Top             =   480
      Width           =   2640
      Begin VB.TextBox txt_embarque 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   2370
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   1305
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   1710
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "frmoracle_packing_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter



Private Sub cmd_facturas_por_embarque_Click()
   If IsNumeric(Me.txt_embarque) Then
      strconsulta = "select char_emb_estatus from xxvia_tb_encabezado_embarques where embarque = ?"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_embarque))
           .Parameters.Append parametro
      End With
      Set rsaux = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      var_posible = 1
      If Not rsaux.EOF Then
         If IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus) = "I" Or IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus) = "F" Or IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus) = "E" Then
            var_posible = 1
         Else
            var_posible = 1
         End If
      Else
         var_posible = 2
      End If
      rsaux.Close
      If var_posible = 1 Then
         cnn.BeginTrans
         rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) AS CONSECUTIVO FROM TB_TEMP_FACTURAs_EXPORTACION", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
         Else
            var_consecutivo = 1
         End If
         rsaux.Close
         rsaux.Open "INSERT INTO TB_TEMP_FACTURAs_EXPORTACION (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
            
         rsaux10.Open "select distinct source_header_number from xxvia_tb_Salidas_cajas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux10.EOF
         
               rs.Open "SELECT distinct inte_emb_embarque as embarque, source_header_number  as pedido, collector_id as agente, name as nombre_agente, customer_id as cliente, customer_name as nombre_cliente FROM XXVIA_TB_salidas_cajas a, xxvia_tb_encabezado_embarques, xxvia_system_items_b b, oe_order_headers_all oh where inte_emb_embarque = embarque and organizacion = b.organization_id and a.inventory_item_id = b.inventory_item_id and order_number = a.source_header_number and nvl(oh.ship_from_org_id,93) = organizacion  and floa_sal_Cantidad_leida >0 and source_header_number = " + CStr(rsaux10!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  rsaux.Open "select * from oe_order_headers_all where order_number = " + CStr(rs!pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If rsaux!ORDER_TYPE_ID = 1002 Then
                     var_agente = rs!Agente
                     var_nombre_cliente = rs!nombre_cliente
                     var_nombre_agente = rs!NOMBRE_AGENTE
                     var_cliente = rs!Cliente
                     rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        var_nombre_cliente = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                     End If
                     rsaux2.Close
                  Else
                     var_agente = rs!Agente
                     var_nombre_cliente = rs!nombre_cliente
                     var_nombre_agente = rs!NOMBRE_AGENTE
                     var_cliente = rs!Cliente
                  End If
                  rsaux.Close
                  rsaux.Open "alter session set nls_language= 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         
                  var_pedido = rs!pedido
                  strconsulta = "SELECT * FROM RA_CUSTOMER_tRX_ALL WHERE CT_REFERENCE = ? "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CDbl(var_pedido))
                       .Parameters.Append parametro
                  End With
                  Set rsaux11 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  While Not rsaux11.EOF
                        VAR_HEADER_ID = rsaux11!customer_Trx_id
                  
         
                        rsaux.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        strconsulta = "select  c.attribute3 , SERIE||D.TRX_NUMBER AS FACTURA, D.TRX_DATE, A1.CT_REFERENCE, a.inventory_item_id, b.segment1, b.description, b.UNIT_WEIGHT as peso, round(sum(unit_selling_price),4) PRECIO, CANTIDAD "
                        strconsulta = strconsulta + "from RA_CUSTOMER_TRX_ALL A1, ra_customer_trx_lines_all a, xxvia_system_items_b b, XXVIA_VW_CANTIDAD_LINEAS_FACT C, xxvia_vw_documento_fiscales D "
                        strconsulta = strconsulta + " Where A.inventory_item_id = b.inventory_item_id "
                        strconsulta = strconsulta + " And b.organization_id = 93 "
                        strconsulta = strconsulta + "    AND A.CUSTOMER_TRX_ID = C.CUSTOMER_TRX_ID "
                        strconsulta = strconsulta + " AND A.CUSTOMER_TRX_ID = D.CUSTOMER_TRX_ID "
                        strconsulta = strconsulta + " AND A.INVENTORY_ITEM_ID = c.INVENTORY_ITEM_ID "
                        strconsulta = strconsulta + " AND A.CUSTOMER_TRX_ID = A1.CUSTOMER_TRX_ID"
                        strconsulta = strconsulta + " AND A1.CT_REFERENCE  = ? "
                        strconsulta = strconsulta + " and a1.customer_trx_id = ? "
                        strconsulta = strconsulta + " and unit_selling_price >= 0 "
                        strconsulta = strconsulta + " and c.attribute3 = a.interface_line_attribute3"
                        strconsulta = strconsulta + " group by c.attribute3 , SERIE||D.TRX_NUMBER, D.TRX_DATE, A1.CT_REFERENCE,a.inventory_item_id, b.segment1, b.description, b.UNIT_WEIGHT,CANTIDAD"
                         
                  
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(var_pedido))
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, VAR_HEADER_ID)
                            .Parameters.Append parametro
                        End With
                        Set rsaux1 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        While Not rsaux1.EOF
                              var_factura = rsaux1!FACTURA
                              
                              var_dia = CStr(Day(CDate(rsaux1!trx_date)))
                              var_mes = CStr(Month(CDate(rsaux1!trx_date)))
                              var_año = CStr(Year(CDate(rsaux1!trx_date)))
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
                                                
                                
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                                         
                              'rsaux.Open "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, NVL(a.description,'') AS localizador, B.UNIT_WEIGHT FROM (select INVENTORY_ITEM_ID, description, cross_reference from mtl_cross_references_b) A, (select inventory_item_id, DESCRIPTION, organization_id, segment1, UNIT_WEIGHT from xxvia_system_items_b) B Where a.inventory_item_id = B.inventory_item_id AND B.organization_id = " + var_unidad_organizacional + " AND b.segment1 = '" + IIf(IsNull(rs!codigo), "", rs!codigo) + "' AND SUBSTR(cross_reference,1,3) = '000' and substr(cross_Reference,9,3) = '000'  AND substr(cross_Reference,12,1) IN( '0','1','2','3','4','5','6','7','8','9')"
                              strconsulta = "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, NVL(a.description,'') AS localizador, B.UNIT_WEIGHT FROM (select INVENTORY_ITEM_ID, description, cross_reference from mtl_cross_references_b) A, (select inventory_item_id, DESCRIPTION, organization_id, segment1, UNIT_WEIGHT from xxvia_system_items_b) B Where a.inventory_item_id = B.inventory_item_id AND B.organization_id = ? AND b.segment1 = ? AND SUBSTR(cross_reference,1,3) = '000' and substr(cross_Reference,9,3) = '000'  AND substr(cross_Reference,12,1) IN( '0','1','2','3','4','5','6','7','8','9')"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux1!SEGMENT1), "", rsaux1!SEGMENT1))
                                   .Parameters.Append parametro
                              End With
                              Set rsaux = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
               
                              If Not rsaux.EOF Then
                                 var_codigo_barras = IIf(IsNull(rsaux!cross_reference), "", rsaux!cross_reference)
                              Else
                                 var_codigo_barras = ""
                              End If
                              rsaux.Close
                        
                         
                              strconsulta = "select * from xxvia_tb_complementos_pk_list where codigo = ?"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux1!SEGMENT1), "", rsaux1!SEGMENT1))
                                  .Parameters.Append parametro
                              End With
                              Set rsaux9 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                              If Not rsaux9.EOF Then
                                 var_arancel = IIf(IsNull(rsaux9!fraccion_arancelaria), 0, rsaux9!fraccion_arancelaria)
                                 VAR_COMPOSICION = IIf(IsNull(rsaux9!composicion), "", rsaux9!composicion)
                                 VAR_CONTENIDO = IIf(IsNull(rsaux9!contenido), "", rsaux9!contenido)
                                 VAR_ORIGEN = IIf(IsNull(rsaux9!originario), "", rsaux9!originario)
                                 VAR_APLICA_USA = IIf(IsNull(rsaux9!aplica_usa), "", rsaux9!aplica_usa)
                                 VAR_APLICA_CA = IIf(IsNull(rsaux9!aplica_ca), "", rsaux9!aplica_ca)
                                 VAR_HECHO_EN = IIf(IsNull(rsaux9!hecho_en), "", rsaux9!hecho_en)
                                 VAR_COMPLEMENTO_1 = IIf(IsNull(rsaux9!complemento_1), "", rsaux9!complemento_1)
                                 VAR_PRECIO_1 = var_precio * (IIf(IsNull(rsaux9!precio_1), 0, rsaux9!precio_1) / 100)
                                 VAR_COMPLEMENTO_2 = IIf(IsNull(rsaux9!complemento_2), "", rsaux9!complemento_2)
                                 VAR_PRECIO_2 = var_precio * (IIf(IsNull(rsaux9!precio_2), 0, rsaux9!precio_2) / 100)
                                 VAR_COMPLEMENTO_3 = IIf(IsNull(rsaux9!complemento_3), "", rsaux9!complemento_3)
                                 VAR_PRECIO_3 = var_precio * (IIf(IsNull(rsaux9!precio_3), 0, rsaux9!precio_3) / 100)
                                 VAR_COMPLEMENTO_4 = IIf(IsNull(rsaux9!complemento_4), "", rsaux9!complemento_4)
                                 VAR_PRECIO_4 = var_precio * (IIf(IsNull(rsaux9!precio_4), 0, rsaux9!precio_4) / 100)
                                 VAR_COMPLEMENTO_5 = IIf(IsNull(rsaux9!complemento_5), "", rsaux9!complemento_5)
                                 VAR_PRECIO_5 = var_precio * (IIf(IsNull(rsaux9!precio_5), 0, rsaux9!precio_5) / 100)
                                 VAR_COMPLEMENTO_6 = IIf(IsNull(rsaux9!complemento_6), "", rsaux9!complemento_6)
                                 VAR_PRECIO_6 = var_precio * (IIf(IsNull(rsaux9!precio_6), 0, rsaux9!precio_6) / 100)
                                 VAR_COMPLEMENTO_7 = IIf(IsNull(rsaux9!complemento_7), "", rsaux9!complemento_7)
                                 VAR_PRECIO_7 = var_precio * (IIf(IsNull(rsaux9!precio_7), 0, rsaux9!precio_7) / 100)
                                 VAR_COMPLEMENTO_8 = IIf(IsNull(rsaux9!complemento_8), "", rsaux9!complemento_8)
                                 VAR_PRECIO_8 = var_precio * (IIf(IsNull(rsaux9!precio_8), 0, rsaux9!precio_8) / 100)
                                 VAR_COMPLEMENTO_9 = IIf(IsNull(rsaux9!complemento_9), "", rsaux9!complemento_9)
                                 VAR_PRECIO_9 = var_precio * (IIf(IsNull(rsaux9!precio_9), 0, rsaux9!precio_9) / 100)
                                 VAR_COMPLEMENTO_10 = IIf(IsNull(rsaux9!complemento_10), "", rsaux9!complemento_10)
                                 VAR_PRECIO_10 = var_precio * (IIf(IsNull(rsaux9!precio_10), 0, rsaux9!precio_10) / 100)
                                 VAR_COMPLEMENTO_11 = IIf(IsNull(rsaux9!complemento_11), "", rsaux9!complemento_11)
                                 VAR_PRECIO_11 = var_precio * (IIf(IsNull(rsaux9!precio_11), 0, rsaux9!precio_11) / 100)
                                 VAR_COMPLEMENTO_12 = IIf(IsNull(rsaux9!complemento_12), "", rsaux9!complemento_12)
                                 VAR_PRECIO_12 = var_precio * (IIf(IsNull(rsaux9!precio_12), 0, rsaux9!precio_12) / 100)
                                 VAR_COMPLEMENTO_13 = IIf(IsNull(rsaux9!complemento_13), "", rsaux9!complemento_13)
                                 VAR_PRECIO_13 = var_precio * (IIf(IsNull(rsaux9!precio_13), 0, rsaux9!precio_13) / 100)
                                 VAR_COMPLEMENTO_14 = IIf(IsNull(rsaux9!complemento_14), "", rsaux9!complemento_14)
                                 VAR_PRECIO_14 = var_precio * (IIf(IsNull(rsaux9!precio_14), 0, rsaux9!precio_14) / 100)
                                 VAR_COMPLEMENTO_15 = IIf(IsNull(rsaux9!complemento_15), "", rsaux9!complemento_15)
                                 VAR_PRECIO_15 = var_precio * (IIf(IsNull(rsaux9!precio_15), 0, rsaux9!precio_15) / 100)
                                 VAR_COMPLEMENTO_16 = IIf(IsNull(rsaux9!complemento_16), "", rsaux9!complemento_16)
                                 VAR_PRECIO_16 = var_precio * (IIf(IsNull(rsaux9!precio_16), 0, rsaux9!precio_16) / 100)
                                 VAR_ARANCEL_AMERICANO = IIf(IsNull(rsaux9!fraccion_americana), 0, rsaux9!fraccion_americana)
                           
                              Else
                                 var_arancel = 0
                                 var_comosicion = ""
                                 var_contenito = ""
                                 VAR_ORIGEN = ""
                                 VAR_APLICA_USA = ""
                                 VAR_APLICA_CA = ""
                                 VAR_HECHO_ENT = ""
                                 VAR_COMPLEMENTO_1 = ""
                                 VAR_PRECIO_1 = 0
                                 VAR_COMPLEMENTO_2 = ""
                                 VAR_PRECIO_2 = 0
                                 VAR_COMPLEMENTO_3 = ""
                                 VAR_PRECIO_3 = 0
                                 VAR_COMPLEMENTO_4 = ""
                                 VAR_PRECIO_4 = 0
                                 VAR_COMPLEMENTO_5 = ""
                                 VAR_PRECIO_5 = 0
                                 VAR_COMPLEMENTO_6 = ""
                                 VAR_PRECIO_6 = 0
                                 VAR_COMPLEMENTO_7 = ""
                                 VAR_PRECIO_7 = 0
                                 VAR_COMPLEMENTO_8 = ""
                                 VAR_PRECIO_8 = 0
                                 VAR_COMPLEMENTO_9 = ""
                                 VAR_PRECIO_9 = 0
                                 VAR_COMPLEMENTO_10 = ""
                                 VAR_PRECIO_10 = 0
                                 VAR_COMPLEMENTO_11 = ""
                                 VAR_PRECIO_11 = 0
                                 VAR_COMPLEMENTO_12 = ""
                                 VAR_PRECIO_12 = 0
                                 VAR_COMPLEMENTO_13 = ""
                                 VAR_PRECIO_13 = 0
                                 VAR_COMPLEMENTO_14 = ""
                                 VAR_PRECIO_14 = 0
                                 VAR_COMPLEMENTO_15 = ""
                                 VAR_PRECIO_15 = 0
                                 VAR_COMPLEMENTO_16 = ""
                                 VAR_PRECIO_16 = 0
                                 VAR_ARANCEL_AMERICANO = 0
                        
                              End If
                              rsaux9.Close
               
                   
                              var_cadena = "INSERT INTO TB_TEMP_FACTURAS_EXPORTACION (INTE_TEM_CONSECUTIVO, EMBARQUE, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, PRECIO, PESO, CONTENIDO, INVENTORY_ITEM_ID, ARANCEL, CODIGO_BARRAS, COMPOSICION, ORIGEN, APLICA_USA, APLICA_CA, HECHO_EN, PEDIDO, FACTURA, FECHA_FACTURA, ATRIBUTO, FRACCION_AMERICANA)"
                              var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + "," + CStr(rs!Embarque) + ",'" + CStr(IIf(IsNull(var_agente), "", var_agente)) + "','" + var_nombre_agente + "', '" + CStr(var_cliente) + "','" + IIf(IsNull(var_nombre_cliente), "", var_nombre_cliente) + "','" + rsaux1!SEGMENT1 + "','" + rsaux1!Description + "'," + CStr(rsaux1!cantidad) + "," + CStr(IIf(IsNull(rsaux1!Precio), 0, rsaux1!Precio)) + "," + CStr(IIf(IsNull(rsaux1!PESO), 0, rsaux1!PESO)) + ",'" + VAR_CONTENIDO + "'," + CStr(rsaux1!inventory_item_id) + ", '" + CStr(var_arancel) + "','" + var_codigo_barras + "','" + VAR_COMPOSICION + "','" + VAR_ORIGEN + "','" + VAR_APLICA_USA + "','" + VAR_APLICA_CA + "', '" + VAR_HECHO_EN + "'," + CStr(var_pedido) + ", '" + var_factura + "'," + var_fecha + ",'" + CStr(rsaux1!attribute3) + "'," + CStr(VAR_ARANCEL_AMERICANO) + ")"
                              rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              rsaux1.MoveNext
                        Wend
                        rsaux1.Close
                        
                        strconsulta = "select c.attribute3 , SERIE||D.TRX_NUMBER AS FACTURA, D.TRX_DATE, A1.CT_REFERENCE, a.inventory_item_id, b.segment1, b.description, b.UNIT_WEIGHT as peso, round(sum(unit_selling_price),4) PRECIO, CANTIDAD "
                        strconsulta = strconsulta + "from RA_CUSTOMER_TRX_ALL A1, ra_customer_trx_lines_all a, xxvia_system_items_b b, XXVIA_VW_CANTIDAD_LINEAS_FACT C, xxvia_vw_documento_fiscales D "
                        strconsulta = strconsulta + " Where A.inventory_item_id = b.inventory_item_id "
                        strconsulta = strconsulta + " And b.organization_id = 93 "
                        strconsulta = strconsulta + "    AND A.CUSTOMER_TRX_ID = C.CUSTOMER_TRX_ID "
                        strconsulta = strconsulta + " AND A.CUSTOMER_TRX_ID = D.CUSTOMER_TRX_ID "
                        strconsulta = strconsulta + " AND A.INVENTORY_ITEM_ID = c.INVENTORY_ITEM_ID "
                        strconsulta = strconsulta + " AND A.CUSTOMER_TRX_ID = A1.CUSTOMER_TRX_ID"
                        strconsulta = strconsulta + " AND A1.CT_REFERENCE  = ? "
                        strconsulta = strconsulta + " and a1.customer_trx_id = ? "
                        strconsulta = strconsulta + " and unit_selling_price < 0 "
                        strconsulta = strconsulta + " and c.attribute3 = a.interface_line_attribute3"
                        strconsulta = strconsulta + " group by c.attribute3 , SERIE||D.TRX_NUMBER, D.TRX_DATE, A1.CT_REFERENCE,a.inventory_item_id, b.segment1, b.description, b.UNIT_WEIGHT,CANTIDAD"
                         
                  
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(var_pedido))
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, VAR_HEADER_ID)
                            .Parameters.Append parametro
                        End With
                        Set rsaux1 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        While Not rsaux1.EOF
                              rsaux2.Open "UPDATE TB_TEMP_FACTURAS_EXPORTACION SET PRECIO = PRECIO + " + CStr(rsaux1!Precio) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND INVENTORY_ITEM_ID = " + CStr(rsaux1!inventory_item_id) + " AND ATRIBUTO = " + CStr(rsaux1!attribute3), cnn, adOpenDynamic, adLockOptimistic
                              rsaux1.MoveNext
                        Wend
                        rsaux1.Close
                        
var_x = 0
                        If var_x = 1 Then
                           strconsulta = "SELECT PEDIDO, CAJA, CODIGO, decode(SUBSTR(codigo_barras,9,1), 1,'CHINO','MEXICO') ORIGEN, SUM(CANTIDAD) CANTIDAD FROM XXVIA_TB_BITACORA_LECTURA where pedido = ? and codigo = ? GROUP BY PEDIDO, CAJA, CODIGO, decode(SUBSTR(codigo_barras,9,1), 1,'CHINO','MEXICO')"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(var_pedido), "", var_pedido))
                               .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rs!codigo), "", rs!codigo))
                               .Parameters.Append parametro
                           End With
                           Set rsaux9 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           VAR_CANTIDAD_MEXICO = 0
                           VAR_CANTIDAD_CHINA = 0
                           If Not rsaux9.EOF Then
                              While Not rsaux9.EOF
                                    If rsaux9!ORIGEN = "MEXICO" Then
                                       VAR_CANTIDAD_MEXICO = rsaux9!cantidad
                                    Else
                                       VAR_CANTIDAD_CHINA = rsaux9!cantidad
                                    End If
                                    rsaux9.MoveNext
                              Wend
                           End If
                           rsaux9.Close
                       End If
                         
                       rsaux11.MoveNext
                  Wend
                  rsaux11.Close
               Else
                  MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
               End If
               rs.Close
               
               rsaux10.MoveNext
         Wend
         rsaux10.Close
         rsaux1.Open "DELETE FROM tb_temp_facturas_exportacion WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND EMBARQUE IS NULL", cnn, adOpenDynamic, adLockOptimistic
         rsaux1.Open "select distinct replace(factura,'FAEVII','') AS FACTURA FROM tb_temp_facturas_exportacion WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         
         'rsaux1.Open "SELECT * FROM tb_temp_facturas_exportacion WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         'While Not rsaux1.EOF
         '      strconsulta = "SELECT * FROM WSH_DELIVERABLES_V WHERE SOURCE_HEADER_NUMBER = ? AND INVENTORY_ITEM_ID = ?"
         '      With comandoORA
         '           .ActiveConnection = cnnoracle_4
         '           .CommandType = adCmdText
         '           .CommandText = strconsulta
         '           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux1!pedido), "", rsaux1!pedido))
         '           .Parameters.Append parametro
         '           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux1!INVENTORY_ITEM_ID), "", rsaux1!INVENTORY_ITEM_ID))
         '           .Parameters.Append parametro
         '       End With
         '       Set rsaux9 = comandoORA.execute
         '       Set comandoORA = Nothing
         '       Set parametro = Nothing
         '       If Not rsaux9.EOF Then
         '          rsaux10.Open "UPDATE tb_temp_facturas_exportacion SET CONSECUTIVO_FACTURA = " + CStr(IIf(IsNull(rsaux9!ATTRIBUTE15), "0", rsaux9!ATTRIBUTE15)) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND INVENTORY_ITEM_ID = " + CStr(rsaux1!INVENTORY_ITEM_ID), cnn, adOpenDynamic, adLockOptimistic
         '       End If
         '       rsaux9.Close
         '       rsaux1.MoveNext
         'Wend
         'rsaux1.Close
         
         While Not rsaux1.EOF
               strconsulta = "select cadena as cadena from xxvia_tb_control_doc_fiscales where serie = 'FAEVII' and numero = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(rsaux1!FACTURA))
                    .Parameters.Append parametro
               End With
               Set rsaux = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               var_cadena = rsaux!Cadena
               var_cadena_rfc = Mid(var_cadena, 34, 12)
               VAR_CADENA_STR = ""
               var_consecutivo_FACTURA = 0
               For var_i = 1 To Len(var_cadena)
                   If Asc(Mid(var_cadena, var_i, 1)) = 63 Then
                      If Mid(VAR_CADENA_STR, 1, 6) = "LINEA:" Then
                         var_consecutivo_FACTURA = var_consecutivo_FACTURA + 1
                         var_listo = 0
                         VAR_LISTO_2 = 0
                         var_codigo = ""
                         var_descripcion = ""
                         var_c = 0
                         VAR_CANTIDAD_LETRA = ""
                         
                         x = 1
                         If x = 0 Then
                         For var_j = 7 To Len(VAR_CADENA_STR)
                              
                             VAR_LETRA = Mid(VAR_CADENA_STR, var_j, 1)
                             If var_c = 0 Then
                                If VAR_LETRA <> "|" Then
                                   VAR_CANTIDAD_LETRA = VAR_CANTIDAD_LETRA + Mid(VAR_CADENA_STR, var_j, 1)
                                Else
                                   var_c = 1
                                End If
                             End If
                             If VAR_LETRA = "|" Then
                                If var_listo = 2 Then
                                   var_listo = 3
                                End If
                                If var_listo = 0 Then
                                   var_listo = 2
                                End If
                             End If
                             If var_listo = 2 Then
                                var_codigo = var_codigo + Mid(VAR_CADENA_STR, var_j, 1)
                             End If
                             
                             If var_listo = 3 Then
                                If VAR_LETRA = "|" And var_listo = 3 Then
                                   var_listo = 4
                                Else
                                   var_descripcion = var_descripcion + Mid(VAR_CADENA_STR, var_j, 1)
                                End If
                             End If
                             If var_listo > 3 Then
                                If var_listo = 4 And VAR_LETRA <> "|" Then
                                   var_descripcion = var_descripcion + Mid(VAR_CADENA_STR, var_j, 1)
                                   var_listo = 5
                                Else
                                   If var_listo = 5 Then
                                      If VAR_LETRA <> "|" Then
                                         var_descripcion = var_descripcion + Mid(VAR_CADENA_STR, var_j, 1)
                                      Else
                                         var_listo = 6
                                      End If
                                   End If
                                End If
                                
                             End If
                         Next var_j
                         Else
                                var_precio_s = ""
                                VAR_CODIGO_S = ""
                                VAR_cANTIDAD_S = ""
                                var_descripcion_s = ""
                                var_importe_s = ""
                                var_unidad_s = ""
                                var_Cantidad_completa = 0
                                var_codigo_completo = 1
                                var_precio_completo = 1
                                var_importe_completo = 1
                                var_descripcion_completa = 1
                                var_unidad_completa = 1
                                var_codigo_completo = 0
                         
                                For VAR_Z = 15 To Len(VAR_CADENA_STR)
                                    If var_codigo_completo = 0 Then
                                       If Mid(VAR_CADENA_STR, VAR_Z, 1) = "|" Then
                                          var_descripcion_completa = 0
                                          var_codigo_completo = 1
                                          GoTo completo:
                                       Else
                                          VAR_cANTIDAD_S = VAR_cANTIDAD_S + Mid(VAR_CADENA_STR, VAR_Z, 1)
                                       End If
                                    End If
                                    If var_Cantidad_completa = 0 Then
                                       If Mid(VAR_CADENA_STR, VAR_Z, 1) = "|" Then
                                          var_Cantidad_completa = 1
                                          var_codigo_completo = 0
                                          GoTo completo:
                                       Else
                                          VAR_CODIGO_S = VAR_CODIGO_S + Mid(VAR_CADENA_STR, VAR_Z, 1)
                                       End If
                                    End If
                                    If var_unidad_completa = 0 Then
                                       If Mid(VAR_CADENA_STR, VAR_Z, 1) = "|" Then
                                          var_unidad_completa = 1
                                          var_precio_completo = 0
                                          GoTo completo:
                                       Else
                                          var_unidad_s = var_unidad_s + Mid(VAR_CADENA_STR, VAR_Z, 1)
                                       End If
                                    End If
                                    
                                    
                                    
                                    If var_descripcion_completa = 0 Then
                                       If Mid(VAR_CADENA_STR, VAR_Z, 1) = "|" Then
                                          var_unidad_completa = 0
                                          var_descripcion_completa = 1
                                          GoTo completo:
                                       Else
                                          var_descripcion_s = var_descripcion_s + Mid(VAR_CADENA_STR, VAR_Z, 1)
                                       End If
                                    End If
                                    
                                    
                                    If var_precio_completo = 0 Then
                                       If Mid(VAR_CADENA_STR, VAR_Z, 1) = "|" Then
                                          var_importe_completo = 0
                                          var_precio_completo = 1
                                          GoTo completo:
                                       Else
                                          var_precio_s = var_precio_s + Mid(VAR_CADENA_STR, VAR_Z, 1)
                                       End If
                                    End If
                                    
                                    If var_importe_completo = 0 Then
                                       If Mid(VAR_CADENA_STR, VAR_Z, 1) = "|" Then
                                          var_importe_completo = 1
                                          var_precio_completo = 1
                                          GoTo completo:
                                       Else
                                          var_importe_s = var_importe_s + Mid(VAR_CADENA_STR, VAR_Z, 1)
                                       End If
                                    End If
                                    
                                    
completo:
                                Next VAR_Z
                          
                         
                         
                         End If
                         
                         If Mid(var_codigo, 2, Len(var_codigo)) = "00085680" Then
                            var_codigo = var_codigo
                         End If
                         VAR_CANTIDAD_LETRA = CDbl(VAR_cANTIDAD_S)
                         'rsaux2.Open "UPDATE tb_temp_facturas_exportacion SET CONSECUTIVO_FACTURA = " + CStr(var_consecutivo_FACTURA) + ", descripcion = '" + Replace(VAR_DESCRIPCION, "*", "") + "' WHERE CODIGO = '" + CStr(Mid(var_codigo, 2, Len(var_codigo))) + "' AND FACTURA = 'FAEVII" + CStr(rsaux1!FACTURA) + "' AND INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND CANTIDAD = " + VAR_CANTIDAD_LETRA, cnn, adOpenDynamic, adLockOptimistic
                         rsaux2.Open "UPDATE tb_temp_facturas_exportacion SET CONSECUTIVO_FACTURA = " + CStr(var_consecutivo_FACTURA) + " WHERE CODIGO = '" + VAR_CODIGO_S + "' AND FACTURA = 'FAEVII" + CStr(rsaux1!FACTURA) + "' AND INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND CANTIDAD = " + CStr(VAR_CANTIDAD_LETRA), cnn, adOpenDynamic, adLockOptimistic
                      End If
                      VAR_CADENA_STR = ""
                   Else
                      VAR_CADENA_STR = VAR_CADENA_STR + Mid(var_cadena, var_i, 1)
                   End If
              Next var_i
              rsaux1.MoveNext
         Wend
         rsaux1.Close
         
         
         
         rsaux1.Open "select distinct factura FROM tb_temp_facturas_exportacion WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux1.EOF
               x = 0
               If x = 1 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_facturas_exportaciones.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_FACTURAS_EXPORTACIONES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_FACTURAS_EXPORTACIONES.factura} = '" + rsaux1!FACTURA + "'"
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\hoja_trabajo_factura_" + rsaux1!FACTURA + " " & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               Else
                   
                     Set oexcel = CreateObject("Excel.Application")
                     Set owbook = oexcel.Workbooks.Add
                     Set osheet = owbook.Worksheets(1)
                     osheet.Name = "HOJA DE TRABAJO"
                     Screen.MousePointer = vbHourglass
                     iFila = 1
                     ifila2 = 1
                     icol2 = 1
                     iCol = 1
                     var_cadena = "select EMBARQUE, PEDIDO, FACTURA, FECHA_FACTURA, NOMBRE_AGENTE, NOMBRE_CLIENTE, CODIGO, CODIGO_BARRAS, DESCRIPCION, 'C62_1' UNIDAD_MEDIDA, CANTIDAD, CANTIDAD * PRECIO IMPORTE, 'USD' MONEDA, PESO, CANTIDAD * PESO TOTAL_PESO, ARANCEL, CONTENIDO, HECHO_EN, APLICA_USA, APLICA_CA, PRECIO, CONSECUTIVO_FACTURA, FRACCION_AMERICANA from VW_ORACLE_FACTURAS_EXPORTACIONES WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND CODIGO IS NOT NULL order by CONSECUTIVO_FACTURA"
                     rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     For i = 0 To rsaux10.Fields.Count - 1
                         osheet.Cells(iFila, i + 1) = rsaux10.Fields(i).Name
                     Next
                     iFila = iFila + 1
                     With osheet
                         ' carga los registros del recordset
                         .Cells(iFila, iCol).CopyFromRecordset rsaux10
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.NumberFormat = "#,##0.00"
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.Font.Color = vbRed
                         .Columns.AutoFit ' ajusta el ancho de las columnas
                     End With
                     archivo = "c:\reportessid\hoja_trabajo_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     owbook.SaveAs archivo
                     oexcel.Visible = True
                     Set oexcel = Nothing
                     Screen.MousePointer = vbDefault
                     rsaux10.Close
                   
                   
               End If
               MsgBox "Se a terminado de guardar el archivo " + archivo
               rsaux1.MoveNext
         Wend
         rsaux1.Close
         rsaux2.Open "delete from tb_temp_facturas_exportacion where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      Else
         If var_posible = 0 Then
            MsgBox "El embarque no a sido cerrado.", vbOKOnly, "ATENCION"
         End If
         If var_posible = 2 Then
            MsgBox "El embarque no existe.", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub cmd_imprimir_Click()
   If IsNumeric(Me.txt_embarque) Then
      strconsulta = "select char_emb_estatus from xxvia_tb_encabezado_embarques where embarque = ?"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_embarque))
           .Parameters.Append parametro
      End With
      Set rsaux = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      var_posible = 1
      If Not rsaux.EOF Then
         If IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus) = "I" Or IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus) = "F" Or IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus) = "E" Then
            var_posible = 1
         Else
            var_posible = 1
         End If
      Else
         var_posible = 2
      End If
      rsaux.Close
      
      If var_posible = 1 Then
         If rsaux8.State = 1 Then
            rsaux8.Close
         End If
         rsaux8.Open "SELECT DISTINCT SOURCE_HEADER_NUMBER AS PEDIDO FROM XXVIA_TB_salidas_cajas WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " and floa_sal_cantidad_leida > 0", cnnoracle_4, adOpenDynamic, adLockOptimistic
         'rs.Open "SELECT inte_emb_embarque as embarque, inte_paq_caja as caja, source_header_number  as pedido, segment1 as codigo, collector_id as agente,name as nombre_agente, customer_id as cliente, customer_name as nombre_cliente, inventory_item_id, caja_pedido, sello, item_description as descripcion, floa_sal_cantidad_leida as cantidad   FROM XXVIA_TB_salidas_cajas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux8.EOF Then
            cnn.BeginTrans
            rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) AS CONSECUTIVO FROM TB_TEMP_ORACLE_DETALLE_CAJAS", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rsaux.Close
            rsaux.Open "INSERT INTO TB_TEMP_ORACLE_DETALLE_CAJAS (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            While Not rsaux8.EOF
                  If rs.State = 1 Then
                     rs.Close
                  End If
                  rs.Open "SELECT floa_sal_Cantidad_leida as cantidad, organizacion,  inte_emb_embarque as embarque, inte_paq_caja as caja, source_header_number  as pedido, a.segment1 as codigo, collector_id as agente,name as nombre_agente, customer_id as cliente, customer_name as nombre_cliente, a.inventory_item_id, caja_pedido, sello, UNIT_WEIGHT as peso, item_description as descripcion, tipo_caja    FROM XXVIA_TB_salidas_cajas a, xxvia_tb_encabezado_embarques, xxvia_system_items_b b, oe_order_headers_all oh where inte_emb_embarque = embarque and organizacion = b.organization_id and a.inventory_item_id = b.inventory_item_id and order_number = a.source_header_number and  oh.ship_from_org_id = organizacion and  inte_emb_embarque = " + Me.txt_embarque + " and floa_sal_Cantidad_leida >0 AND SOURCE_HEADER_NUMBER = " + CStr(rsaux8!pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  'rs.Open "SELECT floa_sal_Cantidad_leida as cantidad, organizacion,  inte_emb_embarque as embarque, inte_paq_caja as caja, source_header_number  as pedido, a.segment1 as codigo, collector_id as agente,name as nombre_agente, customer_id as cliente, customer_name as nombre_cliente, a.inventory_item_id, caja_pedido, sello, UNIT_WEIGHT as peso, item_description as descripcion, tipo_caja    FROM XXVIA_TB_salidas_cajas a, xxvia_tb_encabezado_embarques, xxvia_system_items_b b, oe_order_headers_all oh where inte_emb_embarque = embarque and organizacion = b.organization_id and a.inventory_item_id = b.inventory_item_id and order_number = a.source_header_number and  oh.ship_from_org_id = organizacion and  inte_emb_embarque = " + Me.txt_embarque + " and floa_sal_Cantidad_leida >0 AND SOURCE_HEADER_NUMBER = " + CStr(rsaux8!pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  
                  rsaux.Open "select * from oe_order_headers_all where order_number = " + CStr(rs!pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If rsaux!ORDER_TYPE_ID = 1002 Then
                     var_agente = rs!Agente
                     var_nombre_cliente = rs!nombre_cliente
                     var_nombre_agente = rs!NOMBRE_AGENTE
                     var_cliente = rs!Cliente
                     rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        var_nombre_cliente = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                     End If
                     rsaux2.Close
                  Else
                     var_agente = rs!Agente
                     var_nombre_cliente = rs!nombre_cliente
                     var_nombre_agente = rs!NOMBRE_AGENTE
                     var_cliente = rs!Cliente
                  End If
                  rsaux.Close
                  rsaux.Open "alter session set nls_language= 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rs.EOF
                        'var_cadena = "select ic.category_concat_segs,fft.description from mtl_item_categories_v ic, fnd_flex_value_sets   ffvs, fnd_flex_values_vl ffv,  fnd_flex_values_tl fft,mtl_parameters mtp where UPPER (ic.category_set_name) LIKE '%VIANNEY%EXPORTACION%' AND ic.inventory_item_id = " + CStr(rs!INVENTORY_ITEM_ID) + " AND ic.organization_id = mtp.organization_id AND mtp.organization_code = 'MTO' AND ic.category_concat_segs =  ffv.flex_value_meaning AND ffvs.flex_value_set_name = 'VIANNEY_INV_EXPORTACION' AND ffvs.flex_value_set_id  =  ffv.flex_value_set_id AND ffv.flex_value_id  =  fft.flex_value_id AND fft.language  =  USERENV('LANG') "
                        'If rsaux.State = 1 Then
                        '   rsaux.Close
                        'End If
                        'rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'If Not rsaux.EOF Then
                        '   var_arancel = CStr(IIf(IsNull(rsaux!Description), "", rsaux!Description))
                        'Else
                        '   var_arancel = ""
                        'End If
                        'rsaux.Close
               
                        var_peso = IIf(IsNull(rs!PESO), 0, rs!PESO)
               
                        'var_pedido = rs!pedido
                        'rsaux.Open "select unit_selling_price from oe_order_headers_all oh, oe_order_lines_all ol where order_number = " + CStr(var_pedido) + " and oh.header_id = ol.header_id and oh.ship_from_org_id = " + var_unidad_organizacional + " and ol.inventory_item_id = " + CStr(IIf(IsNull(rs!INVENTORY_ITEM_ID), 0, rs!INVENTORY_ITEM_ID)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'If Not rsaux.EOF Then
                        '   var_precio = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                        'Else
                        '   var_precio = 0
                        'End If
                        'rsaux.Close
                        var_precio = 0
                        
                        'MsgBox Len(rs!sello)
                        var_cadena = "INSERT INTO TB_TEMP_ORACLE_DETALLE_CAJAS (INTE_TEM_CONSECUTIVO, EMBARQUE, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, PRECIO, PESO, CAJA, CONTENIDO, INVENTORY_ITEM_ID, ARANCEL, CODIGO_BARRAS, SELLO, CAJA_PEDIDO, PEDIDO, bulto)"
                        var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + "," + CStr(rs!Embarque) + ",'" + CStr(IIf(IsNull(var_agente), "", var_agente)) + "','" + IIf(IsNull(var_nombre_agente), "", var_nombre_agente) + "', '" + CStr(var_cliente) + "','" + IIf(IsNull(var_nombre_cliente), "", var_nombre_cliente) + "','" + rs!codigo + "','" + Mid(IIf(IsNull(rs!Descripcion), "", rs!Descripcion), 1, 100) + "'," + CStr(rs!cantidad) + "," + CStr(var_precio) + "," + CStr(var_peso) + "," + CStr(rs!Caja) + ",''," + CStr(rs!inventory_item_id) + ", '" + IIf(IsNull(var_arancel), "", var_arancel) + "','" + IIf(IsNull(var_codigo_barras), "", var_codigo_barras) + "','" + Trim(IIf(IsNull(rs!sello), "", rs!sello)) + "'," + CStr(IIf(IsNull(rs!caja_pedido), rs!Caja, rs!caja_pedido)) + "," + CStr(rs!pedido) + ",'" + IIf(IsNull(rs!tipo_caja), "", rs!tipo_caja) + "')"
                        'MsgBox var_cadena
                        rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        
                        
                        rs.MoveNext
                  Wend
                  rsaux1.Open "DELETE FROM TB_TEMP_ORACLE_DETALLE_CAJAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND EMBARQUE IS NULL", cnn, adOpenDynamic, adLockOptimistic
             
                  rsaux1.Open "SELECT DISTINCT CAJA FROM TB_TEMP_ORACLE_DETALLE_CAJAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux8!pedido), cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveFirst
                  VAR_CANTIDAD_CAJAS = 0
                  While Not rsaux1.EOF
                        VAR_CANTIDAD_CAJAS = VAR_CANTIDAD_CAJAS + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "UPDATE TB_TEMP_ORACLE_DETALLE_CAJAS SET NUMERO_CAJAS = " + CStr(VAR_CANTIDAD_CAJAS) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux8!pedido), cnn, adOpenDynamic, adLockOptimistic
                  rsaux8.MoveNext
                  rs.Close
            Wend
            rsaux8.Close
            rsaux8.Open "SELECT DISTINCT PEDIDO FROM TB_TEMP_ORACLE_DETALLE_CAJAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux8.EOF
                  VAR_CADENA_BULTOS = ""
                  strconsulta = "SELECT tipo_caja, COUNT(*) CANTIDAD FROM XXVIA_VW_CAJAS_POR_PEDIDO WHERE SOURCE_HEADER_NUMBER = ? GROUP BY TIPO_CAJA"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 20, rsaux8!pedido)
                       .Parameters.Append parametro
                  End With
                  Set rsaux9 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  While Not rsaux9.EOF
                        If VAR_CADENA_BULTOS = "" Then
                           VAR_CADENA_BULTOS = rsaux9!tipo_caja + ": " + CStr(rsaux9!cantidad)
                        Else
                           VAR_CADENA_BULTOS = VAR_CADENA_BULTOS + ",    " + rsaux9!tipo_caja + ": " + CStr(rsaux9!cantidad)
                        End If
                        rsaux9.MoveNext
                  Wend
                  rsaux9.Close
                  rsaux9.Open "UPDATE TB_TEMP_ORACLE_DETALLE_CAJAS SET VCHA_PAQ_TIPO_BULTOS = '" + IIf(IsNull(VAR_CADENA_BULTOS), "", VAR_CADENA_BULTOS) + "' WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux8!pedido), cnn, adOpenDynamic, adLockOptimistic
                  rsaux8.MoveNext
            Wend
            rsaux8.Close
         
         
         
            
            
            
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_packing_list.rpt")
            reporte.RecordSelectionFormula = "{VW_ORACLE_DETALLE_CAJAS.EMBARQUE} = " + txt_embarque + " and {VW_ORACLE_DETALLE_CAJAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Packing List"
            frmvistasprevias.Show 1
            Set reporte = Nothing
       
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_packing_list.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_DETALLE_CAJAS.EMBARQUE} = " + txt_embarque + " and {VW_ORACLE_DETALLE_CAJAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\packing_list" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               
               MsgBox "Se a terminado de guardar el archivo " + archivo
               var_si = MsgBox("Desea enviar el packing list por correo?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  VAR_CORREO_ELECTRONICO = "vluna@vianney.com.mx"
                  If Trim(VAR_CORREO_ELECTRONICO) <> "" Then
                     If MAPISession1.SessionID = 0 Then
                        MAPISession1.SignOn
                     End If
                     MAPIMessages1.SessionID = MAPISession1.SessionID
                     MAPIMessages1.Compose
                     MAPIMessages1.RecipDisplayName = VAR_CORREO_ELECTRONICO
                     MAPIMessages1.RecipAddress = VAR_CORREO_ELECTRONICO
                     MAPIMessages1.AddressResolveUI = True
                     MAPIMessages1.ResolveName
                     MAPIMessages1.MsgSubject = "Packing list"
                     MAPIMessages1.MsgNoteText = "Se anexa archivo de packing list"
                     MAPIMessages1.AttachmentPathName = archivo
                     MAPIMessages1.send True
                     If MAPISession1.SessionID > 0 Then
                        MAPISession1.SignOff
                     End If
                  End If
               End If
            End If
            rsaux.Open "delete from TB_TEMP_ORACLE_DETALLE_CAJAS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
      Else
         If var_posible = 0 Then
            MsgBox "El embarque no a sido cerrado.", vbOKOnly, "ATENCION"
         End If
         If var_posible = 2 Then
            MsgBox "El embarque no existe.", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_embarque = ""
   Me.txt_embarque.SetFocus
End Sub

Private Sub cmd_resumen_Click()
   If IsNumeric(Me.txt_embarque) Then
      strconsulta = "select char_emb_estatus from xxvia_tb_encabezado_embarques where embarque = ?"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_embarque))
           .Parameters.Append parametro
      End With
      Set rsaux = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      var_posible = 1
      If Not rsaux.EOF Then
         If IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus) = "I" Or IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus) = "F" Or IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus) = "E" Then
            var_posible = 1
         Else
            var_posible = 1
         End If
      Else
         var_posible = 2
      End If
      rsaux.Close
      If var_posible = 1 Then
         rsaux8.Open "SELECT DISTINCT SOURCE_HEADER_NUMBER AS PEDIDO FROM XXVIA_TB_salidas_cajas WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux8.EOF Then
            cnn.BeginTrans
            rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) AS CONSECUTIVO FROM TB_TEMP_ORACLE_DETALLE_CAJAS", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rsaux.Close
            rsaux.Open "INSERT INTO TB_TEMP_ORACLE_DETALLE_CAJAS (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            While Not rsaux8.EOF
                  rs.Open "SELECT floa_sal_Cantidad_leida as cantidad, organizacion,  inte_emb_embarque as embarque, inte_paq_caja as caja, source_header_number  as pedido, a.segment1 as codigo, collector_id as agente,name as nombre_agente, customer_id as cliente, customer_name as nombre_cliente, a.inventory_item_id, caja_pedido, sello, UNIT_WEIGHT as peso, item_description as descripcion, tipo_caja    FROM XXVIA_TB_salidas_cajas a, xxvia_tb_encabezado_embarques, xxvia_system_items_b b, oe_order_headers_all oh where inte_emb_embarque = embarque and organizacion = b.organization_id and a.inventory_item_id = b.inventory_item_id and order_number = a.source_header_number and  oh.ship_from_org_id = organizacion and  inte_emb_embarque = " + Me.txt_embarque + " and floa_sal_Cantidad_leida >0 AND SOURCE_HEADER_NUMBER = " + CStr(rsaux8!pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  rsaux.Open "select * from oe_order_headers_all where order_number = " + CStr(rs!pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If rsaux!ORDER_TYPE_ID = 1002 Then
                     var_agente = rs!Agente
                     var_nombre_cliente = rs!nombre_cliente
                     var_nombre_agente = rs!NOMBRE_AGENTE
                     var_cliente = rs!Cliente
                     rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        var_nombre_cliente = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                     End If
                     rsaux2.Close
                  Else
                     var_agente = rs!Agente
                     var_nombre_cliente = rs!nombre_cliente
                     var_nombre_agente = rs!NOMBRE_AGENTE
                     var_cliente = rs!Cliente
                  End If
                  rsaux.Close
                  rsaux.Open "alter session set nls_language= 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rs.EOF
                        'var_cadena = "select ic.category_concat_segs,fft.description from mtl_item_categories_v ic, fnd_flex_value_sets   ffvs, fnd_flex_values_vl ffv,  fnd_flex_values_tl fft,mtl_parameters mtp where UPPER (ic.category_set_name) LIKE '%VIANNEY%EXPORTACION%' AND ic.inventory_item_id = " + CStr(rs!INVENTORY_ITEM_ID) + " AND ic.organization_id = mtp.organization_id AND mtp.organization_code = 'MTO' AND ic.category_concat_segs =  ffv.flex_value_meaning AND ffvs.flex_value_set_name = 'VIANNEY_INV_EXPORTACION' AND ffvs.flex_value_set_id  =  ffv.flex_value_set_id AND ffv.flex_value_id  =  fft.flex_value_id AND fft.language  =  USERENV('LANG') "
                        'If rsaux.State = 1 Then
                        '   rsaux.Close
                        'End If
                        'rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'If Not rsaux.EOF Then
                        '   var_arancel = CStr(IIf(IsNull(rsaux!Description), "", rsaux!Description))
                        'Else
                        '   var_arancel = ""
                        'End If
                        'rsaux.Close
                  
                        var_peso = IIf(IsNull(rs!PESO), 0, rs!PESO)
                  
                        'var_pedido = rs!pedido
                        'rsaux.Open "select unit_selling_price from oe_order_headers_all oh, oe_order_lines_all ol where order_number = " + CStr(var_pedido) + " and oh.header_id = ol.header_id and oh.ship_from_org_id = " + var_unidad_organizacional + " and ol.inventory_item_id = " + CStr(IIf(IsNull(rs!INVENTORY_ITEM_ID), 0, rs!INVENTORY_ITEM_ID)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'If Not rsaux.EOF Then
                        '   var_precio = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                        'Else
                        '   var_precio = 0
                        'End If
                        'rsaux.Close
                        var_precio = 0
                        var_cadena = "INSERT INTO TB_TEMP_ORACLE_DETALLE_CAJAS (INTE_TEM_CONSECUTIVO, EMBARQUE, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, PRECIO, PESO, CAJA, CONTENIDO, INVENTORY_ITEM_ID, ARANCEL, CODIGO_BARRAS, SELLO, CAJA_PEDIDO, PEDIDO, bulto)"
                        var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + "," + CStr(rs!Embarque) + ",'" + CStr(IIf(IsNull(var_agente), "", var_agente)) + "','" + var_nombre_agente + "', '" + CStr(var_cliente) + "','" + IIf(IsNull(var_nombre_cliente), "", var_nombre_cliente) + "','" + rs!codigo + "','" + rs!Descripcion + "'," + CStr(rs!cantidad) + "," + CStr(var_precio) + "," + CStr(var_peso) + "," + CStr(rs!Caja) + ",''," + CStr(rs!inventory_item_id) + ", '" + var_arancel + "','" + var_codigo_barras + "','" + Trim(IIf(IsNull(rs!sello), "", rs!sello)) + "'," + CStr(IIf(IsNull(rs!caja_pedido), rs!Caja, rs!caja_pedido)) + "," + CStr(rs!pedido) + ",'" + IIf(IsNull(rs!tipo_caja), "", rs!tipo_caja) + "')"
                        rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        rs.MoveNext
                  Wend
                  rsaux1.Open "DELETE FROM TB_TEMP_ORACLE_DETALLE_CAJAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND EMBARQUE IS NULL", cnn, adOpenDynamic, adLockOptimistic
            
                  rsaux1.Open "SELECT DISTINCT CAJA FROM TB_TEMP_ORACLE_DETALLE_CAJAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux8!pedido), cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveFirst
                  VAR_CANTIDAD_CAJAS = 0
                  While Not rsaux1.EOF
                        VAR_CANTIDAD_CAJAS = VAR_CANTIDAD_CAJAS + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "UPDATE TB_TEMP_ORACLE_DETALLE_CAJAS SET NUMERO_CAJAS = " + CStr(VAR_CANTIDAD_CAJAS) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux8!pedido), cnn, adOpenDynamic, adLockOptimistic
                  rsaux8.MoveNext
                  rs.Close
            Wend
            rsaux8.Close
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_packing_list_resumen.rpt")
            reporte.RecordSelectionFormula = "{VW_ORACLE_PACKING_LIST_RESUMEN.EMBARQUE} = " + txt_embarque + " and {VW_ORACLE_PACKING_LIST_RESUMEN.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Packing List"
            frmvistasprevias.Show 1
            Set reporte = Nothing
       
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_packing_list_resumen.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_PACKING_LIST_RESUMEN.EMBARQUE} = " + txt_embarque + " and {VW_ORACLE_PACKING_LIST_RESUMEN.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\packing_list_resumen_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
               var_si = MsgBox("Desea enviar el packing list por correo?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  VAR_CORREO_ELECTRONICO = "vluna@vianney.com.mx"
                  If Trim(VAR_CORREO_ELECTRONICO) <> "" Then
                     If MAPISession1.SessionID = 0 Then
                        MAPISession1.SignOn
                     End If
                     MAPIMessages1.SessionID = MAPISession1.SessionID
                     MAPIMessages1.Compose
                     MAPIMessages1.RecipDisplayName = VAR_CORREO_ELECTRONICO
                     MAPIMessages1.RecipAddress = VAR_CORREO_ELECTRONICO
                     MAPIMessages1.AddressResolveUI = True
                     MAPIMessages1.ResolveName
                     MAPIMessages1.MsgSubject = "Packing list"
                     MAPIMessages1.MsgNoteText = "Se anexa archivo de resumen de packing list"
                     MAPIMessages1.AttachmentPathName = archivo
                     MAPIMessages1.send True
                     If MAPISession1.SessionID > 0 Then
                        MAPISession1.SignOff
                     End If
                  End If
               End If
            End If
            rsaux.Open "delete from TB_TEMP_ORACLE_DETALLE_CAJAS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
      Else
         If var_posible = 0 Then
            MsgBox "El embarque no a sido cerrado.", vbOKOnly, "ATENCION"
         End If
         If var_posible = 2 Then
            MsgBox "El embarque no existe.", vbOKOnly, "ATENCION"
         End If
         
      End If
   Else
      MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   If IsNumeric(Me.txt_embarque) Then
      strconsulta = "select char_emb_estatus from xxvia_tb_encabezado_embarques where embarque = ?"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_embarque))
           .Parameters.Append parametro
      End With
      Set rsaux = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      var_posible = 1
      If Not rsaux.EOF Then
         If IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus) = "I" Or IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus) = "F" Or IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus) = "E" Then
            var_posible = 1
         Else
            var_posible = 1
         End If
      Else
         var_posible = 2
      End If
      rsaux.Close
      If var_posible = 1 Then
         cnn.BeginTrans
         rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) AS CONSECUTIVO FROM TB_TEMP_ORACLE_DETALLE_CAJAS", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
         Else
            var_consecutivo = 1
         End If
         rsaux.Close
         rsaux.Open "INSERT INTO TB_TEMP_ORACLE_DETALLE_CAJAS (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
            
         rsaux10.Open "select distinct source_header_number from xxvia_tb_Salidas_cajas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux10.EOF
               rs.Open "SELECT floa_sal_Cantidad_leida as cantidad, organizacion,  inte_emb_embarque as embarque, inte_paq_caja as caja, source_header_number  as pedido, a.segment1 as codigo, collector_id as agente,name as nombre_agente, customer_id as cliente, customer_name as nombre_cliente, a.inventory_item_id, caja_pedido, sello, UNIT_WEIGHT as peso,  item_description as descripcion    FROM XXVIA_TB_salidas_cajas a, xxvia_tb_encabezado_embarques, xxvia_system_items_b b, oe_order_headers_all oh where inte_emb_embarque = embarque and organizacion = b.organization_id and a.inventory_item_id = b.inventory_item_id and order_number = a.source_header_number and nvl(oh.ship_from_org_id,93) = organizacion and inte_emb_embarque = " + Me.txt_embarque + " and floa_sal_Cantidad_leida >0 and source_header_number = " + CStr(rsaux10!source_header_number) + " order by source_header_number, inte_paq_caja", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  rsaux.Open "select * from oe_order_headers_all where order_number = " + CStr(rs!pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If rsaux!ORDER_TYPE_ID = 1002 Then
                     var_agente = rs!Agente
                     var_nombre_cliente = rs!nombre_cliente
                     var_nombre_agente = rs!NOMBRE_AGENTE
                     var_cliente = rs!Cliente
                     rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        var_nombre_cliente = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                     End If
                     rsaux2.Close
                  Else
                     var_agente = rs!Agente
                     var_nombre_cliente = rs!nombre_cliente
                     var_nombre_agente = rs!NOMBRE_AGENTE
                     var_cliente = rs!Cliente
                  End If
                  rsaux.Close
                  rsaux.Open "alter session set nls_language= 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         
                  var_pedido = rs!pedido
                  strconsulta = "select header_id from oe_order_headers_all oh where order_number = ? "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_pedido))
                       .Parameters.Append parametro
                  End With
                  Set rsaux = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  VAR_HEADER_ID = rsaux!header_id
                  rsaux.Close
         
         
                  While Not rs.EOF
                        'var_cadena = "select ic.category_concat_segs,fft.description from mtl_item_categories_v ic, fnd_flex_value_sets   ffvs, fnd_flex_values_vl ffv,  fnd_flex_values_tl fft,mtl_parameters mtp where UPPER (ic.category_set_name) LIKE '%VIANNEY%EXPORTACION%' AND ic.inventory_item_id = " + CStr(rs!INVENTORY_ITEM_ID) + " AND ic.organization_id = mtp.organization_id AND mtp.organization_code = 'MTO' AND ic.category_concat_segs =  ffv.flex_value_meaning AND ffvs.flex_value_set_name = 'VIANNEY_INV_EXPORTACION' AND ffvs.flex_value_set_id  =  ffv.flex_value_set_id AND ffv.flex_value_id  =  fft.flex_value_id AND fft.language  =  USERENV('LANG') "
                        'If rsaux.State = 1 Then
                        '   rsaux.Close
                        'End If
                        'rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'If Not rsaux.EOF Then
                        '   var_arancel = CStr(IIf(IsNull(rsaux!Description), "", rsaux!Description))
                        'Else
                        '   var_arancel = ""
                        'End If
                        'rsaux.Close
               
                        var_peso = IIf(IsNull(rs!PESO), 0, rs!PESO)
                
                        var_pedido = rs!pedido
                  
                        'strconsulta = "select unit_selling_price from oe_order_headers_all oh, oe_order_lines_all ol where order_number = ? and oh.header_id = ol.header_id and oh.ship_from_org_id = ? and ol.inventory_item_id = ?"
                        strconsulta = "select unit_selling_price from oe_order_lines_all ol where header_id = ? and ol.inventory_item_id = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(VAR_HEADER_ID))
                             .Parameters.Append parametro
                             'Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
                             '.Parameters.Append parametro
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
                 
                        'rsaux.Open "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, NVL(a.description,'') AS localizador, B.UNIT_WEIGHT FROM (select INVENTORY_ITEM_ID, description, cross_reference from mtl_cross_references_b) A, (select inventory_item_id, DESCRIPTION, organization_id, segment1, UNIT_WEIGHT from xxvia_system_items_b) B Where a.inventory_item_id = B.inventory_item_id AND B.organization_id = " + var_unidad_organizacional + " AND b.segment1 = '" + IIf(IsNull(rs!codigo), "", rs!codigo) + "' AND SUBSTR(cross_reference,1,3) = '000' and substr(cross_Reference,9,3) = '000'  AND substr(cross_Reference,12,1) IN( '0','1','2','3','4','5','6','7','8','9')"
                        strconsulta = "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, NVL(a.description,'') AS localizador, B.UNIT_WEIGHT FROM (select INVENTORY_ITEM_ID, description, cross_reference from mtl_cross_references_b) A, (select inventory_item_id, DESCRIPTION, organization_id, segment1, UNIT_WEIGHT from xxvia_system_items_b) B Where a.inventory_item_id = B.inventory_item_id AND B.organization_id = ? AND b.segment1 = ? AND SUBSTR(cross_reference,1,3) = '000' and substr(cross_Reference,9,3) = '000'  AND substr(cross_Reference,12,1) IN( '0','1','2','3','4','5','6','7','8','9')"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rs!codigo), "", rs!codigo))
                             .Parameters.Append parametro
                        End With
                        Set rsaux = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
               
                        If Not rsaux.EOF Then
                           var_codigo_barras = IIf(IsNull(rsaux!cross_reference), "", rsaux!cross_reference)
                        Else
                           var_codigo_barras = ""
                        End If
                        rsaux.Close
                  
                  
                        strconsulta = "select * from xxvia_tb_complementos_pk_list where codigo = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rs!codigo), "", rs!codigo))
                            .Parameters.Append parametro
                        End With
                        Set rsaux9 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If Not rsaux9.EOF Then
                           var_arancel = IIf(IsNull(rsaux9!fraccion_arancelaria), 0, rsaux9!fraccion_arancelaria)
                           VAR_COMPOSICION = IIf(IsNull(rsaux9!composicion), "", rsaux9!composicion)
                           VAR_CONTENIDO = IIf(IsNull(rsaux9!contenido), "", rsaux9!contenido)
                           VAR_ORIGEN = IIf(IsNull(rsaux9!originario), "", rsaux9!originario)
                           VAR_APLICA_USA = IIf(IsNull(rsaux9!aplica_usa), "", rsaux9!aplica_usa)
                           VAR_APLICA_CA = IIf(IsNull(rsaux9!aplica_ca), "", rsaux9!aplica_ca)
                           VAR_HECHO_EN = IIf(IsNull(rsaux9!hecho_en), "", rsaux9!hecho_en)
                           VAR_COMPLEMENTO_1 = IIf(IsNull(rsaux9!complemento_1), "", rsaux9!complemento_1)
                           VAR_PRECIO_1 = var_precio * (IIf(IsNull(rsaux9!precio_1), 0, rsaux9!precio_1) / 100)
                           VAR_COMPLEMENTO_2 = IIf(IsNull(rsaux9!complemento_2), "", rsaux9!complemento_2)
                           VAR_PRECIO_2 = var_precio * (IIf(IsNull(rsaux9!precio_2), 0, rsaux9!precio_2) / 100)
                           VAR_COMPLEMENTO_3 = IIf(IsNull(rsaux9!complemento_3), "", rsaux9!complemento_3)
                           VAR_PRECIO_3 = var_precio * (IIf(IsNull(rsaux9!precio_3), 0, rsaux9!precio_3) / 100)
                           VAR_COMPLEMENTO_4 = IIf(IsNull(rsaux9!complemento_4), "", rsaux9!complemento_4)
                           VAR_PRECIO_4 = var_precio * (IIf(IsNull(rsaux9!precio_4), 0, rsaux9!precio_4) / 100)
                           VAR_COMPLEMENTO_5 = IIf(IsNull(rsaux9!complemento_5), "", rsaux9!complemento_5)
                           VAR_PRECIO_5 = var_precio * (IIf(IsNull(rsaux9!precio_5), 0, rsaux9!precio_5) / 100)
                           VAR_COMPLEMENTO_6 = IIf(IsNull(rsaux9!complemento_6), "", rsaux9!complemento_6)
                           VAR_PRECIO_6 = var_precio * (IIf(IsNull(rsaux9!precio_6), 0, rsaux9!precio_6) / 100)
                           VAR_COMPLEMENTO_7 = IIf(IsNull(rsaux9!complemento_7), "", rsaux9!complemento_7)
                           VAR_PRECIO_7 = var_precio * (IIf(IsNull(rsaux9!precio_7), 0, rsaux9!precio_7) / 100)
                           VAR_COMPLEMENTO_8 = IIf(IsNull(rsaux9!complemento_8), "", rsaux9!complemento_8)
                           VAR_PRECIO_8 = var_precio * (IIf(IsNull(rsaux9!precio_8), 0, rsaux9!precio_8) / 100)
                           VAR_COMPLEMENTO_9 = IIf(IsNull(rsaux9!complemento_9), "", rsaux9!complemento_9)
                           VAR_PRECIO_9 = var_precio * (IIf(IsNull(rsaux9!precio_9), 0, rsaux9!precio_9) / 100)
                           VAR_COMPLEMENTO_10 = IIf(IsNull(rsaux9!complemento_10), "", rsaux9!complemento_10)
                           VAR_PRECIO_10 = var_precio * (IIf(IsNull(rsaux9!precio_10), 0, rsaux9!precio_10) / 100)
                           VAR_COMPLEMENTO_11 = IIf(IsNull(rsaux9!complemento_11), "", rsaux9!complemento_11)
                           VAR_PRECIO_11 = var_precio * (IIf(IsNull(rsaux9!precio_11), 0, rsaux9!precio_11) / 100)
                           VAR_COMPLEMENTO_12 = IIf(IsNull(rsaux9!complemento_12), "", rsaux9!complemento_12)
                           VAR_PRECIO_12 = var_precio * (IIf(IsNull(rsaux9!precio_12), 0, rsaux9!precio_12) / 100)
                           VAR_COMPLEMENTO_13 = IIf(IsNull(rsaux9!complemento_13), "", rsaux9!complemento_13)
                           VAR_PRECIO_13 = var_precio * (IIf(IsNull(rsaux9!precio_13), 0, rsaux9!precio_13) / 100)
                           VAR_COMPLEMENTO_14 = IIf(IsNull(rsaux9!complemento_14), "", rsaux9!complemento_14)
                           VAR_PRECIO_14 = var_precio * (IIf(IsNull(rsaux9!precio_14), 0, rsaux9!precio_14) / 100)
                           VAR_COMPLEMENTO_15 = IIf(IsNull(rsaux9!complemento_15), "", rsaux9!complemento_15)
                           VAR_PRECIO_15 = var_precio * (IIf(IsNull(rsaux9!precio_15), 0, rsaux9!precio_15) / 100)
                           VAR_COMPLEMENTO_16 = IIf(IsNull(rsaux9!complemento_16), "", rsaux9!complemento_16)
                           VAR_PRECIO_16 = var_precio * (IIf(IsNull(rsaux9!precio_16), 0, rsaux9!precio_16) / 100)
                           VAR_FRACCION_AMERICANA = IIf(IsNull(rsaux9!fraccion_americana), 0, rsaux9!fraccion_americana)
                           var_criterio_usa = IIf(IsNull(rsaux9!criterio_usa), "", rsaux9!criterio_usa)
                           var_criterio_ca = IIf(IsNull(rsaux9!criterio_ca), "", rsaux9!criterio_ca)
                        Else
                           var_arancel = 0
                           var_comosicion = ""
                           var_contenito = ""
                           VAR_ORIGEN = ""
                           VAR_APLICA_USA = ""
                           VAR_APLICA_CA = ""
                           VAR_HECHO_ENT = ""
                           VAR_COMPLEMENTO_1 = ""
                           VAR_PRECIO_1 = 0
                           VAR_COMPLEMENTO_2 = ""
                           VAR_PRECIO_2 = 0
                           VAR_COMPLEMENTO_3 = ""
                           VAR_PRECIO_3 = 0
                           VAR_COMPLEMENTO_4 = ""
                           VAR_PRECIO_4 = 0
                           VAR_COMPLEMENTO_5 = ""
                           VAR_PRECIO_5 = 0
                           VAR_COMPLEMENTO_6 = ""
                           VAR_PRECIO_6 = 0
                           VAR_COMPLEMENTO_7 = ""
                           VAR_PRECIO_7 = 0
                           VAR_COMPLEMENTO_8 = ""
                           VAR_PRECIO_8 = 0
                           VAR_COMPLEMENTO_9 = ""
                           VAR_PRECIO_9 = 0
                           VAR_COMPLEMENTO_10 = ""
                           VAR_PRECIO_10 = 0
                           VAR_COMPLEMENTO_11 = ""
                           VAR_PRECIO_11 = 0
                           VAR_COMPLEMENTO_12 = ""
                           VAR_PRECIO_12 = 0
                           VAR_COMPLEMENTO_13 = ""
                           VAR_PRECIO_13 = 0
                           VAR_COMPLEMENTO_14 = ""
                           VAR_PRECIO_14 = 0
                           VAR_COMPLEMENTO_15 = ""
                           VAR_PRECIO_15 = 0
                           VAR_COMPLEMENTO_16 = ""
                           VAR_PRECIO_16 = 0
                           VAR_FRACCION_AMERICANA = 0
                           var_criterio_usa = ""
                           var_criterio_ca = ""
                        
                        End If
                        rsaux9.Close
               
                        strconsulta = "SELECT PEDIDO, CAJA, CODIGO, decode(SUBSTR(codigo_barras,9,1), 1,'CHINO','MEXICO') ORIGEN, SUM(CANTIDAD) CANTIDAD FROM XXVIA_TB_BITACORA_LECTURA where pedido = ? and codigo = ? GROUP BY PEDIDO, CAJA, CODIGO, decode(SUBSTR(codigo_barras,9,1), 1,'CHINO','MEXICO')"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(var_pedido), "", var_pedido))
                            .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rs!codigo), "", rs!codigo))
                            .Parameters.Append parametro
                        End With
                        Set rsaux9 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        VAR_CANTIDAD_MEXICO = 0
                        VAR_CANTIDAD_CHINA = 0
                        If Not rsaux9.EOF Then
                           While Not rsaux9.EOF
                                 If rsaux9!ORIGEN = "MEXICO" Then
                                    VAR_CANTIDAD_MEXICO = rsaux9!cantidad
                                 Else
                                    VAR_CANTIDAD_CHINA = rsaux9!cantidad
                                 End If
                                 rsaux9.MoveNext
                           Wend
                        End If
                        rsaux9.Close
                   
                   
                        var_cadena = "INSERT INTO TB_TEMP_ORACLE_DETALLE_CAJAS (INTE_TEM_CONSECUTIVO, EMBARQUE, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, PRECIO, PESO, CAJA, CONTENIDO, INVENTORY_ITEM_ID, ARANCEL, CODIGO_BARRAS, COMPOSICION, ORIGEN, APLICA_USA, APLICA_CA, HECHO_EN, PEDIDO, COMPLEMENTO_1, PRECIO_1, COMPLEMENTO_2, PRECIO_2, COMPLEMENTO_3, PRECIO_3, COMPLEMENTO_4, PRECIO_4, COMPLEMENTO_5, PRECIO_5, COMPLEMENTO_6, PRECIO_6, COMPLEMENTO_7, PRECIO_7, COMPLEMENTO_8, PRECIO_8, COMPLEMENTO_9, PRECIO_9, COMPLEMENTO_10, PRECIO_10, COMPLEMENTO_11, PRECIO_11, COMPLEMENTO_12, PRECIO_12, COMPLEMENTO_13, PRECIO_13, COMPLEMENTO_14, PRECIO_14, COMPLEMENTO_15, PRECIO_15, COMPLEMENTO_16, PRECIO_16, FRACCION_AMERICANA, criterio_usa, criterio_ca, CANTIDAD_MEXICO, CANTIDAD_CHINA)"
                        var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + "," + CStr(rs!Embarque) + ",'" + CStr(IIf(IsNull(var_agente), "", var_agente)) + "','" + var_nombre_agente + "', '" + CStr(var_cliente) + "','" + IIf(IsNull(var_nombre_cliente), "", var_nombre_cliente) + "','" + rs!codigo + "','" + rs!Descripcion + "'," + CStr(rs!cantidad) + "," + CStr(var_precio) + "," + CStr(var_peso) + "," + CStr(rs!Caja) + ",'" + VAR_CONTENIDO + "'," + CStr(rs!inventory_item_id) + ", '" + CStr(var_arancel) + "','" + var_codigo_barras + "','" + VAR_COMPOSICION + "','" + VAR_ORIGEN + "','" + VAR_APLICA_USA + "','" + VAR_APLICA_CA + "', '" + VAR_HECHO_EN + "'," + CStr(rs!pedido) + ",'" + VAR_COMPLEMENTO_1 + "'," + CStr(VAR_PRECIO_1) + ",'" + VAR_COMPLEMENTO_2 + "'," + CStr(VAR_PRECIO_2) + ",'" + VAR_COMPLEMENTO_3 + "'," + CStr(VAR_PRECIO_3) + ",'" + VAR_COMPLEMENTO_4 + "'," + CStr(VAR_PRECIO_4) + ",'" + VAR_COMPLEMENTO_5 + "'," + CStr(VAR_PRECIO_5)
                        var_cadena = var_cadena + ",'" + VAR_COMPLEMENTO_6 + "'," + CStr(VAR_PRECIO_6) + ",'" + VAR_COMPLEMENTO_7 + "'," + CStr(VAR_PRECIO_7) + ",'" + VAR_COMPLEMENTO_8 + "'," + CStr(VAR_PRECIO_8) + ",'" + VAR_COMPLEMENTO_9 + "'," + CStr(VAR_PRECIO_9) + ",'" + VAR_COMPLEMENTO_10 + "'," + CStr(VAR_PRECIO_10) + ",'" + VAR_COMPLEMENTO_11 + "'," + CStr(VAR_PRECIO_11) + ",'" + VAR_COMPLEMENTO_12 + "'," + CStr(VAR_PRECIO_12) + ",'" + VAR_COMPLEMENTO_13 + "'," + CStr(VAR_PRECIO_13) + ",'" + VAR_COMPLEMENTO_14 + "'," + CStr(VAR_PRECIO_14) + ",'" + VAR_COMPLEMENTO_15 + "'," + CStr(VAR_PRECIO_15) + ",'" + VAR_COMPLEMENTO_16 + "'," + CStr(VAR_PRECIO_16) + "," + CStr(VAR_FRACCION_AMERICANA) + ",'" + var_criterio_usa + "','" + var_criterio_ca + "'," + CStr(VAR_CANTIDAD_MEXICO) + "," + CStr(VAR_CANTIDAD_CHINA) + ")"
                        rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               
                        var_cadena = "SELECT b.vcha_tit_titular_id, D.vcha_esb_establecimient_id FROM XXVIA_VW_CLIENTES_PEDIDOS B, OE_ORDER_HEADERS_ALL C, XXVIA_VW_ESTABLECIMIENTOS_PED D, XXVIA_VW_ESTABLECIMIENTOS_PED E Where c.order_number = ? AND c.SOLD_TO_ORG_ID = B.CUST_ACCOUNT_ID AND D.SITE_USE_ID    = C.SHIP_TO_ORG_ID AND E.SITE_USE_ID    = C.INVOICE_TO_ORG_ID"
                        'MsgBox var_cadena
                        strconsulta = var_cadena
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!pedido)
                             .Parameters.Append parametro
                        End With
                        Set rsaux9 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
               
               
                        var_Titular_vianney_catalog = rsaux9!vcha_tit_titular_id
                        var_si_vianney_Catalog = 0
                        var_almacen_Destino = rsaux9!vcha_esb_establecimient_id
                   
                        If var_Titular_vianney_catalog = "T000000343" Then
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
                  
                        var_si_vianney_Catalog = 0
                        If var_si_vianney_Catalog = 1 Then
                           var_cadena = "select * from openquery(icgcentral, 'select * from BDICGDESA_USA.DBO.IT_PEDIDOCOMPRA where no_embarque = ''" + Me.txt_embarque + "'' and no_pedido = ''" + CStr(rs!pedido) + "'' and no_caja =  ''" + CStr(rs!Caja) + "'' and codigo = ''" + rs!codigo + "''') A"
                           var_cadena = "SELECT * FROM OPENQUERY(ICGCENTRAL, 'SELECT * FROM BDICGDESA_USA.DBO.IT_PEDIDOCOMPRA where no_embarque = ''" + Me.txt_embarque + "'' and no_pedido = ''" + CStr(rs!pedido) + "'' and no_caja =  ''" + CStr(rs!Caja) + "'' and codigo = ''" + rs!codigo + "''')A"
                           rsaux10.Open var_cadena, cnn_icg_usa, adOpenDynamic, adLockOptimistic
                           'rsaux10.Open "select * from IT_PEDIDOCOMPRA where no_embarque = " + Me.txt_embarque + " and no_pedido = '" + CStr(rs!pedido) + "' and no_caja =  '" + CStr(rs!Caja) + "' and codigo = '" + rs!codigo + "'", cnn_icg_usa, adOpenDynamic, adLockOptimistic
                           If rsaux10.EOF Then
                              var_cadena = "insert into IT_PEDIDOCOMPRA (OU, SUBINVENTORY_CODE, TRANSFER_SUBINVENTORY, FECHA, NO_EMBARQUE, NO_PEDIDO, NO_CAJA, CODIGO, CANTIDAD, PRECIO, DESCRIPCION) values (381,'" + var_almacen_Destino + "','CDI_ALMPT', " + var_fecha + ", " + Me.txt_embarque + ",'" + CStr(rs!pedido) + "','" + CStr(rs!Caja) + "','" + rs!codigo + "'," + CStr(rs!cantidad) + "," + CStr(var_precio) + ",'" + rs!Descripcion + "')"
                        
                              rsaux11.Open "insert into IT_PEDIDOCOMPRA (OU, SUBINVENTORY_CODE, TRANSFER_SUBINVENTORY, FECHA, NO_EMBARQUE, NO_PEDIDO, NO_CAJA, CODIGO, CANTIDAD, PRECIO, DESCRIPCION) values (381,'" + var_almacen_Destino + "','CDI_ALMPT', " + var_fecha + ", " + Me.txt_embarque + ",'" + CStr(rs!pedido) + "','" + CStr(rs!Caja) + "','" + rs!codigo + "'," + CStr(rs!cantidad) + "," + CStr(var_precio) + ",'" + rs!Descripcion + "')", cnn_icg_usa, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux10.Close
                        End If
                        rs.MoveNext
                  Wend
               Else
                  MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
               End If
               rs.Close
               
               rsaux10.MoveNext
         Wend
         rsaux10.Close
         rsaux1.Open "DELETE FROM TB_TEMP_ORACLE_DETALLE_CAJAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND EMBARQUE IS NULL", cnn, adOpenDynamic, adLockOptimistic
         rsaux1.Open "SELECT DISTINCT CAJA FROM TB_TEMP_ORACLE_DETALLE_CAJAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         rs.Open "SELECT floa_sal_Cantidad_leida as cantidad, organizacion,  inte_emb_embarque as embarque, inte_paq_caja as caja, source_header_number  as pedido, a.segment1 as codigo, collector_id as agente,name as nombre_agente, customer_id as cliente, customer_name as nombre_cliente, a.inventory_item_id, caja_pedido, sello, UNIT_WEIGHT as peso,  item_description as descripcion    FROM XXVIA_TB_salidas_cajas a, xxvia_tb_encabezado_embarques, xxvia_system_items_b b, oe_order_headers_all oh where inte_emb_embarque = embarque and organizacion = b.organization_id and a.inventory_item_id = b.inventory_item_id and order_number = a.source_header_number and nvl(oh.ship_from_org_id,93) = organizacion and inte_emb_embarque = " + Me.txt_embarque + " and floa_sal_Cantidad_leida >0 order by source_header_number, inte_paq_caja", cnnoracle_4, adOpenDynamic, adLockOptimistic
         
         rs.MoveFirst
         VAR_CANTIDAD_CAJAS = 0
         While Not rsaux1.EOF
               VAR_CANTIDAD_CAJAS = VAR_CANTIDAD_CAJAS + 1
               rsaux1.MoveNext
         Wend
         rs.Close
         rsaux1.Close
         rsaux1.Open "UPDATE TB_TEMP_ORACLE_DETALLE_CAJAS SET NUMERO_CAJAS = " + CStr(VAR_CANTIDAD_CAJAS) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         
         rs.Open "SELECT DISTINCT PEDIDO FROM TB_TEMP_ORACLE_DETALLE_CAJAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
         
         While Not rs.EOF
               var_cajas_eliminadas = ""
               rsaux.Open "select caja from TB_ORACLE_CAJAS_ELIMINADAS where pedido = '" + CStr(rs!pedido) + "' order by caja"
               While Not rsaux.EOF
                     If var_cajas_eliminadas = "" Then
                        var_cajas_eliminadas = CStr(rsaux!Caja)
                     Else
                        var_cajas_eliminadas = var_cajas_eliminadas + ", " + CStr(rsaux!Caja)
                     End If
                     rsaux.MoveNext
               Wend
               rsaux.Close
               rsaux.Open "update TB_TEMP_ORACLE_DETALLE_CAJAS set cajas_eliminadas = '" + var_cajas_eliminadas + "' where  INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = '" + CStr(rs!pedido) + "'", cnn, adOpenDynamic, adLockOptimistic
               rs.MoveNext
         Wend
         rs.Close
         
         'Set reporte = appl.OpenReport(App.Path + "\rep_oracle_packing_list_codigo_barras.rpt")
         'reporte.RecordSelectionFormula = "{VW_ORACLE_DETALLE_CAJAS.EMBARQUE} = " + txt_embarque + " and {VW_ORACLE_DETALLE_CAJAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
         'frmvistasprevias.cr.ReportSource = reporte
         'For ntablas = 1 To reporte.Database.Tables.Count
         '   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         'Next ntablas
         'frmvistasprevias.cr.ViewReport
         'frmvistasprevias.Caption = "Packing List"
         'frmvistasprevias.Show 1
         'Set reporte = Nothing
    
         'var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
         'If var_si = 6 Then
         var_si = MsgBox("¿Desea el reporte con complementos?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_packing_list_codigo_barras_complementos.rpt")
            reporte.RecordSelectionFormula = "{VW_ORACLE_DETALLE_CAJAS.EMBARQUE} = " + txt_embarque + " and {VW_ORACLE_DETALLE_CAJAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Packing List"
            frmvistasprevias.Show 1
            Set reporte = Nothing
         
         Else
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_packing_list_codigo_barras.rpt")
         End If
         'var_si = MsgBox("¿Desea el reporte con formato", vbMsgBoxSetForeground, "ATENCION")
         var_aplica_PL = 1
         frmoracle_aplica_PL.Show 1
         If var_si = 6 Then
            'reporte.RecordSelectionFormula = "{VW_ORACLE_DETALLE_CAJAS.EMBARQUE} = " + txt_embarque + " and {VW_ORACLE_DETALLE_CAJAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            'For ntablas = 1 To reporte.Database.Tables.Count
            '    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            'Next ntablas
            'reporte.ExportOptions.FormatType = crEFTExcel80
            'reporte.ExportOptions.DestinationType = crEDTDiskFile
            'archivo = "c:\reportessid\packing_list" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            'reporte.ExportOptions.DiskFileName = archivo
            'reporte.Export False
            'Set reporte = Nothing
         Else
            Dim iFila As Long, iCol As Integer, i As Integer
            Set oexcel = CreateObject("Excel.Application")
            Set owbook = oexcel.Workbooks.Add
            Set osheet = owbook.Worksheets(1)
            osheet.Name = "EMBARQUE " + Me.txt_embarque
            Screen.MousePointer = vbHourglass
            iFila = 1
            ifila2 = 1
            icol2 = 1
            iCol = 1
            'var_cadena = "SELECT EMBARQUE, PEDIDO, CAJA, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, PRECIO, PRECIO * CANTIDAD AS IMPORTE FROM VW_ORACLE_DETALLE_CAJAS EMBARQUE WHERE EMBARQUE = " + txt_embarque + " and INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " ORDER BY PEDIDO, CAJA"
            If var_aplica_PL = 1 Then
               var_cadena = "SELECT CAJA, CODIGO, CODIGO_BARRAS, DESCRIPCION, CANTIDAD, CANTIDAD_MEXICO, CANTIDAD_CHINA,PRECIO, cantidad* precio IMPORTE, PESO, cantidad * peso TOTAL_PESO, ARANCEL FRACCION_ARANCELARIA, FRACCION_AMERICANA A_AMERICANO, CONTENIDO, COMPOSICION, APLICA_USA, HECHO_EN, PEDIDO, EMBARQUE, CAJAS_ELIMINADAS FROM VW_ORACLE_DETALLE_CAJAS EMBARQUE WHERE EMBARQUE = " + txt_embarque + " and INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " ORDER BY PEDIDO, CAJA"
            End If
            If var_aplica_PL = 2 Then
               var_cadena = "SELECT CAJA, CODIGO, CODIGO_BARRAS, DESCRIPCION, CANTIDAD, CANTIDAD_MEXICO, CANTIDAD_CHINA,PRECIO, cantidad* precio IMPORTE, PESO, cantidad * peso TOTAL_PESO, ARANCEL FRACCION_ARANCELARIA, CONTENIDO, COMPOSICION, APLICA_CA,HECHO_EN, PEDIDO, EMBARQUE, CAJAS_ELIMINADAS FROM VW_ORACLE_DETALLE_CAJAS EMBARQUE WHERE EMBARQUE = " + txt_embarque + " and INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " ORDER BY PEDIDO, CAJA"
            End If
            If var_aplica_PL = 3 Then
               var_cadena = "SELECT CAJA, CODIGO, CODIGO_BARRAS, DESCRIPCION, CANTIDAD, CANTIDAD_MEXICO, CANTIDAD_CHINA,PRECIO, cantidad* precio IMPORTE, PESO, cantidad * peso TOTAL_PESO, ARANCEL FRACCION_ARANCELARIA, CONTENIDO, COMPOSICION, FOLIO_COLOMBIA, COMPLEMENTO, CRITERIO_AP,HECHO_EN, PEDIDO, EMBARQUE, CAJAS_ELIMINADAS FROM VW_ORACLE_DETALLE_CAJAS EMBARQUE WHERE EMBARQUE = " + txt_embarque + " and INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " ORDER BY PEDIDO, CAJA"
            End If
            rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            For i = 0 To rsaux10.Fields.Count - 1
                osheet.Cells(iFila, i + 1) = rsaux10.Fields(i).Name
            Next
            iFila = iFila + 1
            With osheet
                 ' carga los registros del recordset
                 .Cells(iFila, iCol).CopyFromRecordset rsaux10
                 'oExcel.Columns(1).Select
                 'oExcel.Selection.NumberFormat = "#,##0.00"
                 'oExcel.Columns(1).Select
                 'oExcel.Selection.Font.Color = vbRed
                 .Columns.AutoFit ' ajusta el ancho de las columnas
            End With
            owbook.SaveAs "c:\reportessid\packing_list" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            oexcel.Visible = True
            Set oexcel = Nothing
            Screen.MousePointer = vbDefault
            rsaux10.Close
         End If
         x = 1
         If x = 0 Then
         Set reporte = appl.OpenReport(App.Path + "\rep_oracle_packing_list_codigo_barras_hoja_trabajo.rpt")
         reporte.RecordSelectionFormula = "{VW_ORACLE_DETALLE_CAJAS.EMBARQUE} = " + txt_embarque + " and {VW_ORACLE_DETALLE_CAJAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         archivo = "c:\reportessid\hoja_trabajo_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
         End If
         MsgBox "Se a terminado de guardar el archivo " + archivo
         var_si = MsgBox("Desea enviar el packing list por correo?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            VAR_CORREO_ELECTRONICO = "vluna@vianney.com.mx"
            If Trim(VAR_CORREO_ELECTRONICO) <> "" Then
               If MAPISession1.SessionID = 0 Then
                  MAPISession1.SignOn
               End If
               MAPIMessages1.SessionID = MAPISession1.SessionID
               MAPIMessages1.Compose
               MAPIMessages1.RecipDisplayName = VAR_CORREO_ELECTRONICO
               MAPIMessages1.RecipAddress = VAR_CORREO_ELECTRONICO
               MAPIMessages1.AddressResolveUI = True
               MAPIMessages1.ResolveName
               MAPIMessages1.MsgSubject = "Packing list"
               MAPIMessages1.MsgNoteText = "Se anexa archivo de packing list"
               MAPIMessages1.AttachmentPathName = archivo
               MAPIMessages1.send True
               If MAPISession1.SessionID > 0 Then
                  MAPISession1.SignOff
               End If
            End If
         End If
         'End If
         rsaux.Open "delete from TB_TEMP_ORACLE_DETALLE_CAJAS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      Else
         If var_posible = 0 Then
            MsgBox "El embarque no a sido cerrado.", vbOKOnly, "ATENCION"
         End If
         If var_posible = 2 Then
            MsgBox "El embarque no existe.", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command2_Click()
      strconsulta = "select cadena as cadena from xxvia_tb_control_doc_fiscales where serie = 'FAEVII' and numero = ?"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(3245))
           .Parameters.Append parametro
      End With
      Set rsaux = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
                      var_cadena = rsaux!Cadena
                      var_cadena_rfc = Mid(var_cadena, 34, 12)
                      VAR_CADENA_STR = ""
                      var_consecutivo = 0
                      For var_i = 1 To Len(var_cadena)
                          If Asc(Mid(var_cadena, var_i, 1)) = 63 Then
                             If Mid(VAR_CADENA_STR, 1, 6) = "LINEA:" Then
                                var_consecutivo = var_consecutivo + 1
                                var_listo = 0
                                VAR_LISTO_2 = 0
                                var_codigo = ""
                                For var_j = 7 To Len(VAR_CADENA_STR)
                                    VAR_LETRA = Mid(VAR_CADENA_STR, var_j, 1)
                                    If VAR_LETRA = "|" Then
                                       If var_listo = 2 Then
                                          var_listo = 3
                                       End If
                                       If var_listo = 0 Then
                                          var_listo = 2
                                       End If
                                    End If
                                    If var_listo = 2 Then
                                       var_codigo = var_codigo + Mid(VAR_CADENA_STR, var_j, 1)
                                    End If
                                Next var_j
                                MsgBox Mid(var_codigo, 2, Len(var_codigo)) + " " + CStr(var_consecutivo)
                             End If
                             VAR_CADENA_STR = ""
                          Else
                             VAR_CADENA_STR = VAR_CADENA_STR + Mid(var_cadena, var_i, 1)
                          End If
                      Next var_i
End Sub

Private Sub Command3_Click()
   If IsNumeric(Me.txt_embarque) Then
      strconsulta = "select char_emb_estatus from xxvia_tb_encabezado_embarques where embarque = ?"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_embarque))
           .Parameters.Append parametro
      End With
      Set rsaux = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      var_posible = 1
      If Not rsaux.EOF Then
         If IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus) = "I" Or IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus) = "F" Or IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus) = "E" Then
            var_posible = 1
         Else
            var_posible = 1
         End If
      Else
         var_posible = 2
      End If
      rsaux.Close
      
      If var_posible = 1 Then
         If rsaux8.State = 1 Then
            rsaux8.Close
         End If
         rsaux8.Open "SELECT DISTINCT SOURCE_HEADER_NUMBER AS PEDIDO FROM XXVIA_TB_salidas_cajas WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         'rs.Open "SELECT inte_emb_embarque as embarque, inte_paq_caja as caja, source_header_number  as pedido, segment1 as codigo, collector_id as agente,name as nombre_agente, customer_id as cliente, customer_name as nombre_cliente, inventory_item_id, caja_pedido, sello, item_description as descripcion, floa_sal_cantidad_leida as cantidad   FROM XXVIA_TB_salidas_cajas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux8.EOF Then
            cnn.BeginTrans
            rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) AS CONSECUTIVO FROM TB_TEMP_ORACLE_DETALLE_CAJAS", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rsaux.Close
            rsaux.Open "INSERT INTO TB_TEMP_ORACLE_DETALLE_CAJAS (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            While Not rsaux8.EOF
               
                  rs.Open "SELECT floa_sal_Cantidad_leida as cantidad, organizacion,  inte_emb_embarque as embarque, inte_paq_caja as caja, source_header_number  as pedido, a.segment1 as codigo, collector_id as agente,name as nombre_agente, customer_id as cliente, customer_name as nombre_cliente, a.inventory_item_id, caja_pedido, sello, UNIT_WEIGHT as peso, item_description as descripcion, tipo_caja    FROM XXVIA_TB_salidas_cajas a, xxvia_tb_encabezado_embarques, xxvia_system_items_b b, oe_order_headers_all oh where inte_emb_embarque = embarque and organizacion = b.organization_id and a.inventory_item_id = b.inventory_item_id and order_number = a.source_header_number and  oh.ship_from_org_id = organizacion and  inte_emb_embarque = " + Me.txt_embarque + " and floa_sal_Cantidad_leida >0 AND SOURCE_HEADER_NUMBER = " + CStr(rsaux8!pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  rsaux.Open "select * from oe_order_headers_all where order_number = " + CStr(rs!pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If rsaux!ORDER_TYPE_ID = 1002 Then
                     var_agente = rs!Agente
                     var_nombre_cliente = rs!nombre_cliente
                     var_nombre_agente = rs!NOMBRE_AGENTE
                     var_cliente = rs!Cliente
                     rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        var_nombre_cliente = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                     End If
                     rsaux2.Close
                  Else
                     var_agente = rs!Agente
                     var_nombre_cliente = rs!nombre_cliente
                     var_nombre_agente = rs!NOMBRE_AGENTE
                     var_cliente = rs!Cliente
                  End If
                  rsaux.Close
                  rsaux.Open "alter session set nls_language= 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rs.EOF
                        'var_cadena = "select ic.category_concat_segs,fft.description from mtl_item_categories_v ic, fnd_flex_value_sets   ffvs, fnd_flex_values_vl ffv,  fnd_flex_values_tl fft,mtl_parameters mtp where UPPER (ic.category_set_name) LIKE '%VIANNEY%EXPORTACION%' AND ic.inventory_item_id = " + CStr(rs!INVENTORY_ITEM_ID) + " AND ic.organization_id = mtp.organization_id AND mtp.organization_code = 'MTO' AND ic.category_concat_segs =  ffv.flex_value_meaning AND ffvs.flex_value_set_name = 'VIANNEY_INV_EXPORTACION' AND ffvs.flex_value_set_id  =  ffv.flex_value_set_id AND ffv.flex_value_id  =  fft.flex_value_id AND fft.language  =  USERENV('LANG') "
                        'If rsaux.State = 1 Then
                        '   rsaux.Close
                        'End If
                        'rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'If Not rsaux.EOF Then
                        '   var_arancel = CStr(IIf(IsNull(rsaux!Description), "", rsaux!Description))
                        'Else
                        '   var_arancel = ""
                        'End If
                        'rsaux.Close
               
                        var_peso = IIf(IsNull(rs!PESO), 0, rs!PESO)
               
                        'var_pedido = rs!pedido
                        'rsaux.Open "select unit_selling_price from oe_order_headers_all oh, oe_order_lines_all ol where order_number = " + CStr(var_pedido) + " and oh.header_id = ol.header_id and oh.ship_from_org_id = " + var_unidad_organizacional + " and ol.inventory_item_id = " + CStr(IIf(IsNull(rs!INVENTORY_ITEM_ID), 0, rs!INVENTORY_ITEM_ID)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'If Not rsaux.EOF Then
                        '   var_precio = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                        'Else
                        '   var_precio = 0
                        'End If
                        'rsaux.Close
                        var_precio = 0
                        
                        'MsgBox Len(rs!sello)
                        var_cadena = "INSERT INTO TB_TEMP_ORACLE_DETALLE_CAJAS (INTE_TEM_CONSECUTIVO, EMBARQUE, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, PRECIO, PESO, CAJA, CONTENIDO, INVENTORY_ITEM_ID, ARANCEL, CODIGO_BARRAS, SELLO, CAJA_PEDIDO, PEDIDO, bulto)"
                        var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + "," + CStr(rs!Embarque) + ",'" + CStr(IIf(IsNull(var_agente), "", var_agente)) + "','" + IIf(IsNull(var_nombre_agente), "", var_nombre_agente) + "', '" + CStr(var_cliente) + "','" + IIf(IsNull(var_nombre_cliente), "", var_nombre_cliente) + "','" + rs!codigo + "','" + Mid(rs!Descripcion, 1, 100) + "'," + CStr(rs!cantidad) + "," + CStr(var_precio) + "," + CStr(var_peso) + "," + CStr(rs!Caja) + ",''," + CStr(rs!inventory_item_id) + ", '" + IIf(IsNull(var_arancel), "", var_arancel) + "','" + IIf(IsNull(var_codigo_barras), "", var_codigo_barras) + "','" + Trim(IIf(IsNull(rs!sello), "", rs!sello)) + "'," + CStr(IIf(IsNull(rs!caja_pedido), rs!Caja, rs!caja_pedido)) + "," + CStr(rs!pedido) + ",'" + IIf(IsNull(rs!tipo_caja), "", rs!tipo_caja) + "')"
                        'MsgBox var_cadena
                        rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        
                        
                        rs.MoveNext
                  Wend
                  rsaux1.Open "DELETE FROM TB_TEMP_ORACLE_DETALLE_CAJAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND EMBARQUE IS NULL", cnn, adOpenDynamic, adLockOptimistic
             
                  rsaux1.Open "SELECT DISTINCT CAJA FROM TB_TEMP_ORACLE_DETALLE_CAJAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux8!pedido), cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveFirst
                  VAR_CANTIDAD_CAJAS = 0
                  While Not rsaux1.EOF
                        VAR_CANTIDAD_CAJAS = VAR_CANTIDAD_CAJAS + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "UPDATE TB_TEMP_ORACLE_DETALLE_CAJAS SET NUMERO_CAJAS = " + CStr(VAR_CANTIDAD_CAJAS) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux8!pedido), cnn, adOpenDynamic, adLockOptimistic
                  rsaux8.MoveNext
                  rs.Close
            Wend
            rsaux8.Close
            rsaux8.Open "SELECT DISTINCT PEDIDO FROM TB_TEMP_ORACLE_DETALLE_CAJAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux8.EOF
                  VAR_CADENA_BULTOS = ""
                  strconsulta = "SELECT tipo_caja, COUNT(*) CANTIDAD FROM XXVIA_VW_CAJAS_POR_PEDIDO WHERE SOURCE_HEADER_NUMBER = ? GROUP BY TIPO_CAJA"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 20, rsaux8!pedido)
                       .Parameters.Append parametro
                  End With
                  Set rsaux9 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  While Not rsaux9.EOF
                        If VAR_CADENA_BULTOS = "" Then
                           VAR_CADENA_BULTOS = rsaux9!tipo_caja + ": " + CStr(rsaux9!cantidad)
                        Else
                           VAR_CADENA_BULTOS = VAR_CADENA_BULTOS + ",    " + rsaux9!tipo_caja + ": " + CStr(rsaux9!cantidad)
                        End If
                        rsaux9.MoveNext
                  Wend
                  rsaux9.Close
                  rsaux9.Open "UPDATE TB_TEMP_ORACLE_DETALLE_CAJAS SET VCHA_PAQ_TIPO_BULTOS = '" + IIf(IsNull(VAR_CADENA_BULTOS), "", VAR_CADENA_BULTOS) + "' WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux8!pedido), cnn, adOpenDynamic, adLockOptimistic
                  rsaux8.MoveNext
            Wend
            rsaux8.Close
         
         
         
            
            
            
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_packing_list_compreso.rpt")
            reporte.RecordSelectionFormula = "{VW_ORACLE_DETALLE_CAJAS.EMBARQUE} = " + txt_embarque + " and {VW_ORACLE_DETALLE_CAJAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Packing List"
            frmvistasprevias.Show 1
            Set reporte = Nothing
       
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_packing_list_compreso.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_DETALLE_CAJAS.EMBARQUE} = " + txt_embarque + " and {VW_ORACLE_DETALLE_CAJAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\packing_list" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
               var_si = MsgBox("Desea enviar el packing list por correo?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  VAR_CORREO_ELECTRONICO = "vluna@vianney.com.mx"
                  If Trim(VAR_CORREO_ELECTRONICO) <> "" Then
                     If MAPISession1.SessionID = 0 Then
                        MAPISession1.SignOn
                     End If
                     MAPIMessages1.SessionID = MAPISession1.SessionID
                     MAPIMessages1.Compose
                     MAPIMessages1.RecipDisplayName = VAR_CORREO_ELECTRONICO
                     MAPIMessages1.RecipAddress = VAR_CORREO_ELECTRONICO
                     MAPIMessages1.AddressResolveUI = True
                     MAPIMessages1.ResolveName
                     MAPIMessages1.MsgSubject = "Packing list"
                     MAPIMessages1.MsgNoteText = "Se anexa archivo de packing list"
                     MAPIMessages1.AttachmentPathName = archivo
                     MAPIMessages1.send True
                     If MAPISession1.SessionID > 0 Then
                        MAPISession1.SignOff
                     End If
                  End If
               End If
            End If
            rsaux.Open "delete from TB_TEMP_ORACLE_DETALLE_CAJAS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
      Else
         If var_posible = 0 Then
            MsgBox "El embarque no a sido cerrado.", vbOKOnly, "ATENCION"
         End If
         If var_posible = 2 Then
            MsgBox "El embarque no existe.", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3850
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub
