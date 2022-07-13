VERSION 5.00
Begin VB.Form frmoracle_correcion_entradas_compra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Corrección entradas por compra"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm 
      Height          =   780
      Left            =   45
      TabIndex        =   0
      Top             =   -45
      Width           =   4170
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   240
         Left            =   60
         TabIndex        =   4
         Top             =   345
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmd_corregir 
         Caption         =   "Corregir"
         Height          =   435
         Left            =   2610
         TabIndex        =   3
         Top             =   210
         Width           =   1455
      End
      Begin VB.TextBox txt_folio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   765
         TabIndex        =   2
         Top             =   195
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   330
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmoracle_correcion_entradas_compra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_folio As Double
Dim var_primera_vez As Integer
Dim var_cantidad_leida As Double
Dim var_fecha_inicio As Date
Dim var_fecha_fin As Date
Dim var_renglon As Integer

Private Sub cmd_corregir_Click()
   Dim objConn As New ADODB.Connection
   Dim objCmd As New ADODB.Command
   Dim objParm As ADODB.Parameter
   
   If IsNumeric(Me.txt_folio) Then
      rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rs.Open "SELECT * FROM XXVIA_TB_RECEPCIONES WHERE FOLIO =  " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         VAR_ESTATUS = IIf(IsNull(rs!ESTATUS), "", rs!ESTATUS)
         If VAR_ESTATUS = "I" Then
            rsaux.Open "select * from rcv_shipment_headers where attribute12 =  'SIDEC_" + Me.txt_folio + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            
            If Not rsaux.EOF Then
               VAR_SHIPMENT_HEADER = rsaux!shipment_header_id
               var_cadena = "SELECT SEGMENT1 AS CODIGO, DESCRIPTION AS NOMBRE_aRTICULO, po_unit_price as precio, currency_conversion_rate as tipo_cambio,A.SHIPMENT_HEADER_ID, a.SHIPMENT_LINE_ID, a.ITEM_ID, SHIPMENT_NUM, A.quantity_shipped AS CANTIDAD_ENVIADA, A.quantity_received AS CANTIDAD_RECIBIDA, B.ship_to_org_id AS TO_ORGANIZATION_ID, B.organization_id AS FROM_ORGANIZATION_ID, '' AS to_subinventory, d.vendor_id, b.attribute13 FROM rcv_shipment_lines A, RCV_SHIPMENT_HEADERS B, xxvia_system_items_b C, RCV_transactions D Where a.SHIPMENT_HEADER_ID = " + CStr(VAR_SHIPMENT_HEADER) + " AND A.shipment_header_id =  B.shipment_header_id AND A.ITEM_ID = C.INVENTORY_ITEM_ID AND A.to_organization_id = C.organization_id and a.SHIPMENT_HEADER_ID = D.SHIPMENT_HEADER_ID and a.shipment_line_id = d.shipment_line_id and d.destination_type_code = 'INVENTORY'"
               rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If rsaux2.EOF Then
                  cnnoracle_4.BeginTrans
                  If rsaux1.State = 1 Then
                     rsaux1.Close
                  End If
                  rsaux1.Open "update xxvia_tb_folios_entradas set folio = folio + 1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "select * from xxvia_tb_folios_entradas", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  var_nuevo_folio = rsaux1!folio
                  rsaux1.Close
                  rsaux1.Open "update xxvia_Tb_recepciones set folio = " + CStr(var_nuevo_folio) + " Where folio = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  cnnoracle_4.CommitTrans
                  rsaux1.Open "select deliver_to_person_id,  PROMISED_DATE, unit_meas_lookup_code, line_num, vendor_id, vendor_site_id, num_oc SHIPMENT_NUM, PO_HEADER_ID SHIPMENT_HEADER_ID, PO_LINE_ID SHIPMENT_LINE_ID,  ITEM_ID, quantity AS QUANTITY_SHIPPED, 0 QUANTITY_RECEIVED, VENDOR_ID, VENDOR_NAME, item_number, ITEM_DESCRIPTION, ORG_ID , SHIP_TO_ORGANIZATION_ID, line_location_id, ship_to_location_id, country_of_origin_code, UOM_CODE, UNIT_PRICE from xxvia_vw_recepcion_compra where num_oc = " + rs!shipment_num + " AND SHIP_TO_ORGANIZATION_ID = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  var_fecha_compromiso = rsaux1!PROMISED_DATE
                  vendor_site_id = rsaux1!vendor_site_id
                  vendor_id = rsaux1!vendor_id
                          
                  rsaux3.Open "SELECT (next_receipt_num+1) idRec From rcv_parameters WHERE organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     var_next_receipt_num = rsaux3(0).Value
                  End If
                  rsaux3.Close
                           
   
                           
                  var_fecha_fin = Now
                  var_segundo_s = CStr(Second(var_fecha_fin))
                  var_minuto_s = CStr(Minute(var_fecha_fin))
                  var_hora_s = CStr(Hour(var_fecha_fin))
                  var_año_s = CStr(Year(var_fecha_fin))
                  var_mes_s = CStr(Month(var_fecha_fin))
                  var_dia_s = CStr(Day(var_fecha_fin))
                  If Len(var_segundo_s) = 1 Then
                     var_segundo_s = "0" + var_segundo_s
                  End If
                  If Len(var_minuto_s) = 1 Then
                     var_minuto_s = "0" + var_minuto_s
                  End If
                  If Len(var_hora_s) = 1 Then
                     var_hora_s = "0" + var_hora_s
                  End If
                  If Len(var_año_s) = 2 Then
                     var_año_s = "20" + var_año_s
                  End If
                  If Len(var_mes_s) = 1 Then
                     var_mes_s = "0" + var_mes_s
                  End If
                  If Len(var_dia_s) = 1 Then
                     var_dia_s = "0" + var_dia_s
                  End If
                  var_fecha_str_1 = var_dia_s + "/" + var_mes_s + "/" + var_año_s
                  var_fecha_str = var_año_s + "/" + var_mes_s + "/" + var_dia_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
                     
                        
                  var_concurrente = 0
                  objConn.Open var_conexion_oracle
                  With objCmd
                       objConn.BeginTrans
                       .ActiveConnection = objConn
                       .CommandText = "XXVIA_PK_INTERFACES_PO.crear_encabezado_recepcion2"
                       .CommandType = adCmdStoredProc
                                   
                       Set objParm = .CreateParameter("p_expected_receipt_date", adVarChar, adParamInput, 50, var_fecha_str_1)
                       .Parameters.Append objParm
                          
                       Set objParm = .CreateParameter("p_ship_to_organization_id", adNumeric, adParamInput, 200, var_unidad_organizacional)
                       .Parameters.Append objParm
                                      
                       Set objParm = .CreateParameter("p_vendor_id", adNumeric, adParamInput, 200, vendor_id)
                       .Parameters.Append objParm
                                       
                       Set objParm = .CreateParameter("p_vendor_site_id", adNumeric, adParamInput, 200, vendor_site_id)
                       .Parameters.Append objParm
                                         
                       Set objParm = .CreateParameter("p_recepcion_sid", adNumeric, adParamInput, 200, CDbl(var_nuevo_folio))
                       .Parameters.Append objParm
                                       
                       Set objParm = .CreateParameter("p_attrib12", adVarChar, adParamInput, 200, "SIDEC_" + CStr(var_nuevo_folio))
                       .Parameters.Append objParm
                                         
                       Set objParm = .CreateParameter("p_attrib13", adVarChar, adParamInput, 200, "FACTURA: " + rs!factura)
                       .Parameters.Append objParm
                                      
                       Set objParm = .CreateParameter("p_attrib14", adVarChar, adParamInput, 200, "")
                       .Parameters.Append objParm
                                        
                       Set objParm = .CreateParameter("x_header_interface_id", adNumeric, adParamOutput, 200, 0)
                       .Parameters.Append objParm
                                       
                       Set objParm = .CreateParameter("x_group_id", adNumeric, adParamOutput, 200, 0)
                       .Parameters.Append objParm
                                          
                       On Error GoTo salir
                       .execute
                           
                       var_header_interface_id = .Parameters("x_header_interface_id").Value
                       objConn.CommitTrans
                                                  
                       var_group_id = .Parameters("x_group_id").Value
                       'objConn.CommitTrans
                                
                  End With
                  Set objConn = Nothing
                  Set objCmd = Nothing
                           
                  var_cadena = "select xc.po_header_id oc_identificador, xc.num_oc oc_numero, xc.po_line_id oc_linea_identificador, xc.line_num oc_linea_numero, xc.item_number articulo_identificador, xc.item_description articulo_descripcion, xc.quantity cantidad_pendiente, xc.unit_meas_lookup_code oc_unidad_medida, xc.uom_code unidad_medida_primaria, xc.unit_price precio_unitario, xc.currency_code moneda, xc.vendor_name proveedor, xc.quantity+tolerance CANTIDAD_MAXIMA, xc.item_id, xc.closed_code, xc.vendor_id, xc.deliver_to_person_id, xc.line_num, xc.line_location_id, xc.ship_to_location_id, xc.country_of_origin_code, xc.vendor_site_id, xc.RELEASE_NUM,  xc.PO_RELEASE_ID, xc.TYPE_LOOKUP_CODE, cantidad, to_subinventory, factura  FROM xxvia_vw_recepcion_compra xc, xxvia_tb_recepciones xr"
                  var_cadena = var_cadena + " where xc.CLOSED_CODE = 'OPEN' AND xc.num_oc = '" + rsaux1!shipment_num + "' AND xc.ship_to_organization_id =  " + var_unidad_organizacional + " AND xc.org_id  =  " + var_empresa + " and xc.org_id = xr.from_organization_id and xc.ship_to_organization_id = xr.to_organization_id and xc.num_oc = xr.shipment_num and xr.folio = " + CStr(var_nuevo_folio) + " and xc.po_line_id = xr.shipment_line_id  AND XR.LINE_LOCATION_ID = XC.LINE_LOCATION_ID"
                                                
                  rsaux3.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rsaux3.EOF
                                 
                        rsaux4.Open "SELECT rcv_transactions_interface_s.NEXTVAL FROM DUAL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux4.EOF Then
                           VAR_INTERFACE_TRANSACTION_ID = rsaux4(0).Value
                        End If
                        rsaux4.Close
                                 
                        var_release_num = CStr(IIf(IsNull(rsaux3!RELEASE_NUM), 0, rsaux3!RELEASE_NUM))
                        If var_release_num = "0" Then
                           var_release_num = "NULL"
                        End If
                        var_cadena = "Insert Into rcv_transactions_interface (interface_transaction_id, GROUP_ID,last_update_date, last_updated_by, creation_date, created_by, last_update_login, transaction_type,transaction_date, processing_status_code, processing_mode_code, transaction_status_code, quantity, unit_of_measure, item_id, employee_id, auto_transact_code, po_header_id, po_line_id, po_line_location_id, receipt_source_code, to_organization_code, source_document_code, document_num, destination_type_code, deliver_to_person_id, deliver_to_location_id, subinventory, header_interface_id, validation_flag, release_num)"
                        var_cadena = var_cadena + " VALUES (" + CStr(VAR_INTERFACE_TRANSACTION_ID) + "," + CStr(var_group_id) + ",SYSDATE, 1170, SYSDATE, 1170, 0, 'RECEIVE', SYSDATE, 'PENDING', 'BATCH', 'PENDING', " + CStr(IIf(IsNull(rsaux3!Cantidad), 0, rsaux3!Cantidad)) + ",'" + CStr(rsaux3!oc_unidad_medida) + "'," + CStr(rsaux3!ITEM_ID) + ",NULL,'DELIVER'," + CStr(rsaux3!oc_identificador) + "," + CStr(rsaux3!oc_linea_identificador) + "," + CStr(rsaux3!line_location_id) + ",'VENDOR','CDI','PO'," + rsaux1!shipment_num + ",'INVENTORY'," + CStr(rsaux3!deliver_to_person_id) + "," + CStr(rsaux3!ship_to_location_id) + ",'" + rsaux3!TO_SUBINVENTORY + "'," + CStr(var_header_interface_id) + ",'Y'," + var_release_num + ")"
                        rsaux4.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux3.MoveNext
                  Wend
                  rsaux3.Close
                           
                        
                        
                       
                  var_concurrente = 0
                  objConn.Open var_conexion_oracle
                  With objCmd
                       objConn.BeginTrans
                       .ActiveConnection = objConn
                       .CommandText = "XXVIA_PK_INVENTARIOS.XXVIA_SP_CONCURRENTE_MAT"
                       .CommandType = adCmdStoredProc
                                
                       Set objParm = .CreateParameter("x_concurrente", adNumeric, adParamOutput, 50, 0)
                       .Parameters.Append objParm
                             
                       Set objParm = .CreateParameter("p_tipo_movimiento", adVarChar, adParamInput, 200, "Traspasos")
                       .Parameters.Append objParm
                                     
                       Set objParm = .CreateParameter("p_organization_id", adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
                       .Parameters.Append objParm
                                       
                       Set objParm = .CreateParameter("p_group_id", adNumeric, adParamInput, 200, var_group_id)
                       .Parameters.Append objParm
                                          
                       On Error GoTo salir
                       .execute
                            
                       var_concurrente = .Parameters("x_concurrente").Value
                       objConn.CommitTrans
                                
                  End With
                  Set objConn = Nothing
                  Set objCmd = Nothing
                      
                         
                        
                  var_mensaje = ""
                  While var_mensaje <> "EXITO"
                        var_concurrente = 0
                        objConn.Open var_conexion_oracle
                        With objCmd
                             objConn.BeginTrans
                             .ActiveConnection = objConn
                             .CommandText = "XXVIA_PK_RECEPCIONES_MP.xxvia_sp_eje_concurr_0"
                             .CommandType = adCmdStoredProc
                                
                             Set objParm = .CreateParameter("p_application", adVarChar, adParamInput, 50, "PO")
                             .Parameters.Append objParm
                             
                             Set objParm = .CreateParameter("p_program", adVarChar, adParamInput, 200, "RCVLCMWS")
                             .Parameters.Append objParm
                                         
                             Set objParm = .CreateParameter("p_description", adVarChar, adParamInput, 200, "Integracion de costo extendido SID")
                             .Parameters.Append objParm
                                         
                             Set objParm = .CreateParameter("p_usuario", adNumeric, adParamInput, 200, 1170)
                             .Parameters.Append objParm
                                         
                             Set objParm = .CreateParameter("p_resp", adNumeric, adParamInput, 200, 20560)
                             .Parameters.Append objParm
                                          
                             Set objParm = .CreateParameter("p_app", adNumeric, adParamInput, 200, 706)
                             .Parameters.Append objParm
                                      
                             Set objParm = .CreateParameter("p_mensaje", adVarChar, adParamOutput, 200, "")
                             .Parameters.Append objParm
                             
                             Set objParm = .CreateParameter("p_concurrente", adNumeric, adParamOutput, 200, 0)
                             .Parameters.Append objParm
                                
                                          
                             On Error GoTo salir
                             .execute
                          
                             var_concurrente = .Parameters("p_concurrente").Value
                     
                             var_mensaje = IIf(IsNull(.Parameters("p_mensaje").Value), "", .Parameters("p_mensaje").Value)
                             objConn.CommitTrans
                             
                        End With
                        Set objConn = Nothing
                        Set objCmd = Nothing
                  Wend
                             
                        
                  var_concurrente = 0
                  objConn.Open var_conexion_oracle
                  With objCmd
                       objConn.BeginTrans
                       .ActiveConnection = objConn
                       .CommandText = "XXVIA_PK_INVENTARIOS.XXVIA_SP_CONCURRENTE_MAT"
                       .CommandType = adCmdStoredProc
                                    
                       Set objParm = .CreateParameter("x_concurrente", adNumeric, adParamOutput, 50, 0)
                       .Parameters.Append objParm
                              
                       Set objParm = .CreateParameter("p_tipo_movimiento", adVarChar, adParamInput, 200, "ImpotarIterface")
                       .Parameters.Append objParm
                                          
                       Set objParm = .CreateParameter("p_organization_id", adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
                       .Parameters.Append objParm
                                         
                       Set objParm = .CreateParameter("p_group_id", adNumeric, adParamInput, 200, var_group_id)
                       .Parameters.Append objParm
                                          
                       On Error GoTo salir
                       .execute
                          
                       var_concurrente = .Parameters("x_concurrente").Value
                       objConn.CommitTrans
                                 
                  End With
                  Set objConn = Nothing
                  Set objCmd = Nothing
                        
                        
                  var_concurrente = 0
                  objConn.Open var_conexion_oracle
                  With objCmd
                       objConn.BeginTrans
                       .ActiveConnection = objConn
                       .CommandText = "XXVIA_PK_RECEPCIONES_MP.XXVIA_SP_ESPERA_DELIVER"
                       .CommandType = adCmdStoredProc
                                   
                       Set objParm = .CreateParameter("p_header_interface_id", adNumeric, adParamInput, 200, var_header_interface_id)
                       .Parameters.Append objParm
                           
                              
                       On Error GoTo salir
                       .execute
                            
                  End With
                  Set objConn = Nothing
                  Set objCmd = Nothing

                  rsaux10.Open "UPDATE XXVIA_TB_RECEPCIONES SET ESTATUS = 'I', fecha_fin = to_date('" + var_fecha_str + "','yyyy/mm/dd hh24:mi:ss') WHERE receipt_source_code = 'VENDOR' AND FOLIO = " + CStr(var_nuevo_folio), cnnoracle_4, adOpenDynamic, adLockOptimistic
                   
                  rsaux10.Open "select * from rcv_shipment_headers where attribute12 =  'SIDEC_" + CStr(var_nuevo_folio) + "'", cnnoracle_4
                  If Not rsaux10.EOF Then
                     MsgBox "Se genero el folio " + CStr(var_nuevo_folio) + ", entre al movimiento y mandelo a imprimir", vbOKOnly, "ATENCION"
                  Else
                      MsgBox "El movimiento no se a terminado de generar en ORACLE, espere un momento por favor y vuelvalo a intentar", vbOKOnly, "ATENCION"
                  End If
                  rsaux10.Close
                  If rsaux1.State = 1 Then
                     rsaux1.Close
                  End If
               Else
                  MsgBox "La recepción no a sido cerrada", vbOKOnly, "ATENCION"
                  Me.txt_folio = ""
               End If
               rsaux2.Close
            Else
'----
               cnnoracle_4.BeginTrans
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
               rsaux1.Open "update xxvia_tb_folios_entradas set folio = folio + 1", cnnoracle_4, adOpenDynamic, adLockOptimistic
               rsaux1.Open "select * from xxvia_tb_folios_entradas", cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_nuevo_folio = rsaux1!folio
               rsaux1.Close
               rsaux1.Open "update xxvia_Tb_recepciones set folio = " + CStr(var_nuevo_folio) + " Where folio = " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
               cnnoracle_4.CommitTrans
               rsaux1.Open "select deliver_to_person_id,  PROMISED_DATE, unit_meas_lookup_code, line_num, vendor_id, vendor_site_id, num_oc SHIPMENT_NUM, PO_HEADER_ID SHIPMENT_HEADER_ID, PO_LINE_ID SHIPMENT_LINE_ID,  ITEM_ID, quantity AS QUANTITY_SHIPPED, 0 QUANTITY_RECEIVED, VENDOR_ID, VENDOR_NAME, item_number, ITEM_DESCRIPTION, ORG_ID , SHIP_TO_ORGANIZATION_ID, line_location_id, ship_to_location_id, country_of_origin_code, UOM_CODE, UNIT_PRICE from xxvia_vw_recepcion_compra where num_oc = " + rs!shipment_num + " AND SHIP_TO_ORGANIZATION_ID = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_fecha_compromiso = rsaux1!PROMISED_DATE
               vendor_site_id = rsaux1!vendor_site_id
               vendor_id = rsaux1!vendor_id
                       
               rsaux4.Open "SELECT (next_receipt_num+1) idRec From rcv_parameters WHERE organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                  var_next_receipt_num = rsaux4(0).Value
               End If
               rsaux4.Close
                        

                        
               var_fecha_fin = Now
               var_segundo_s = CStr(Second(var_fecha_fin))
               var_minuto_s = CStr(Minute(var_fecha_fin))
               var_hora_s = CStr(Hour(var_fecha_fin))
               var_año_s = CStr(Year(var_fecha_fin))
               var_mes_s = CStr(Month(var_fecha_fin))
               var_dia_s = CStr(Day(var_fecha_fin))
               If Len(var_segundo_s) = 1 Then
                  var_segundo_s = "0" + var_segundo_s
               End If
               If Len(var_minuto_s) = 1 Then
                  var_minuto_s = "0" + var_minuto_s
               End If
               If Len(var_hora_s) = 1 Then
                  var_hora_s = "0" + var_hora_s
               End If
               If Len(var_año_s) = 2 Then
                  var_año_s = "20" + var_año_s
               End If
               If Len(var_mes_s) = 1 Then
                  var_mes_s = "0" + var_mes_s
               End If
               If Len(var_dia_s) = 1 Then
                  var_dia_s = "0" + var_dia_s
               End If
               var_fecha_str_1 = var_dia_s + "/" + var_mes_s + "/" + var_año_s
               var_fecha_str = var_año_s + "/" + var_mes_s + "/" + var_dia_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
                 
                        
               var_concurrente = 0
               objConn.Open var_conexion_oracle
               With objCmd
                    objConn.BeginTrans
                    .ActiveConnection = objConn
                    .CommandText = "XXVIA_PK_INTERFACES_PO.crear_encabezado_recepcion2"
                    .CommandType = adCmdStoredProc
                                
                    Set objParm = .CreateParameter("p_expected_receipt_date", adVarChar, adParamInput, 50, var_fecha_str_1)
                    .Parameters.Append objParm
                       
                    Set objParm = .CreateParameter("p_ship_to_organization_id", adNumeric, adParamInput, 200, var_unidad_organizacional)
                    .Parameters.Append objParm
                                   
                    Set objParm = .CreateParameter("p_vendor_id", adNumeric, adParamInput, 200, vendor_id)
                    .Parameters.Append objParm
                                    
                    Set objParm = .CreateParameter("p_vendor_site_id", adNumeric, adParamInput, 200, vendor_site_id)
                    .Parameters.Append objParm
                                      
                    Set objParm = .CreateParameter("p_recepcion_sid", adNumeric, adParamInput, 200, CDbl(var_nuevo_folio))
                    .Parameters.Append objParm
                                    
                    Set objParm = .CreateParameter("p_attrib12", adVarChar, adParamInput, 200, "SIDEC_" + CStr(var_nuevo_folio))
                    .Parameters.Append objParm
                                      
                    Set objParm = .CreateParameter("p_attrib13", adVarChar, adParamInput, 200, "FACTURA: " + rs!factura)
                    .Parameters.Append objParm
                                   
                    Set objParm = .CreateParameter("p_attrib14", adVarChar, adParamInput, 200, "")
                    .Parameters.Append objParm
                                     
                    Set objParm = .CreateParameter("x_header_interface_id", adNumeric, adParamOutput, 200, 0)
                    .Parameters.Append objParm
                                       
                    Set objParm = .CreateParameter("x_group_id", adNumeric, adParamOutput, 200, 0)
                    .Parameters.Append objParm
                                         
                    On Error GoTo salir
                    .execute
                           
                    var_header_interface_id = .Parameters("x_header_interface_id").Value
                    objConn.CommitTrans
                                               
                    var_group_id = .Parameters("x_group_id").Value
                    'objConn.CommitTrans
                                
               End With
               Set objConn = Nothing
               Set objCmd = Nothing
                        
               var_cadena = "select xc.po_header_id oc_identificador, xc.num_oc oc_numero, xc.po_line_id oc_linea_identificador, xc.line_num oc_linea_numero, xc.item_number articulo_identificador, xc.item_description articulo_descripcion, xc.quantity cantidad_pendiente, xc.unit_meas_lookup_code oc_unidad_medida, xc.uom_code unidad_medida_primaria, xc.unit_price precio_unitario, xc.currency_code moneda, xc.vendor_name proveedor, xc.quantity+tolerance CANTIDAD_MAXIMA, xc.item_id, xc.closed_code, xc.vendor_id, xc.deliver_to_person_id, xc.line_num, xc.line_location_id, xc.ship_to_location_id, xc.country_of_origin_code, xc.vendor_site_id, xc.RELEASE_NUM,  xc.PO_RELEASE_ID, xc.TYPE_LOOKUP_CODE, cantidad, to_subinventory, factura  FROM xxvia_vw_recepcion_compra xc, xxvia_tb_recepciones xr"
               var_cadena = var_cadena + " where xc.CLOSED_CODE = 'OPEN' AND xc.num_oc = '" + rsaux1!shipment_num + "' AND xc.ship_to_organization_id =  " + var_unidad_organizacional + " AND xc.org_id  =  " + var_empresa + " and xc.org_id = xr.from_organization_id and xc.ship_to_organization_id = xr.to_organization_id and xc.num_oc = xr.shipment_num and xr.folio = " + CStr(var_nuevo_folio) + " and xc.po_line_id = xr.shipment_line_id  AND XR.LINE_LOCATION_ID = XC.LINE_LOCATION_ID"
                                             
               rsaux3.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               While Not rsaux3.EOF
                               
                     rsaux5.Open "SELECT rcv_transactions_interface_s.NEXTVAL FROM DUAL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux5.EOF Then
                        VAR_INTERFACE_TRANSACTION_ID = rsaux5(0).Value
                     End If
                     rsaux5.Close
                              
                     var_release_num = CStr(IIf(IsNull(rsaux3!RELEASE_NUM), 0, rsaux3!RELEASE_NUM))
                     If var_release_num = "0" Then
                        var_release_num = "NULL"
                     End If
                     var_cadena = "Insert Into rcv_transactions_interface (interface_transaction_id, GROUP_ID,last_update_date, last_updated_by, creation_date, created_by, last_update_login, transaction_type,transaction_date, processing_status_code, processing_mode_code, transaction_status_code, quantity, unit_of_measure, item_id, employee_id, auto_transact_code, po_header_id, po_line_id, po_line_location_id, receipt_source_code, to_organization_code, source_document_code, document_num, destination_type_code, deliver_to_person_id, deliver_to_location_id, subinventory, header_interface_id, validation_flag, release_num)"
                     var_cadena = var_cadena + " VALUES (" + CStr(VAR_INTERFACE_TRANSACTION_ID) + "," + CStr(var_group_id) + ",SYSDATE, 1170, SYSDATE, 1170, 0, 'RECEIVE', SYSDATE, 'PENDING', 'BATCH', 'PENDING', " + CStr(IIf(IsNull(rsaux3!Cantidad), 0, rsaux3!Cantidad)) + ",'" + CStr(rsaux3!oc_unidad_medida) + "'," + CStr(rsaux3!ITEM_ID) + ",NULL,'DELIVER'," + CStr(rsaux3!oc_identificador) + "," + CStr(rsaux3!oc_linea_identificador) + "," + CStr(rsaux3!line_location_id) + ",'VENDOR','CDI','PO'," + rsaux1!shipment_num + ",'INVENTORY'," + CStr(rsaux3!deliver_to_person_id) + "," + CStr(rsaux3!ship_to_location_id) + ",'" + rsaux3!TO_SUBINVENTORY + "'," + CStr(var_header_interface_id) + ",'Y'," + var_release_num + ")"
                     rsaux5.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     rsaux3.MoveNext
               Wend
               rsaux3.Close
                        
                     
                    
                    
               var_concurrente = 0
               objConn.Open var_conexion_oracle
               With objCmd
                    objConn.BeginTrans
                    .ActiveConnection = objConn
                    .CommandText = "XXVIA_PK_INVENTARIOS.XXVIA_SP_CONCURRENTE_MAT"
                    .CommandType = adCmdStoredProc
                             
                    Set objParm = .CreateParameter("x_concurrente", adNumeric, adParamOutput, 50, 0)
                    .Parameters.Append objParm
                          
                    Set objParm = .CreateParameter("p_tipo_movimiento", adVarChar, adParamInput, 200, "Traspasos")
                    .Parameters.Append objParm
                                     
                    Set objParm = .CreateParameter("p_organization_id", adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
                    .Parameters.Append objParm
                                   
                    Set objParm = .CreateParameter("p_group_id", adNumeric, adParamInput, 200, var_group_id)
                    .Parameters.Append objParm
                                       
                    On Error GoTo salir
                    .execute
                         
                    var_concurrente = .Parameters("x_concurrente").Value
                    objConn.CommitTrans
                               
               End With
               Set objConn = Nothing
               Set objCmd = Nothing
                    
                        
                        
               var_mensaje = ""
               While var_mensaje <> "EXITO"
                     var_concurrente = 0
                     objConn.Open var_conexion_oracle
                     With objCmd
                          objConn.BeginTrans
                          .ActiveConnection = objConn
                          .CommandText = "XXVIA_PK_RECEPCIONES_MP.xxvia_sp_eje_concurr_0"
                          .CommandType = adCmdStoredProc
                             
                          Set objParm = .CreateParameter("p_application", adVarChar, adParamInput, 50, "PO")
                          .Parameters.Append objParm
                          
                          Set objParm = .CreateParameter("p_program", adVarChar, adParamInput, 200, "RCVLCMWS")
                          .Parameters.Append objParm
                                      
                          Set objParm = .CreateParameter("p_description", adVarChar, adParamInput, 200, "Integracion de costo extendido SID")
                          .Parameters.Append objParm
                                        
                          Set objParm = .CreateParameter("p_usuario", adNumeric, adParamInput, 200, 1170)
                          .Parameters.Append objParm
                                         
                          Set objParm = .CreateParameter("p_resp", adNumeric, adParamInput, 200, 20560)
                          .Parameters.Append objParm
                                         
                          Set objParm = .CreateParameter("p_app", adNumeric, adParamInput, 200, 706)
                          .Parameters.Append objParm
                                      
                          Set objParm = .CreateParameter("p_mensaje", adVarChar, adParamOutput, 200, "")
                          .Parameters.Append objParm
                            
                          Set objParm = .CreateParameter("p_concurrente", adNumeric, adParamOutput, 200, 0)
                          .Parameters.Append objParm
                             
                                       
                          On Error GoTo salir
                          .execute
                       
                          var_concurrente = .Parameters("p_concurrente").Value
                    
                          var_mensaje = IIf(IsNull(.Parameters("p_mensaje").Value), "", .Parameters("p_mensaje").Value)
                          objConn.CommitTrans
                          
                     End With
                     Set objConn = Nothing
                     Set objCmd = Nothing
               Wend
                             
                        
               var_concurrente = 0
               objConn.Open var_conexion_oracle
               With objCmd
                    objConn.BeginTrans
                    .ActiveConnection = objConn
                    .CommandText = "XXVIA_PK_INVENTARIOS.XXVIA_SP_CONCURRENTE_MAT"
                    .CommandType = adCmdStoredProc
                                 
                    Set objParm = .CreateParameter("x_concurrente", adNumeric, adParamOutput, 50, 0)
                    .Parameters.Append objParm
                           
                    Set objParm = .CreateParameter("p_tipo_movimiento", adVarChar, adParamInput, 200, "ImpotarIterface")
                    .Parameters.Append objParm
                                       
                    Set objParm = .CreateParameter("p_organization_id", adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
                    .Parameters.Append objParm
                                         
                    Set objParm = .CreateParameter("p_group_id", adNumeric, adParamInput, 200, var_group_id)
                    .Parameters.Append objParm
                                       
                    On Error GoTo salir
                    .execute
                          
                    var_concurrente = .Parameters("x_concurrente").Value
                    objConn.CommitTrans
                                 
               End With
               Set objConn = Nothing
               Set objCmd = Nothing
                        
                        
               var_concurrente = 0
               objConn.Open var_conexion_oracle
               With objCmd
                    objConn.BeginTrans
                    .ActiveConnection = objConn
                    .CommandText = "XXVIA_PK_RECEPCIONES_MP.XXVIA_SP_ESPERA_DELIVER"
                    .CommandType = adCmdStoredProc
                                
                    Set objParm = .CreateParameter("p_header_interface_id", adNumeric, adParamInput, 200, var_header_interface_id)
                    .Parameters.Append objParm
                           
                              
                    On Error GoTo salir
                    .execute
                         
               End With
               Set objConn = Nothing
               Set objCmd = Nothing
               rsaux10.Open "UPDATE XXVIA_TB_RECEPCIONES SET ESTATUS = 'I', fecha_fin = to_date('" + var_fecha_str + "','yyyy/mm/dd hh24:mi:ss') WHERE receipt_source_code = 'VENDOR' AND FOLIO = " + CStr(var_nuevo_folio), cnnoracle_4, adOpenDynamic, adLockOptimistic
                   
               rsaux10.Open "select * from rcv_shipment_headers where attribute12 =  'SIDEC_" + CStr(var_nuevo_folio) + "'", cnnoracle_4
               If Not rsaux10.EOF Then
                  MsgBox "Se genero el folio " + CStr(var_nuevo_folio) + ", entre al movimiento y mandelo a imprimir", vbOKOnly, "ATENCION"
               Else
                  MsgBox "El movimiento no se a terminado de generar en ORACLE, espere un momento por favor y vuelvalo a intentar", vbOKOnly, "ATENCION"
               End If
               rsaux10.Close
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If

            
'----
            
            End If
            rsaux.Close
         Else
            MsgBox "La recepción no a sido cerrada", vbOKOnly, "ATENCION"
            Me.txt_folio = ""
         End If
      Else
         Me.txt_folio = ""
         MsgBox "La recepción número " + Me.txt_folio + " no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      Me.txt_folio = ""
      MsgBox "Folio incorrecto", vbOKOnly, "ATENCION"
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
      If rsaux10.State = 1 Then
         rsaux10.Close
      End If
    
    End If
                     
                     
                     
                     
                     

End Sub



Private Sub Command1_Click()
   Call cantidad_leida_por_persona(20, "-")
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_lineas)
End Sub

Private Sub txt_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_corregir.SetFocus
   End If
End Sub

Private Sub txt_folio_LostFocus()
   Dim objConn As New ADODB.Connection
   Dim objCmd As New ADODB.Command
   Dim objParm As ADODB.Parameter
   
   If IsNumeric(Me.txt_folio) Then
      rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rs.Open "SELECT * FROM XXVIA_TB_RECEPCIONES WHERE FOLIO =  " + Me.txt_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         VAR_ESTATUS = IIf(IsNull(rs!ESTATUS), "", rs!ESTATUS)
         If VAR_ESTATUS = "I" Then
            rsaux.Open "select * from rcv_shipment_headers where attribute12 =  'SIDEC_" + Me.txt_folio + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               VAR_SHIPMENT_HEADER = rsaux!shipment_header_id
               var_cadena = "SELECT SEGMENT1 AS CODIGO, DESCRIPTION AS NOMBRE_aRTICULO, po_unit_price as precio, currency_conversion_rate as tipo_cambio,A.SHIPMENT_HEADER_ID, a.SHIPMENT_LINE_ID, a.ITEM_ID, SHIPMENT_NUM, A.quantity_shipped AS CANTIDAD_ENVIADA, A.quantity_received AS CANTIDAD_RECIBIDA, B.ship_to_org_id AS TO_ORGANIZATION_ID, B.organization_id AS FROM_ORGANIZATION_ID, '' AS to_subinventory, d.vendor_id, b.attribute13 FROM rcv_shipment_lines A, RCV_SHIPMENT_HEADERS B, xxvia_system_items_b C, RCV_transactions D Where a.SHIPMENT_HEADER_ID = " + CStr(VAR_SHIPMENT_HEADER) + " AND A.shipment_header_id =  B.shipment_header_id AND A.ITEM_ID = C.INVENTORY_ITEM_ID AND A.to_organization_id = C.organization_id and a.SHIPMENT_HEADER_ID = D.SHIPMENT_HEADER_ID and a.shipment_line_id = d.shipment_line_id and d.destination_type_code = 'INVENTORY'"
               rsaux1.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  MsgBox "La recepción no tiene problemas", vbOKOnly, "ATENCION"
                  Me.txt_folio = ""
               End If
               rsaux1.Close
            End If
            rsaux.Close
         Else
            MsgBox "La recepción no a sido cerrada", vbOKOnly, "ATENCION"
            Me.txt_folio = ""
         End If
      Else
         Me.txt_folio = ""
         MsgBox "La recepción número " + Me.txt_folio + " no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      Me.txt_folio = ""
      MsgBox "Folio incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub
