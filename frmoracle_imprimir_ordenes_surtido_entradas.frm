VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_imprimir_ordenes_surtido_entradas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir OS por entrega"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   " Pedido "
      Height          =   735
      Left            =   60
      TabIndex        =   2
      Top             =   45
      Width           =   2550
      Begin VB.TextBox txt_pedido 
         Height          =   390
         Left            =   90
         TabIndex        =   0
         Top             =   255
         Width           =   2370
      End
   End
   Begin MSComctlLib.ListView lv_entradas 
      Height          =   2205
      Left            =   60
      TabIndex        =   1
      Top             =   855
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   3889
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Entrada"
         Object.Width           =   4410
      EndProperty
   End
End
Attribute VB_Name = "frmoracle_imprimir_ordenes_surtido_entradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub lv_entradas_KeyPress(KeyAscii As Integer)
   If Me.lv_entradas.ListItems.Count > 0 Then
      If IsNumeric(Me.lv_entradas.selectedItem) Then
         var_cadena = " SELECT HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, a.source_header_type_name, oha.source_document_id, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.LAST_UPDATE_DATE,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, C.ATTRIBUTE2 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND  to_number(source_header_number)  BETWEEN " + CStr(Me.txt_pedido) + " AND " + CStr(Me.txt_pedido) + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID "
         var_cadena = var_cadena + " AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y' and A.delivery_id = " + Me.lv_entradas.selectedItem
         
         
         rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            cnn.BeginTrans
            rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_tEMP_ORACLE_ORDEN_SURTIDO", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rsaux.Close
            rsaux1.Open "insert into tb_Temp_oracle_orden_surtido (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            While Not rs.EOF
                  VAR_ESTABLECIMIENTO = rs!ship_to_org_id
                  rsaux.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'SHIP_TO' AND csu.LOCATION = ps.party_site_number and csu.site_use_id = " + CStr(VAR_ESTABLECIMIENTO), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux!VCHA_eSB_NOMBRE), "", rsaux!VCHA_eSB_NOMBRE)
                  Else
                     VAR_NOMBRE_ESTABLECIMIENTO = ""
                  End If
                  rsaux.Close
                  
                  var_dia = CStr(Day(CDate(rs!LAST_UPDATE_DATE)))
                  var_mes = CStr(Month(CDate(rs!LAST_UPDATE_DATE)))
                  var_año = CStr(Year(CDate(rs!LAST_UPDATE_DATE)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                  If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rs!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                     If var_pedido_tienda = 0 Then
                        rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rs!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID) + " AND SITE_USE_ID = " + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           VAR_PROVEEDOR = rsaux4!COLlECTOR_ID
                           VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                           rsaux4.Close
                        Else
                           rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID) + " AND SITE_USE_ID = " + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           VAR_PROVEEDOR = rsaux4!COLlECTOR_ID
                           VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                          rsaux4.Close
                        End If
                        rsaux2.Close
                     Else
                        rsaux2.Open "select a.attribute8, B.description from oe_order_headers_all a, MTL_SECONDARY_INVENTORIES b where order_number = " + CStr(rs!source_header_number) + " and a.attribute8 = b.secondary_inventory_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           VAR_PROVEEDOR = rsaux4!COLlECTOR_ID
                           VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                           rsaux4.Close
                        Else
                           rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           VAR_PROVEEDOR = rsaux4!COLlECTOR_ID
                           VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                           rsaux4.Close
                        End If
                        rsaux2.Close
                     End If
                  Else
                     rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID) + " AND SITE_USE_ID = " + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     VAR_PROVEEDOR = rsaux4!COLlECTOR_ID
                     VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                     rsaux4.Close
                  End If
                  var_cadena = "insert into tb_temp_oracle_orden_surtido (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, DELIVERY_ID, DELIVERY_DETAIL_ID, ORGANIZATION_ID, SUBINVENTORY, DELIVERY_LINE_ID, INVENTORY_ITEM_ID, ITEM_DESCRIPTION, SOURCE_LINE_NUMBER, src_requested_quantity, RELEASED_STATUS, CUSTOMER_NAME, SEGMENT1, COLLECTOR_ID, NAME, date_requested, UBICACION, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO)  values "
                  var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rs!source_header_number), "", rs!source_header_number) + "', " + CStr(IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)) + ", " + CStr(IIf(IsNull(rs!DELIVERY_DETAIL_ID), 0, rs!DELIVERY_DETAIL_ID)) + ", " + CStr(IIf(IsNull(rs!organization_id), 0, rs!organization_id)) + ", '" + IIf(IsNull(rs!subinventory), "", rs!subinventory) + "', " + CStr(IIf(IsNull(rs!delivery_line_id), 0, rs!delivery_line_id)) + ", " + CStr(IIf(IsNull(rs!INVENTORY_ITEM_ID), "", rs!INVENTORY_ITEM_ID)) + ", '" + IIf(IsNull(rs!Description), "", rs!Description) + "', '" + IIf(IsNull(rs!source_line_number), "", rs!source_line_number) + "', " + CStr(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)) + ", '" + IIf(IsNull(rs!released_status), "", rs!released_status) + "', '" + IIf(IsNull(rs!CUSTOMER_NAME), "", rs!CUSTOMER_NAME) + "', '" + IIf(IsNull(rs!segment1), "", rs!segment1) + "'"
                  var_cadena = var_cadena + ", " + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "'," + var_fecha + ",'" + IIf(IsNull(rs!ATTRIBUTE2), "", rs!ATTRIBUTE2) + "','" + CStr(VAR_ESTABLECIMIENTO) + "','" + VAR_NOMBRE_ESTABLECIMIENTO + "')"
                  rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rsaux.Open "select distinct source_header_number from tb_temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER is not null", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               While Not rsaux.EOF
                     Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido.rpt")
                     reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.SOURCE_HEADER_NUMBER} = '" + rsaux(0).Value + "' and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.PrintOut False
                     Set reporte = Nothing
                     x = 0
                     If x = 1 Then
                        Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido.rpt")
                        reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.SOURCE_HEADER_NUMBER} = '" + rsaux(0).Value + "' and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                        For ntablas = 1 To reporte.Database.Tables.Count
                            reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                        Next ntablas
                        reporte.ExportOptions.FormatType = crEFTExcel80
                        reporte.ExportOptions.DestinationType = crEDTDiskFile
                        archivo = "c:\reportessid\ORDEN_SURTIDO_" + CStr(rsaux(0).Value) & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                        reporte.ExportOptions.DiskFileName = archivo
                        reporte.Export False
                        Set reporte = Nothing
                     End If
                     rsaux.MoveNext
               Wend
            End If
            rsaux.Close
         End If
         rs.Close
      Else
         MsgBox "Número de entrega incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No existen entrega por imprimir"
   End If
End Sub

Private Sub txt_pedido_Change()
   Me.lv_entradas.ListItems.Clear
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_pedido) Then
         rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_cadena = "SELECT DISTINCT A.delivery_id from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND  to_number(source_header_number)  BETWEEN " + Me.txt_pedido + " AND " + Me.txt_pedido + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID "
         var_cadena = var_cadena + " AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y' order by A.delivery_id desc"
         rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Dim list_item As ListItem
            While Not rs.EOF
                  Set list_item = lv_entradas.ListItems.Add(, , rs!delivery_id)
                  rs.MoveNext
            Wend
            Me.lv_entradas.SetFocus
         Else
            MsgBox "El número de pedido no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Número de pedido incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub
