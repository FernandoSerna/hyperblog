VERSION 5.00
Begin VB.Form frmoracle_reporte_reservas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reservas"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   2265
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_generar_reporte 
      Caption         =   "Generar reporte"
      Height          =   495
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   1920
   End
End
Attribute VB_Name = "frmoracle_reporte_reservas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub cmd_generar_reporte_Click()
   cnn.BeginTrans
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open "select max(inte_tem_consecutivo) from tb_temp_oracle_reservas", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
   Else
      var_consecutivo = 1
   End If
   rs.Close
   rs.Open "insert into tb_temp_oracle_reservas (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
   cnn.CommitTrans
   var_cadena = "SELECT mr.requirement_date AS fecha, HCAS.CUST_ACCOUNT_ID, TL.NAME AS tipo_pedido, oha.source_document_id, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID, HPS.LOCATION_ID, HL.ADDRESS1 AS cliente, MR.inventory_item_id, to_number(oha.order_number) AS pedido, c.description AS descripcion, mr.reservation_quantity AS cantidad, c.segment1 as codigo, j.name as ruta, u.user_name as usuario, u.description as nombre_usuario, mr.subinventory_code AS ALMACEN FROM hz_cust_acct_sites_all HCAS, OE_ORDER_LINES_ALL OLA, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, MTL_RESERVATIONS MR, OE_TRANSACTION_TYPES_TL TL, mtl_parameters mpa, xxvia_vendedores j, fnd_user u Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID = OHA.INVOICE_TO_ORG_ID AND OLA.HEADER_ID = OHA.HEADER_ID  AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND "
   var_cadena = var_cadena + " MR.inventory_item_id   = c.inventory_item_id AND OHA.ship_from_org_id   = C.ORGANIZATION_ID AND mpa.organization_code  = 'CDI'"
   var_cadena = var_cadena + " AND OLA.LINE_ID  = MR.ORIG_DEMAND_SOURCE_LINE_ID and oha.created_by  = u.user_id"
   var_cadena = var_cadena + " AND OHA.ORDER_TYPE_ID = TL.TRANSAcTION_TYPE_ID  AND TL.language = 'ESA' AND c.organization_id = mpa.organization_id AND mpa.organization_code  = 'CDI' AND TL.NAME NOT IN ('VIA_VXT_PUBLICIDAD','VIA_PEDIDO_INTERNO') and oha.salesrep_id = j.salesrep_id order by to_number(order_number)"
   rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            var_dia_str = CStr(Day(rs!Fecha))
            var_mes_str = CStr(Month(rs!Fecha))
            var_año_str = CStr(Year(rs!Fecha))
            If Len(var_dia_str) = 1 Then
               var_dia_str = "0" + var_dia_str
            End If
            If Len(var_mes_str) = 1 Then
               var_mes_str = "0" + var_mes_str
            End If
            If Len(var_año_str) = 2 Then
               var_año_str = "20" + var_año_str
            End If
            var_fecha = "{d '" + var_año_str + "-" + var_mes_str + "-" + var_dia_str + "'}"
            'MsgBox "insert into tb_temp_oracle_reservas (inte_tem_consecutivo, tipo_pedido, pedido, fecha, cliente, codigo, descripcion, cantidad) values (" + CStr(var_consecutivo) + ",'" + IIf(IsNull(rs!tipo_pedido), "", rs!tipo_pedido) + "','" + CStr(rs!pedido) + "'," + var_fecha + ",'" + IIf(IsNull(rs!Cliente), "", rs!Cliente) + "','" + IIf(IsNull(rs!codigo), "", rs!codigo) + "','" + IIf(IsNull(rs!descripcion), "", rs!descripcion) + "'," + CStr(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) + ")"
            rsaux.Open "insert into tb_temp_oracle_reservas (inte_tem_consecutivo, tipo_pedido, pedido, fecha, cliente, codigo, descripcion, cantidad, RUTA, USUARIO, NOMBRE_USUARIO, CLAVE_ALMACEN) values (" + CStr(var_consecutivo) + ",'" + IIf(IsNull(rs!tipo_pedido), "", rs!tipo_pedido) + "','" + CStr(rs!pedido) + "'," + var_fecha + ",'" + Replace(IIf(IsNull(rs!Cliente), "", rs!Cliente), "'", " ") + "','" + IIf(IsNull(rs!codigo), "", rs!codigo) + "','" + Replace(IIf(IsNull(rs!Descripcion), "", rs!Descripcion), "'", " ") + "'," + CStr(IIf(IsNull(rs!cantidad), 0, rs!cantidad)) + ",'" + IIf(IsNull(rs!ruta), "", rs!ruta) + "','" + IIf(IsNull(rs!USUARIO), "", rs!USUARIO) + "','" + IIf(IsNull(rs!nombre_usuario), "", rs!nombre_usuario) + "','" + rs!ALMACEN + "')", cnn, adOpenDynamic, adLockOptimistic
            rs.MoveNext
      Wend
   End If
   rs.Close
   var_cadena = "SELECT OHA.ship_from_org_id, mr.requirement_date AS fecha, HCAS.CUST_ACCOUNT_ID, 'VIA_PEDIDO_INTERNO' AS tipo_pedido, oha.source_document_id, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID, HPS.LOCATION_ID, inv.description AS cliente, MR.inventory_item_id, to_number(oha.order_number) AS pedido, c.description AS descripcion, mr.reservation_quantity AS cantidad, c.segment1 AS codigo, j.name AS ruta, u.user_name AS usuario, u.description AS nombre_usuario, mr.subinventory_code as almacen FROM hz_cust_acct_sites_all HCAS, OE_ORDER_LINES_ALL OLA, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, MTL_RESERVATIONS MR, mtl_parameters mpa, xxvia_vendedores j, fnd_user u, po_requisition_headers_ALL PO, MTL_SECONDARY_INVENTORIES INV"
   var_cadena = var_cadena + " Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID          =HL.LOCATION_ID AND HCSU.SITE_USE_ID         = OHA.INVOICE_TO_ORG_ID AND OLA.HEADER_ID            = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID   = HCAS.CUST_ACCT_SITE_ID AND MR.inventory_item_id     = c.inventory_item_id AND OHA.ship_from_org_id     = C.ORGANIZATION_ID AND mpa.organization_code    = 'CDI' AND OLA.LINE_ID              = MR.ORIG_DEMAND_SOURCE_LINE_ID AND oha.created_by           = u.user_id AND secondary_inventory_name = PO.ATTRIBUTE1 AND OHA.source_document_id   = PO.requisition_header_id AND oha.order_type_id        = 1002 AND c.organization_id        = mpa.organization_id AND mpa.organization_code    = 'CDI' AND HCAS.cust_acct_site_id   = 1100 AND HCAS.party_site_id       = 7021"
   var_cadena = var_cadena + " AND HCAS.cust_account_id     = 2040 AND oha.salesrep_id          = j.salesrep_id AND OHA.SOLD_TO_ORG_ID = 2040 AND OHA.SHIP_TO_ORG_ID = 1061 AND OHA.INVOICE_TO_ORG_ID = 1060 ORDER BY to_number(order_number)"
   rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            var_dia_str = CStr(Day(rs!Fecha))
            var_mes_str = CStr(Month(rs!Fecha))
            var_año_str = CStr(Year(rs!Fecha))
            If Len(var_dia_str) = 1 Then
               var_dia_str = "0" + var_dia_str
            End If
            If Len(var_mes_str) = 1 Then
               var_mes_str = "0" + var_mes_str
            End If
            If Len(var_año_str) = 2 Then
               var_año_str = "20" + var_año_str
            End If
            var_fecha = "{d '" + var_año_str + "-" + var_mes_str + "-" + var_dia_str + "'}"
            'MsgBox "insert into tb_temp_oracle_reservas (inte_tem_consecutivo, tipo_pedido, pedido, fecha, cliente, codigo, descripcion, cantidad) values (" + CStr(var_consecutivo) + ",'" + IIf(IsNull(rs!tipo_pedido), "", rs!tipo_pedido) + "','" + CStr(rs!pedido) + "'," + var_fecha + ",'" + IIf(IsNull(rs!Cliente), "", rs!Cliente) + "','" + IIf(IsNull(rs!codigo), "", rs!codigo) + "','" + IIf(IsNull(rs!descripcion), "", rs!descripcion) + "'," + CStr(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) + ")"
            rsaux.Open "insert into tb_temp_oracle_reservas (inte_tem_consecutivo, tipo_pedido, pedido, fecha, cliente, codigo, descripcion, cantidad, RUTA, USUARIO, NOMBRE_USUARIO, clave_almacen) values (" + CStr(var_consecutivo) + ",'" + IIf(IsNull(rs!tipo_pedido), "", rs!tipo_pedido) + "','" + CStr(rs!pedido) + "'," + var_fecha + ",'" + Replace(IIf(IsNull(rs!Cliente), "", rs!Cliente), "'", " ") + "','" + IIf(IsNull(rs!codigo), "", rs!codigo) + "','" + Replace(IIf(IsNull(rs!Descripcion), "", rs!Descripcion), "'", " ") + "'," + CStr(IIf(IsNull(rs!cantidad), 0, rs!cantidad)) + ",'" + IIf(IsNull(rs!ruta), "", rs!ruta) + "','" + IIf(IsNull(rs!USUARIO), "", rs!USUARIO) + "','" + IIf(IsNull(rs!nombre_usuario), "", rs!nombre_usuario) + "','" + IIf(IsNull(rs!ALMACEN), "", rs!ALMACEN) + "')", cnn, adOpenDynamic, adLockOptimistic
            rs.MoveNext
      Wend
   End If
   rs.Close
      
   
                     Set oexcel = CreateObject("Excel.Application")
                     Set owbook = oexcel.Workbooks.Add
                     Set osheet = owbook.Worksheets(1)
                     osheet.Name = "RESERVAS"
                     Screen.MousePointer = vbHourglass
                     iFila = 1
                     ifila2 = 1
                     icol2 = 1
                     iCol = 1
                     var_cadena = "select GETDATE() AS FECHA_REPORTE, CLAVE_ALMACEN, TIPO_PEDIDO, PEDIDO, FECHA, CLIENTE, CODIGO, DESCRIPCION,CANTIDAD, RUTA, USUARIO, NOMBRE_USUARIO from VW_ORACLE_RESERVAS WHERE INTE_TEM_CONSECUTIVO =  " + CStr(var_consecutivo) + " AND FECHA IS NOT NULL"
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
                     archivo = "c:\reportessid\reporte_reservas_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     owbook.SaveAs archivo
                     oexcel.Visible = True
                     Set oexcel = Nothing
                     Screen.MousePointer = vbDefault
   
                     rsaux10.Close
         
   'Set reporte = appl.OpenReport(App.Path + "\rep_oracle_reserva.rpt")
   'reporte.RecordSelectionFormula = "{VW_ORACLE_RESERVAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
   'For ntablas = 1 To reporte.Database.Tables.Count
   '    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   'Next ntablas
   'reporte.ExportOptions.FormatType = crEFTExcel80
   'reporte.ExportOptions.DestinationType = crEDTDiskFile
   'archivo = "c:\reportessid\reservas_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
   'reporte.ExportOptions.DiskFileName = archivo
   'reporte.Export False
   'Set reporte = Nothing
   'MsgBox "Se a terminado de guardar el archivo " + archivo
   var_si = MsgBox("¿Desea el reporte concentrado?", vbYesNo, "ATENCION")
   If var_si = 6 Then
                     Set oexcel = CreateObject("Excel.Application")
                     Set owbook = oexcel.Workbooks.Add
                     Set osheet = owbook.Worksheets(1)
                     osheet.Name = "RESERVAS CONCENTRADO"
                     Screen.MousePointer = vbHourglass
                     iFila = 1
                     ifila2 = 1
                     icol2 = 1
                     iCol = 1
                     var_cadena = "SELECT GETDATE() AS FECHA_REPORTE, CLAVE_ALMACEN, TIPO_PEDIDO, PEDIDO, FECHA, CLIENTE,RUTA,    EXPR1 AS CANTIDAD, USUARIO, NOMBRE_USUARIO, CLAVE_ALMACEN   FROM VW_ORACLE_RECERVAS_CONCENTRADO WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND FECHA IS NOT NULL"
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
                     archivo = "c:\reportessid\reporte_reservas_concentrado" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     owbook.SaveAs archivo
                     oexcel.Visible = True
                     Set oexcel = Nothing
                     Screen.MousePointer = vbDefault
   
                     rsaux10.Close
      
      'Set reporte = appl.OpenReport(App.Path + "\rep_oracle_recervas_concentrado.rpt")
      'reporte.RecordSelectionFormula = "{VW_ORACLE_RECERVAS_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
      'For ntablas = 1 To reporte.Database.Tables.Count
      '    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      'Next ntablas
      'reporte.ExportOptions.FormatType = crEFTExcel80
      'reporte.ExportOptions.DestinationType = crEDTDiskFile
      'archivo = "c:\reportessid\reservas_concentrado_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
      'reporte.ExportOptions.DiskFileName = archivo
      'reporte.Export False
      'Set reporte = Nothing
      'MsgBox "Se a terminado de guardar el archivo " + archivo
   End If
          
      
End Sub

Private Sub Form_Load()
   Top = 3200
   Left = 4200
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub
