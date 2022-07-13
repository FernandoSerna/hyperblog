VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_libera_pedidos_vxt 
   Caption         =   "Liberación de ordenes de surtido"
   ClientHeight    =   10290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   15270
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_actualizar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   15
      Picture         =   "frmoracle_libera_pedidos_vxt.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Actualizar pedidos"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   14850
      Picture         =   "frmoracle_libera_pedidos_vxt.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   345
      Picture         =   "frmoracle_libera_pedidos_vxt.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   270
      Width           =   15195
   End
   Begin MSComctlLib.ListView lv_pedidos 
      Height          =   10635
      Left            =   45
      TabIndex        =   0
      Top             =   420
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   18759
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "   Agente"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cliente"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Referencia"
         Object.Width           =   2911
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Pedido "
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Fecha"
         Object.Width           =   2064
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Cant. Pedido"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Importe Pedido"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Cant. O.S."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ESTATUS"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "tipo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Credito"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Importe O.S."
         Object.Width           =   2822
      EndProperty
   End
End
Attribute VB_Name = "frmoracle_libera_pedidos_vxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter


Private Sub cmd_actualizar_Click()
   On Error GoTo SALIR:
   rsaux.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   
   'var_cadena = "SELECT HCAS.CUST_ACCOUNT_ID, oha.order_type_id, oha.header_id, oha.ordered_date, oha.order_number,  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, f.orig_system_reference from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU, hz_cust_acct_sites_all f Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID and oha.order_type_id in (1106,1161,1049, 1556) and HCSU.site_use_code = 'BILL_TO' and f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID  and oha.order_number > 25000"
   'var_cadena = var_cadena + " ORDER BY oha.order_number"
   'var_cadena = "SELECT HCAS.CUST_ACCOUNT_ID, oha.order_type_id, oha.header_id, oha.ordered_date, oha.order_number,  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, f.orig_system_reference from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU, hz_cust_acct_sites_all f Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID and oha.order_type_id in (1106,1161,1049, 1556) and HCSU.site_use_code = 'BILL_TO' and f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID  and oha.order_number IN (select SOURCE_HEADER_NUMBER from WSH_DELIVERABLES_V where RELEASED_STATUS = 'Y') ORDER BY oha.order_number"
   var_cadena = "SELECT HCAS.CUST_ACCOUNT_ID, oha.order_type_id, oha.header_id, oha.ordered_date, oha.order_number,  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, f.orig_system_reference, DUE_DAYS from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU, hz_cust_acct_sites_all f,  RA_TERMS_TL A, RA_TERMS_LINES B Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID and oha.order_type_id in (1106,1161,1049, 1556, 1421) and HCSU.site_use_code = 'BILL_TO' and f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID  and oha.order_number IN (select SOURCE_HEADER_NUMBER from WSH_DELIVERABLES_V where RELEASED_STATUS = 'Y') AND OHA.PAYMENT_TERM_ID =  A.TERM_ID AND A.TERM_ID = B.TERM_ID AND A.LANGUAGE = 'ESA' ORDER BY oha.order_number"
   
   
   Me.lv_pedidos.ListItems.Clear
   rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_cadena_headers = ""
      While Not rs.EOF
            If var_cadena_headers = "" Then
               var_cadena_headers = CStr(rs!header_id)
            Else
               var_cadena_headers = var_cadena_headers + "," + CStr(rs!header_id)
            End If
            rs.MoveNext
      Wend
      rs.MoveFirst
      
      
      rsaux.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux.Open "mo_global.set_policy_context('S', 92)", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux.Open "fnd_global.apps_initialize(fnd_profile.value(1130),fnd_profile.value(50755),fnd_profile.value(660))", cnnoracle_4, adOpenDynamic, adLockOptimistic
      'MsgBox Err.Number
      var_i = 0
      rsaux.Open "SELECT header_id, unit_selling_price, tax_value, line_id, pricing_quantity  FROM oe_order_lines_v WHERE  HEADER_ID IN (" + var_cadena_headers + ") ", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rsaux.EOF
            strconsulta = "select * from xxvia_tb_precios_pedidos where header_id = ? and line_id = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux!header_id)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux!line_id)
                 .Parameters.Append parametro
            End With
            Set rsaux9 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If rsaux9.EOF Then
               strconsulta = "insert into xxvia_tb_precios_pedidos (header_id, unit_selling_price, tax_value, line_id, pricing_quantity) values (?, ?, ? ,?, ?)"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux!header_id)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux!unit_selling_price)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux!tax_value)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux!line_id)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux!PRICING_QUANTITY)
                    .Parameters.Append parametro
               End With
               Set rsaux10 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
            Else
               strconsulta = "update xxvia_tb_precios_pedidos  set unit_selling_price = ?, tax_value = ?, pricing_quantity = ? where   header_id = ? and line_id = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux!unit_selling_price)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux!tax_value)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux!PRICING_QUANTITY)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux!header_id)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux!line_id)
                    .Parameters.Append parametro
               End With
               Set rsaux10 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
            End If
            rsaux.MoveNext
      Wend
      rsaux.Close
      
      While Not rs.EOF
            var_i = var_i + 1
            rsaux.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rsaux4.Open "select released_status from WSH_DELIVERABLES_V where source_header_number = '" + CStr(rs!order_number) + "' AND released_status = 'Y'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If rsaux4.EOF Then
               VAR_ESTATUS = ""
            Else
               VAR_ESTATUS = IIf(IsNull(rsaux4!released_status), "", rsaux4!released_status)
            End If
            rsaux4.Close
            If VAR_ESTATUS = "Y" Then
               rsaux.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
               Set list_item = lv_pedidos.ListItems.Add(, , rsaux!Name)
               rsaux.Close
               list_item.SubItems(1) = IIf(IsNull(rs!customer_name), "", rs!customer_name)
               list_item.SubItems(2) = Trim(IIf(IsNull(rs!orig_system_reference), 0, rs!orig_system_reference))
               list_item.SubItems(3) = IIf(IsNull(rs!order_number), "", rs!order_number)
               list_item.SubItems(4) = Format(IIf(IsNull(rs!ORDERED_DATE), "", rs!ORDERED_DATE), "Short date")
               list_item.SubItems(9) = IIf(IsNull(rs!order_type_id), "", rs!order_type_id)
               'rsaux.Open "select sum(pricing_quantity) as cantidad, sum((pricing_quantity * unit_selling_price)  + (pricing_quantity * tax_value)) as precio from oe_order_lines_v where header_id = " + CStr(rs!header_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
               rsaux.Open "select sum(pricing_quantity) as cantidad, sum((pricing_quantity * unit_selling_price)  + (pricing_quantity * tax_value)) as precio from xxvia_Tb_precios_pedidos where header_id = " + CStr(rs!header_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  list_item.SubItems(5) = Format(IIf(IsNull(rsaux!Cantidad), 0, rsaux!Cantidad), "###,###,##0.00")
                  list_item.SubItems(6) = Format(IIf(IsNull(rsaux!Precio), 0, rsaux!Precio), "###,###,##0.00")
               Else
                  list_item.SubItems(5) = Format(0, "###,###,##0.00")
                  list_item.SubItems(6) = Format(0, "###,###,##0.00")
               End If
               rsaux.Close
               
               '22-11-13 rsaux.Open "select sum(b.requested_quantity) as cantidad,  sum((b.requested_quantity * unit_selling_price)  + (pricing_quantity * tax_value)) as importe from oe_order_lines_v a, WSH_dELIVERABLES_V b where a. header_id = " + CStr(rs!header_id) + " and b.source_header_id = a.header_id and b.source_line_id = a.line_id and b.released_status = 'Y'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               '01-11-13 rsaux.Open "SELECT sum(requested_quantity) as cantidad, sum(requested_quantity * unit_price) as importe FROM WSH_dELIVERABLES_V WHERE SOURCE_HEADER_NUMBER = " + CStr(rs!order_number) + " and released_status = 'Y'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               rsaux.Open "select sum(b.requested_quantity) as cantidad,  sum((b.requested_quantity * unit_selling_price)  + (pricing_quantity * tax_value)) as importe from xxvia_tb_precios_pedidos a, WSH_dELIVERABLES_V b where a. header_id = " + CStr(rs!header_id) + " and b.source_header_id = a.header_id and b.source_line_id = a.line_id and b.released_status = 'Y'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  list_item.SubItems(7) = Format(IIf(IsNull(rsaux!Cantidad), 0, rsaux!Cantidad), "###,###,##0.00")
                  list_item.SubItems(11) = Format(IIf(IsNull(rsaux!Importe), 0, rsaux!Importe), "###,###,##0.00")
               Else
                  list_item.SubItems(7) = Format(0, "###,###,##0.00")
                  list_item.SubItems(11) = Format(0, "###,###,##0.00")
               End If
               rsaux.Close
   
               rsaux.Open "select * from OE_ORDER_HOLDS_ALL where header_id = " + CStr(rs!header_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  VAR_ESTATUS = IIf(IsNull(rsaux!released_flag), "N", rsaux!released_flag)
               Else
                  VAR_ESTATUS = "Y"
               End If
               rsaux.Close
               list_item.SubItems(8) = VAR_ESTATUS
               list_item.SubItems(10) = IIf(IsNull(rs!due_Days), 0, rs!due_Days)
            End If
            rs.MoveNext
      Wend
      'MsgBox var_i
      For var_i = 1 To Me.lv_pedidos.ListItems.Count
          Me.lv_pedidos.ListItems.Item(var_i).Selected = True
          If Trim(lv_pedidos.selectedItem.SubItems(8)) = "Y" Then
             If Trim(Me.lv_pedidos.selectedItem.SubItems(9)) = "1106" Or Trim(Me.lv_pedidos.selectedItem.SubItems(9)) = "1049" Or Trim(Me.lv_pedidos.selectedItem.SubItems(9)) = "1421" Then
                lv_pedidos.ListItems.Item(var_i).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(1).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(2).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(3).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(4).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(5).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(6).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(7).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(11).Bold = False
                lv_pedidos.ListItems.Item(var_i).ForeColor = &HC000&
                lv_pedidos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HC000&
                lv_pedidos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HC000&
                lv_pedidos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HC000&
                lv_pedidos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HC000&
                lv_pedidos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HC000&
                lv_pedidos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HC000&
                lv_pedidos.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HC000&
                lv_pedidos.ListItems.Item(var_i).ListSubItems(11).ForeColor = &HC000&
             End If
             If Trim(Me.lv_pedidos.selectedItem.SubItems(9)) = "1161" Or Trim(Me.lv_pedidos.selectedItem.SubItems(9)) = "1556" Then
                lv_pedidos.ListItems.Item(var_i).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(1).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(2).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(3).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(4).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(5).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(6).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(7).Bold = False
                lv_pedidos.ListItems.Item(var_i).ForeColor = &HFF0000
                lv_pedidos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF0000
                lv_pedidos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF0000
                lv_pedidos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF0000
                lv_pedidos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF0000
                lv_pedidos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF0000
                lv_pedidos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF0000
                lv_pedidos.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HFF0000
                lv_pedidos.ListItems.Item(var_i).ListSubItems(11).ForeColor = &HFF0000
             End If
             If CInt(lv_pedidos.selectedItem.SubItems(10)) > 0 Then
                lv_pedidos.ListItems.Item(var_i).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(1).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(2).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(3).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(4).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(5).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(6).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(7).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(11).Bold = False
                lv_pedidos.ListItems.Item(var_i).ForeColor = &HFF0000
                lv_pedidos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF0000
                lv_pedidos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF0000
                lv_pedidos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF0000
                lv_pedidos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF0000
                lv_pedidos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF0000
                lv_pedidos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF0000
                lv_pedidos.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HFF0000
                lv_pedidos.ListItems.Item(var_i).ListSubItems(11).ForeColor = &HFF0000
             End If
             
             rsaux4.Open "select * from tb_oracle_pedidos_vxt_impresos where pedido = " + Me.lv_pedidos.selectedItem.SubItems(3), cnn, adOpenDynamic, adLockOptimistic
             If Not rsaux4.EOF Then
                lv_pedidos.ListItems.Item(var_i).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(1).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(2).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(3).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(4).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(5).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(6).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(7).Bold = False
                lv_pedidos.ListItems.Item(var_i).ListSubItems(11).Bold = False
                lv_pedidos.ListItems.Item(var_i).ForeColor = &HFF&
                lv_pedidos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF&
                lv_pedidos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF&
                lv_pedidos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF&
                lv_pedidos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF&
                lv_pedidos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF&
                lv_pedidos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF&
                lv_pedidos.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HFF&
                lv_pedidos.ListItems.Item(var_i).ListSubItems(11).ForeColor = &HFF&
             End If
             rsaux4.Close
          
          Else
             lv_pedidos.ListItems.Item(var_i).Bold = False
             lv_pedidos.ListItems.Item(var_i).ListSubItems(1).Bold = False
             lv_pedidos.ListItems.Item(var_i).ListSubItems(2).Bold = False
             lv_pedidos.ListItems.Item(var_i).ListSubItems(3).Bold = False
             lv_pedidos.ListItems.Item(var_i).ListSubItems(4).Bold = False
             lv_pedidos.ListItems.Item(var_i).ListSubItems(5).Bold = False
             lv_pedidos.ListItems.Item(var_i).ListSubItems(6).Bold = False
             lv_pedidos.ListItems.Item(var_i).ListSubItems(7).Bold = False
             lv_pedidos.ListItems.Item(var_i).ListSubItems(11).Bold = False
             lv_pedidos.ListItems.Item(var_i).ForeColor = &H80000008
             lv_pedidos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000008
             lv_pedidos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000008
             lv_pedidos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000008
             lv_pedidos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000008
             lv_pedidos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H80000008
             lv_pedidos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H80000008
             lv_pedidos.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H80000008
             lv_pedidos.ListItems.Item(var_i).ListSubItems(11).ForeColor = &H80000008
          End If
      Next var_i
   Else
      MsgBox "No existen pedidos", vbOKOnly, "ATENCION"
   End If
   rs.Close
   Exit Sub
SALIR:
   'MsgBox Err.Description
   If Err.Number = -2147217900 Then
      Resume
   End If
   If rs.State = 1 Then
      rs.Close
   End If
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_pedidos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_pedidos, ColumnHeader)
End Sub

Private Sub lv_pedidos_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = 116 Then
      If Me.lv_pedidos.ListItems.Count > 0 Then
         rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Trim(Me.lv_pedidos.selectedItem.SubItems(8)) = "Y" Then
            
            var_cadena = "SELECT HCAS.CUST_ACCOUNT_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, c.attribute2, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) "
            var_cadena = var_cadena + " Between " + Me.lv_pedidos.selectedItem.SubItems(3) + " And " + Me.lv_pedidos.selectedItem.SubItems(3) + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID  AND released_status = 'Y'"
            rs.Open "select shipping_method_code, packing_instructions from oe_order_headers_all where order_number = " + Me.lv_pedidos.selectedItem.SubItems(3), cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_paqueteria = ""
            If Not rs.EOF Then
               VAR_COMENTARIOS = IIf(IsNull(rs!packing_instructions), "", rs!packing_instructions)
               var_tipo_metodo = IIf(IsNull(rs(0).Value), "", rs(0).Value)
               If var_tipo_metodo <> "" Then
                  rsaux1.Open "SELECT description FROM fnd_lookup_values where lookup_type = 'SHIP_METHOD' and lookup_code = '" + var_tipo_metodo + "' AND LANGUAGE = 'ESA'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                     var_paqueteria = IIf(IsNull(rsaux1(0).Value), "", rsaux1(0).Value)
                  End If
                  rsaux1.Close
               End If
            End If
            rs.Close
            
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
                     var_dia = CStr(Day(CDate(rs!DATE_REQUESTED)))
                     var_mes = CStr(Month(CDate(rs!DATE_REQUESTED)))
                     var_año = CStr(Year(CDate(rs!DATE_REQUESTED)))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     rsaux.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     
                     
                     rsaux6.Open "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.invoice_TO_ORG_ID AND oha.order_number = '" + IIf(IsNull(rs!source_header_number), "", rs!source_header_number) + "' AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux6.EOF Then
                        rsaux5.Open "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = '" + IIf(IsNull(rs!source_header_number), "", rs!source_header_number) + "' AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux5.EOF Then
                           var_nombre = IIf(IsNull(rsaux5!customer_name), "", rsaux5!customer_name)
                           var_tel = IIf(IsNull(rsaux5!tel), 0, rsaux5!tel)
                           VAR_DIRECCION = IIf(IsNull(rsaux5!calle), "", rsaux5!calle) + " " + IIf(IsNull(rsaux5!NUMERO), "", rsaux5!NUMERO)
                           VAR_COLONIA = IIf(IsNull(rsaux5!colonia), "", rsaux5!colonia)
                           var_ciudad = IIf(IsNull(rsaux5!ciudad), "", rsaux5!ciudad)
                           VAR_MUNICIPIO = IIf(IsNull(rsaux5!municipio), "", rsaux5!municipio)
                           var_estado = IIf(IsNull(rsaux5!estado), "", rsaux5!estado)
                           var_pais = IIf(IsNull(rsaux5!pais), "", rsaux5!pais)
                           VAR_CP = IIf(IsNull(rsaux5!cp), "", rsaux5!cp)
                           rsaux5.Close
                        Else
                           rsaux5.Close
                           var_nombre = IIf(IsNull(rsaux6!customer_name), "", rsaux6!customer_name)
                           var_tel = IIf(IsNull(rsaux6!tel), 0, rsaux6!tel)
                           VAR_DIRECCION = IIf(IsNull(rsaux6!calle), "", rsaux6!calle) + " " + IIf(IsNull(rsaux6!NUMERO), "", rsaux6!NUMERO)
                           VAR_COLONIA = IIf(IsNull(rsaux6!colonia), "", rsaux6!colonia)
                           var_ciudad = IIf(IsNull(rsaux6!ciudad), "", rsaux6!ciudad)
                           VAR_MUNICIPIO = IIf(IsNull(rsaux6!municipio), "", rsaux6!municipio)
                           var_estado = IIf(IsNull(rsaux6!estado), "", rsaux6!estado)
                           var_pais = IIf(IsNull(rsaux6!pais), "", rsaux6!pais)
                           VAR_CP = IIf(IsNull(rsaux6!cp), "", rsaux6!cp)
                        End If
                     Else
                        var_tel = 0
                        VAR_DIRECCION = ""
                        VAR_COLONIA = ""
                        var_ciudad = ""
                        VAR_MUNICIPIO = ""
                        var_estado = ""
                        var_pais = ""
                        VAR_CP = ""
                     End If
                     rsaux6.Close
                     If var_tel > 0 Then
                        rsaux6.Open "select Phone_Number from hz_contact_points where owner_table_id = " + CStr(var_tel), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux6.EOF Then
                           var_telefono = CStr(IIf(IsNull(rsaux6(0).Value), "", rsaux6(0).Value))
                        Else
                           var_telefono = ""
                        End If
                        rsaux6.Close
                     Else
                        var_telefono = ""
                     End If
                     
                     
                     
                     
                     var_cadena = "insert into tb_temp_oracle_orden_surtido (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, DELIVERY_ID, DELIVERY_DETAIL_ID, ORGANIZATION_ID, SUBINVENTORY, DELIVERY_LINE_ID, INVENTORY_ITEM_ID, ITEM_DESCRIPTION, SOURCE_LINE_NUMBER, src_requested_quantity, RELEASED_STATUS, CUSTOMER_NAME, SEGMENT1, COLLECTOR_ID, NAME, date_requested, ubicacion, pais, municipio, estado, ciudad, colonia, direccion, cp, PAQUETERIA, telefono, nombre_establecimiento, COMENTARIO)  values "
                     var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rs!source_header_number), "", rs!source_header_number) + "', " + CStr(IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)) + ", " + CStr(IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)) + ", " + CStr(IIf(IsNull(rs!organization_id), 0, rs!organization_id)) + ", '" + IIf(IsNull(rs!subinventory), "", rs!subinventory) + "', " + CStr(IIf(IsNull(rs!delivery_line_id), 0, rs!delivery_line_id)) + ", " + CStr(IIf(IsNull(rs!inventory_item_id), "", rs!inventory_item_id)) + ", '" + IIf(IsNull(rs!Description), "", rs!Description) + "', '" + IIf(IsNull(rs!SOURCE_LINE_NUMBER), "", rs!SOURCE_LINE_NUMBER) + "', " + CStr(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)) + ", '" + IIf(IsNull(rs!released_status), "", rs!released_status) + "', '" + IIf(IsNull(rs!customer_name), "", rs!customer_name) + "', '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'"
                     var_cadena = var_cadena + ", " + CStr(IIf(IsNull(rsaux!collector_id), 0, rsaux!collector_id)) + ",'" + IIf(IsNull(rsaux!Name), "", rsaux!Name) + "'," + var_fecha + ",'" + IIf(IsNull(rs!attribute2), "", rs!attribute2) + "','" + var_pais + "','" + VAR_MUNICIPIO + "','" + var_estado + "', '" + var_ciudad + "','" + VAR_COLONIA + "','" + VAR_DIRECCION + "','" + VAR_CP + "','" + var_paqueteria + "','" + var_telefono + "','" + var_nombre + "','" + VAR_COMENTARIOS + "')"
                     rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rsaux.Close
                     rs.MoveNext
               Wend
               rsaux.Open "select distinct source_header_number from tb_temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER is not null", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  While Not rsaux.EOF
                        x = 1
                        If x = 0 Then
                           Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_ft.rpt")
                           reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.SOURCE_HEADER_NUMBER} = '" + rsaux(0).Value + "' and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                           frmvistasprevias.cr.ReportSource = reporte
                           For ntablas = 1 To reporte.Database.Tables.Count
                               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                           Next ntablas
                           frmvistasprevias.cr.ViewReport
                           frmvistasprevias.Caption = "Pedido"
                           frmvistasprevias.Show 1
                           Set reporte = Nothing
                        Else
                           Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_ft.rpt")
                           reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.SOURCE_HEADER_NUMBER} = '" + rsaux(0).Value + "' and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                           'frmvistasprevias.cr.ReportSource = reporte
                           For ntablas = 1 To reporte.Database.Tables.Count
                               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                           Next ntablas
                           'frmvistasprevias.cr.ViewReport
                           'frmvistasprevias.Caption = "ORDENES DE SURTIDO"
                           reporte.PrintOut False
                           Set reporte = Nothing
                        End If
                        rsaux.MoveNext
                  Wend
                  rsaux1.Open "SELECT * FROM TB_ORACLE_PEDIDOS_VXT_IMPRESOS WHERE PEDIDO = " + Me.lv_pedidos.selectedItem.SubItems(3)
                  If rsaux1.EOF Then
                     rsaux2.Open "INSERT INTO TB_ORACLE_PEDIDOS_VXT_IMPRESOS (PEDIDO, MAQUINA, USUARIO, FECHA)  VALUES (" + Me.lv_pedidos.selectedItem.SubItems(3) + ",'" + fun_NombrePc + "','" + var_clave_usuario_global + "', GETDATE())", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux1.Close
                  var_i = Me.lv_pedidos.selectedItem.Index
                  lv_pedidos.ListItems.Item(var_i).Bold = False
                  lv_pedidos.ListItems.Item(var_i).ListSubItems(1).Bold = False
                  lv_pedidos.ListItems.Item(var_i).ListSubItems(2).Bold = False
                  lv_pedidos.ListItems.Item(var_i).ListSubItems(3).Bold = False
                  lv_pedidos.ListItems.Item(var_i).ListSubItems(4).Bold = False
                  lv_pedidos.ListItems.Item(var_i).ListSubItems(5).Bold = False
                  lv_pedidos.ListItems.Item(var_i).ListSubItems(6).Bold = False
                  lv_pedidos.ListItems.Item(var_i).ListSubItems(7).Bold = False
                  lv_pedidos.ListItems.Item(var_i).ForeColor = &HFF&
                  lv_pedidos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF&
                  lv_pedidos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF&
                  lv_pedidos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF&
                  lv_pedidos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF&
                  lv_pedidos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF&
                  lv_pedidos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF&
                  lv_pedidos.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HFF&
               
               End If
               rsaux.Close
            Else
               MsgBox "El pedido no a sido despachado", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "El pedido no a sido liberado", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub
