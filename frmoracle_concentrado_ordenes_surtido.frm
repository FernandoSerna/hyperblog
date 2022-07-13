VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_concentrado_ordenes_surtido 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Concentrado de ordenes de surtido"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   6840
      Left            =   60
      TabIndex        =   3
      Top             =   405
      Width           =   11550
      Begin VB.CommandButton cmd_actualizar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   11100
         Picture         =   "frmoracle_concentrado_ordenes_surtido.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Actualizar pedidos"
         Top             =   135
         Width           =   330
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   9915
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   150
         Width           =   1140
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   8325
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   165
         Width           =   1140
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmoracle_concentrado_ordenes_surtido.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   45
         Picture         =   "frmoracle_concentrado_ordenes_surtido.frx":0318
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1035
         Picture         =   "frmoracle_concentrado_ordenes_surtido.frx":041A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   705
         Picture         =   "frmoracle_concentrado_ordenes_surtido.frx":04EC
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Marcar (Enter)"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1365
         Picture         =   "frmoracle_concentrado_ordenes_surtido.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   135
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_agentes 
         Height          =   6255
         Left            =   45
         TabIndex        =   9
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   11033
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Agente"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre Agente"
            Object.Width           =   5645
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Pedido"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cliente"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   9585
         TabIndex        =   13
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   7785
         TabIndex        =   12
         Top             =   225
         Width           =   420
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmoracle_concentrado_ordenes_surtido.frx":094C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11250
      Picture         =   "frmoracle_concentrado_ordenes_surtido.frx":0A4E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   0
      TabIndex        =   2
      Top             =   270
      Width           =   11625
   End
End
Attribute VB_Name = "frmoracle_concentrado_ordenes_surtido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_actualizar_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
            Me.lv_agentes.ListItems.Clear
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            txt_fecha_inicio = Me.txt_inicio
            txt_fecha_fin = Me.txt_fin
            var_cadena = "SELECT DISTINCT a.SOURCE_HEADER_NUMBER AS PEDIDO from WSH_DLVB_DLVY_V B, hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, XXVIA_VW_AGENTES D Where a.delivery_id = B.delivery_id AND A.delivery_detail_id = B.delivery_detail_id AND HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID and TRUNC(B.creation_date) >= to_date('" + CStr(txt_fecha_inicio) + "','DD-MM-YYYY') AND TRUNC(B.creation_date) < TO_DATE('" + CStr(CDate(txt_fecha_fin) + 1) + "','DD-MM-YYYY') AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND HCAS.CUST_ACCOUNT_ID = D.CUST_ACCOUNT_ID AND OHA.SHIP_FROM_ORG_ID = " + var_unidad_organizacional + " ORDER BY  a.SOURCE_HEADER_NUMBER"
            var_cadena_pedidos = ""
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  If var_cadena_pedidos = "" Then
                     var_cadena_pedidos = CStr(rs!PEDIDO)
                  Else
                     var_cadena_pedidos = var_cadena_pedidos + "," + CStr(rs!PEDIDO)
                  End If
                  rs.MoveNext
            Wend
            rs.Close
            If var_cadena_pedidos <> "" Then
               var_cadena = "SELECT distinct source_document_id, source_header_type_name, TRUNC(A.LAST_UPDATE_DATE) AS FECHA, source_header_number, D.COLLECTOR_ID, HL.ADDRESS1 AS CUSTOMER_NAME, D.NAME from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, xxvia_vw_agentes D Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCAS.CUST_ACCOUNT_ID = D.CUST_ACCOUNT_ID AND to_number(source_header_number) IN (" + var_cadena_pedidos + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y'"
   
               rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               numero_items_permisos = 0
               While Not rs.EOF
                     NOMBRE_CLIENTE = rs!customer_name
             
                     If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rs!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                        rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(IIf(IsNull(rs!source_document_id), "0", rs!source_document_id)) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           NOMBRE_CLIENTE = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                        End If
                        rsaux2.Close
                     End If
                     
                     Set list_item = lv_agentes.ListItems.Add(, , rs!collector_id)
                     list_item.SubItems(1) = IIf(IsNull(rs!Name), "", rs!Name)
                     list_item.SubItems(2) = rs!source_header_number
                     list_item.SubItems(3) = NOMBRE_CLIENTE
                     list_item.SubItems(4) = ""
                     list_item.SubItems(5) = rs!Fecha
                     rs.MoveNext:
                     numero_items_permisos = numero_items_permisos + 1
               Wend
               rs.Close
            End If
         Else
            MsgBox "La fecha inicial debe de ser menor a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha inicial incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
   Dim var_consecutivo As Integer
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_cadena_pedidos = ""
   For var_j = 1 To Me.lv_agentes.ListItems.Count
       Me.lv_agentes.ListItems.Item(var_j).Selected = True
       If Me.lv_agentes.selectedItem.SubItems(4) = "*" Then
          If var_cadena_pedidos = "" Then
             var_cadena_pedidos = Trim(Me.lv_agentes.selectedItem.SubItems(2))
          Else
             var_cadena_pedidos = var_cadena_pedidos + "," + Trim(Me.lv_agentes.selectedItem.SubItems(2))
          End If
       End If
   Next var_j
   If var_cadena_pedidos <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      
      var_cadena = "SELECT site_use_id, source_document_id, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,c.description as item_description, "
      var_cadena = var_cadena + " A.source_line_number,A.requested_quantity,A.released_status, c.segment1, a.source_header_type_name, oha.source_document_id, C.ATTRIBUTE2 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND to_number(source_header_number) IN (" + var_cadena_pedidos + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id "
      var_cadena = var_cadena + " AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y'"
      
      
      
      
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
         rsaux1.Open "DELETE FROM TB_TEMP_ORACLE_EXISTENCIA_ORDEN_SURTIDO WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
         While Not rs.EOF
               VAR_ESTABLECIMIENTO = rs!ship_to_org_id
               rsaux.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'SHIP_TO' AND csu.LOCATION = ps.party_site_number and csu.site_use_id = " + CStr(VAR_ESTABLECIMIENTO), cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux!vcha_esb_nombre), "", rsaux!vcha_esb_nombre)
               Else
                  VAR_NOMBRE_ESTABLECIMIENTO = ""
               End If
               rsaux.Close
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
               rsaux3.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux3.EOF Then
                  var_clave_agente = CStr(IIf(IsNull(rsaux3!collector_id), "", rsaux3!collector_id))
                  var_nombre_agente = CStr(IIf(IsNull(rsaux3!Name), "", rsaux3!Name))
               End If
               rsaux3.Close
               
               
               
               If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rs!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                  If rsaux2.State = 1 Then
                     rsaux2.Close
                  End If
                  rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(IIf(IsNull(rs!source_document_id), 0, rs!source_document_id)) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     VAR_ESTABLECIMIENTO = rsaux4!collector_id
                     VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                     rsaux4.Close
                  Else
                     rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     VAR_ESTABLECIMIENTO = rsaux4!collector_id
                     VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                     rsaux4.Close
                  End If
                  rsaux2.Close
               End If
               
               
               strconsulta = "select * from Xxvia_vw_existencias_inv where organization_id = ? and subinventory_code = ? and segment1 = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_unidad_organizacional))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rs!subinventory), "", rs!subinventory))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1))
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               
               If Not rsaux9.EOF Then
                  txt_fisico = IIf(IsNull(rsaux9!cantmano), 0, rsaux9!cantmano)
                  'Me.txt_apartado = IIf(IsNull(rsaux9!RESERVADA), 0, rsaux9!RESERVADA)
                  'Me.txt_disponible = IIf(IsNull(rsaux9!disponible), 0, rsaux9!disponible)
               Else
                  txt_fisico = 0
                  'Me.txt_apartado = 0
                  'Me.txt_disponible = 0
               End If
               rsaux9.Close
               rsaux9.Open "INSERT INTO TB_TEMP_ORACLE_EXISTENCIA_ORDEN_SURTIDO (INTE_TEM_CONSECUTIVO, SEGMENT1, EXISTENCIA) VALUES (" + CStr(var_consecutivo) + ",'" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'," + CStr(txt_fisico) + ")", cnn, adOpenDynamic, adLockOptimistic
               var_cadena = "insert into tb_temp_oracle_orden_surtido (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, DELIVERY_ID, DELIVERY_DETAIL_ID, ORGANIZATION_ID, SUBINVENTORY, DELIVERY_LINE_ID, INVENTORY_ITEM_ID, ITEM_DESCRIPTION, SOURCE_LINE_NUMBER, src_requested_quantity, RELEASED_STATUS, CUSTOMER_NAME, SEGMENT1, COLLECTOR_ID, NAME, date_requested, UBICACION, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, ORDENES)  values "
               var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rs!source_header_number), "", rs!source_header_number) + "', " + CStr(IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)) + ", " + CStr(IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)) + ", " + CStr(IIf(IsNull(rs!organization_id), 0, rs!organization_id)) + ", '" + IIf(IsNull(rs!subinventory), "", rs!subinventory) + "', " + CStr(IIf(IsNull(rs!delivery_line_id), 0, rs!delivery_line_id)) + ", " + CStr(IIf(IsNull(rs!inventory_item_id), "", rs!inventory_item_id)) + ", '" + IIf(IsNull(rs!item_description), "", rs!item_description) + "', '" + IIf(IsNull(rs!SOURCE_LINE_NUMBER), "", rs!SOURCE_LINE_NUMBER) + "', " + CStr(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)) + ", '" + IIf(IsNull(rs!released_status), "", rs!released_status) + "', '" + Replace(IIf(IsNull(rs!customer_name), "", rs!customer_name), "'", " ") + "', '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'"
               var_cadena = var_cadena + ", " + var_clave_agente + ",'" + var_nombre_agente + "'," + var_fecha + ",'" + IIf(IsNull(rs!attribute2), "", rs!attribute2) + "','" + CStr(VAR_ESTABLECIMIENTO) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "','" + var_cadena_pedidos + "')"
               'MsgBox var_cadena
               
               
               rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               rs.MoveNext
         Wend
         rs.Close
         rsaux.Open "select distinct source_header_number from tb_temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER is not null", cnn, adOpenDynamic, adLockOptimistic
         Set reporte = appl.OpenReport(App.Path + "\rep_oracle_concentrado_orden_surtido_resumen.rpt")
         var_cadena = "{VW_ORACLE_CONCENTRADO_ORDEN_SURTIDO_RESUMEN.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
         reporte.RecordSelectionFormula = var_cadena
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Concentrado de Ordenes de Surtido"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         
         var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_concentrado_orden_surtido_resumen.rpt")
            var_cadena = "{VW_ORACLE_CONCENTRADO_ORDEN_SURTIDO_RESUMEN.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            reporte.RecordSelectionFormula = var_cadena
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\concentrado_ordenes_surtido_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
         End If
         rsaux.Close
         rs.Open "delete from tb_temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         rs.Open "DELETE FROM TB_TEMP_ORACLE_EXISTENCIA_ORDEN_SURTIDO WHERE INTE_TEM_cONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      Else
         MsgBox "No existen pedidos", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se seleccionaron pedidos"
   End If
   
   
End Sub

Private Sub cmd_invertir_Click()
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      If lv_agentes.selectedItem.SubItems(4) = "*" Then
         lv_agentes.selectedItem.SubItems(4) = ""
         lv_agentes.ListItems.Item(i).Bold = False
         lv_agentes.ListItems.Item(i).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      Else
         lv_agentes.selectedItem.SubItems(4) = "*"
         lv_agentes.ListItems.Item(i).Bold = True
         lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   i = lv_agentes.selectedItem.Index
   If lv_agentes.selectedItem.SubItems(4) = "*" Then
      lv_agentes.selectedItem.SubItems(4) = ""
      lv_agentes.ListItems.Item(i).Bold = False
      lv_agentes.ListItems.Item(i).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_agentes.Refresh
   Else
      lv_agentes.selectedItem.SubItems(4) = "*"
      lv_agentes.ListItems.Item(i).Bold = True
      lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      lv_agentes.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      lv_agentes.selectedItem.SubItems(4) = ""
      lv_agentes.ListItems.Item(i).Bold = False
      lv_agentes.ListItems.Item(i).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
   Next i
   lv_agentes.Refresh
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_agentes.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_agentes.selectedItem.SubItems(4) = "" And var_rellena = True Then
         lv_agentes.selectedItem.SubItems(4) = "*"
         lv_agentes.ListItems.Item(i).Bold = True
         lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_agentes.selectedItem.SubItems(4) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_agentes.selectedItem.SubItems(4) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i

End Sub

Private Sub cmd_todos_Click()
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      lv_agentes.selectedItem.SubItems(4) = "*"
      lv_agentes.ListItems.Item(i).Bold = True
      lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
   Next i
   lv_agentes.Refresh
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
   Dim list_item As ListItem
   txt_fecha_fin = Date
   txt_fecha_inicio = Date
   Me.txt_fin = Date
   Me.txt_inicio = Date
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_cadena = "SELECT DISTINCT a.SOURCE_HEADER_NUMBER AS PEDIDO from WSH_DLVB_DLVY_V B, hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, XXVIA_VW_AGENTES D Where a.delivery_id = B.delivery_id AND A.delivery_detail_id = B.delivery_detail_id AND HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID and TRUNC(B.creation_date) >= to_date('" + CStr(txt_fecha_inicio) + "','DD-MM-YYYY') AND TRUNC(B.creation_date) < TO_DATE('" + CStr(CDate(txt_fecha_fin) + 1) + "','DD-MM-YYYY') AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND HCAS.CUST_ACCOUNT_ID = D.CUST_ACCOUNT_ID AND OHA.SHIP_FROM_ORG_ID = " + var_unidad_organizacional + " ORDER BY  a.SOURCE_HEADER_NUMBER"
   var_cadena_pedidos = ""
   rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         If var_cadena_pedidos = "" Then
            var_cadena_pedidos = CStr(rs!PEDIDO)
         Else
            var_cadena_pedidos = var_cadena_pedidos + "," + CStr(rs!PEDIDO)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_cadena_pedidos <> "" Then
      var_cadena = "SELECT distinct source_document_id, source_header_type_name,TRUNC(A.LAST_UPDATE_DATE) AS FECHA, source_header_number, D.COLLECTOR_ID, HL.ADDRESS1 AS CUSTOMER_NAME, D.NAME from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, xxvia_vw_agentes D Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCAS.CUST_ACCOUNT_ID = D.CUST_ACCOUNT_ID AND to_number(source_header_number) IN (" + var_cadena_pedidos + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y'"
   
      rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
      numero_items_permisos = 0
      While Not rs.EOF
      
      
            NOMBRE_CLIENTE = rs!customer_name
            
            If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rs!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
               rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(IIf(IsNull(rs!source_document_id), "0", rs!source_document_id)) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  NOMBRE_CLIENTE = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
               End If
               rsaux2.Close
            End If
            
            
            Set list_item = lv_agentes.ListItems.Add(, , rs!collector_id)
            list_item.SubItems(1) = IIf(IsNull(rs!Name), "", rs!Name)
            list_item.SubItems(2) = rs!source_header_number
            list_item.SubItems(3) = NOMBRE_CLIENTE
            list_item.SubItems(4) = ""
            list_item.SubItems(5) = rs!Fecha
            rs.MoveNext:
            numero_items_permisos = numero_items_permisos + 1
      Wend
      rs.Close
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_agentes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_agentes, ColumnHeader)
End Sub

Private Sub lv_agentes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_agentes.ListItems.Count > 0 Then
         i = lv_agentes.selectedItem.Index
         If lv_agentes.selectedItem.SubItems(4) = "*" Then
            lv_agentes.selectedItem.SubItems(4) = ""
            lv_agentes.ListItems.Item(i).Bold = False
            lv_agentes.ListItems.Item(i).ForeColor = &H80000012
            lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
            lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
            lv_agentes.ListItems.Item(i).ListSubItems(3).Bold = False
            lv_agentes.ListItems.Item(i).ListSubItems(4).Bold = False
            lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
            lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
            lv_agentes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
            lv_agentes.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
            lv_agentes.Refresh
         Else
            lv_agentes.selectedItem.SubItems(4) = "*"
            lv_agentes.ListItems.Item(i).Bold = True
            lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
            lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_agentes.ListItems.Item(i).ListSubItems(3).Bold = True
            lv_agentes.ListItems.Item(i).ListSubItems(4).Bold = True
            lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
            lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
            lv_agentes.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
            lv_agentes.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
            lv_agentes.Refresh
         End If
      End If
   End If
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes = CDate(Me.txt_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub

