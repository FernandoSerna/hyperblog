VERSION 5.00
Begin VB.Form frmreporte_logistica_negado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Logística de Negado"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   90
      TabIndex        =   2
      Top             =   435
      Width           =   4245
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   315
         Width           =   1140
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   375
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2325
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3975
      Picture         =   "frmreporte_logistica_negado.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmreporte_logistica_negado.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   90
      TabIndex        =   7
      Top             =   270
      Width           =   4230
   End
End
Attribute VB_Name = "frmreporte_logistica_negado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.CommandTimeout = 720
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_ORACLE_REPORTE_NEGADO", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_ORACLE_REPORTE_NEGADO (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            
            var_fecha_inicio = var_dia + "-" + var_mes + "-" + var_año
            VAR_FECHA_INICIO_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = var_dia + "-" + var_mes + "-" + var_año
            
            
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            VAR_FECHA_FIN_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
              
            'var_cadena = "SELECT DISTINCT SOURCE_HEADER_NUMBER AS PEDIDO from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id and a.creation_date >= to_date('" + var_fecha_inicio + "','DD-MM-YYYY') AND A.creation_date < TO_DATE('" + var_fecha_fin + "','DD-MM-YYYY')"
            'var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND OHA.SHIP_FROM_ORG_ID = " + var_unidad_organizacional
            var_cadena = "SELECT DISTINCT SOURCE_HEADER_NUMBER from WSH_DLVB_DLVY_V B, hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where a.delivery_id = B.delivery_id AND A.delivery_detail_id = B.delivery_detail_id AND HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID and TRUNC(B.creation_date) >= to_date('" + var_fecha_inicio + "','DD-MM-YYYY') AND TRUNC(B.creation_date) < TO_DATE('" + var_fecha_fin + "','DD-MM-YYYY') AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND OHA.SHIP_FROM_ORG_ID = " + var_unidad_organizacional + "ORDER BY SOURCE_HEADER_NUMBER"

            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            VAR_CADENA_PEDIDOS = ""
            While Not rs.EOF
                  If VAR_CADENA_PEDIDOS = "" Then
                     VAR_CADENA_PEDIDOS = CStr(rs(0).Value)
                  Else
                     VAR_CADENA_PEDIDOS = VAR_CADENA_PEDIDOS + ", " + CStr(rs(0).Value)
                  End If
                  rs.MoveNext
            Wend
            rs.Close
            If VAR_CADENA_PEDIDOS <> "" Then
               'var_cadena = "SELECT oh.ordered_date, HCSU.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  e.collector_id, E.NAME, oh.header_id, oh.order_number, oh.transactional_curr_code, NVL(ol.ordered_quantity,0) AS CANTIDAD_PEDIDA, NVL(ol.cancelled_quantity,0) AS CANTIDAD_NEGADA, NVL(ol.shipped_quantity,0) AS CANTIDAD_surtida, ol.line_id, ol.ordered_item, g.description, ol.order_quantity_uom, ol.inventory_item_id, ol.price_list_id, ol.unit_selling_price, ol.unit_list_price, DECODE(ol.cancelled_flag,'Y','CANCELADA','SURTIDA') line_status, ol.flow_status_code, NVL(OER.REASON_CODE,'') AS REASON_CODE FROM oe_order_headers_all oh, oe_order_lines_all ol, OE_ORDER_LINES_HISTORY OLH, OE_REASONS OER, hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors E, hz_cust_acct_sites_all f, xxvia_system_items_b g  WHERE order_number IN (" + var_cadena_pedidos + ") "
               'var_cadena = var_cadena + " AND oh.header_id = ol.header_id AND ol.ship_from_org_id = " + var_unidad_organizacional + " AND oL.header_id = oLh.header_id(+) AND OL.LINE_ID = OLH.LINE_ID(+) AND OER.HEADER_ID(+) = oL.header_id AND OER.ENTITY_ID(+) = OL.LINE_ID and HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID = HL.LOCATION_ID AND HCSU.SITE_USE_ID = OH.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id AND f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID and ol.inventory_item_id = g.inventory_item_id and g.organization_id = ol.ship_from_org_id "
               var_cadena = "SELECT oh.ordered_date, HCSU.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  e.collector_id, E.NAME, oh.header_id, oh.order_number, oh.transactional_curr_code, NVL(ol.ordered_quantity,0) AS CANTIDAD_PEDIDA, NVL(ol.cancelled_quantity,0) AS CANTIDAD_NEGADA, NVL(ol.shipped_quantity,0) AS CANTIDAD_surtida, ol.line_id, ol.ordered_item, g.description, ol.order_quantity_uom, ol.inventory_item_id, ol.price_list_id, ol.unit_selling_price, ol.unit_list_price, DECODE(ol.cancelled_flag,'Y','CANCELADA','SURTIDA') line_status, ol.flow_status_code, NVL(OER.REASON_CODE,'') AS REASON_CODE, LINEA FROM oe_order_headers_all oh, oe_order_lines_all ol, OE_ORDER_LINES_HISTORY OLH, OE_REASONS OER, hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors E, hz_cust_acct_sites_all f, xxvia_system_items_b g, xxvia_vw_articulos_cat H WHERE order_number IN (" + VAR_CADENA_PEDIDOS + ") "
               var_cadena = var_cadena + " AND oh.header_id = ol.header_id AND ol.ship_from_org_id = " + var_unidad_organizacional + " AND oL.header_id = oLh.header_id(+) AND OL.LINE_ID = OLH.LINE_ID(+) AND OER.HEADER_ID(+) = oL.header_id AND OER.ENTITY_ID(+) = OL.LINE_ID and HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID = HL.LOCATION_ID AND HCSU.SITE_USE_ID = OH.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id AND f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID and ol.inventory_item_id = g.inventory_item_id and g.organization_id = ol.ship_from_org_id AND G.INVENTORY_ITEM_ID = H.ITEM_ID(+) AND G.ORGANIZATION_ID = H.ORGANIZATION_ID(+) "
               rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  While Not rs.EOF
                        var_es_Catalogo = 0
                        If UCase(IIf(IsNull(rs!linea), "", rs!linea)) = "CATALOGOS" Then
                           var_es_Catalogo = 1
                        End If
                        var_cadena = "insert INTO TB_TEMP_ORACLE_REPORTE_NEGADO (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, FECHA_PEDIDO, FECHA_CERRADO, ESTATUS, CANAL_VENTA, NOMBRE_CANAL_VENTA,       CLAVE_AGENTE, NOMBRE_AGENTE,                 CLAVE_CLIENTE,                     NOMBRE_CLIENTE,         PEDIDO,                   CODIGO,                   DESCRIPCION, CATALOGO, ES_CATALOGO, CANTIDAD_PEDIDA,                 PRECIO,       cantidad_surtida,                     CANTIDAD_NEGADA,                CAUSA_NEGADO, CANTIDAD_NEGADA_PRODUCCION) VALUES"
                        If IIf(IsNull(rs!cantidad_surtida), 0, rs!cantidad_surtida) > 0 Then
                           If IIf(IsNull(rs!REASON_CODE), "", rs!REASON_CODE) = "PRODUCCION" Then
                              var_cadena = var_cadena + "(" + CStr(var_consecutivo) + "," + VAR_FECHA_INICIO_TABLA + "," + VAR_FECHA_FIN_TABLA + ",'" + CStr(rs!ORDERED_DATE) + "','','" + rs!LINE_STATUS + "','','','" + CStr(rs!collector_id) + "','" + rs!Name + "','" + CStr(rs!CUST_ACCT_SITE_ID) + "','" + Mid(rs!CUSTOMER_NAME, 1, 50) + "'," + CStr(rs!ORDER_NUMBER) + ",'" + rs!ordered_item + "','" + rs!Description + "','" + IIf(IsNull(rs!linea), "", rs!linea) + "','" + CStr(var_es_Catalogo) + "', 0," + CStr(rs!unit_selling_price) + "," + CStr(rs!cantidad_surtida) + " , 0,'" + IIf(IsNull(rs!REASON_CODE), "", rs!REASON_CODE) + "'," + CStr(rs!CANTIDAD_NEGADA) + ")"
                           Else
                              var_cadena = var_cadena + "(" + CStr(var_consecutivo) + "," + VAR_FECHA_INICIO_TABLA + "," + VAR_FECHA_FIN_TABLA + ",'" + CStr(rs!ORDERED_DATE) + "','','" + rs!LINE_STATUS + "','','','" + CStr(rs!collector_id) + "','" + rs!Name + "','" + CStr(rs!CUST_ACCT_SITE_ID) + "','" + Mid(rs!CUSTOMER_NAME, 1, 50) + "'," + CStr(rs!ORDER_NUMBER) + ",'" + rs!ordered_item + "','" + rs!Description + "','" + IIf(IsNull(rs!linea), "", rs!linea) + "','" + CStr(var_es_Catalogo) + "', 0," + CStr(rs!unit_selling_price) + "," + CStr(rs!cantidad_surtida) + " , " + CStr(rs!CANTIDAD_NEGADA) + ",'" + IIf(IsNull(rs!REASON_CODE), "", rs!REASON_CODE) + "',0)"
                           End If
                           
                        Else
                           If IIf(IsNull(rs!REASON_CODE), "", rs!REASON_CODE) = "PRODUCCION" Then
                              var_cadena = var_cadena + "(" + CStr(var_consecutivo) + "," + VAR_FECHA_INICIO_TABLA + "," + VAR_FECHA_FIN_TABLA + ",'" + CStr(rs!ORDERED_DATE) + "','','" + rs!LINE_STATUS + "','','','" + CStr(rs!collector_id) + "','" + rs!Name + "','" + CStr(rs!CUST_ACCT_SITE_ID) + "','" + Mid(rs!CUSTOMER_NAME, 1, 50) + "'," + CStr(rs!ORDER_NUMBER) + ",'" + rs!ordered_item + "','" + rs!Description + "','" + IIf(IsNull(rs!linea), "", rs!linea) + "','" + CStr(var_es_Catalogo) + "', " + CStr(rs!cantidad_pedida) + "," + CStr(rs!unit_selling_price) + "," + CStr(rs!cantidad_surtida) + " , 0,'" + IIf(IsNull(rs!REASON_CODE), "", rs!REASON_CODE) + "'," + CStr(rs!CANTIDAD_NEGADA) + ")"
                           Else
                              var_cadena = var_cadena + "(" + CStr(var_consecutivo) + "," + VAR_FECHA_INICIO_TABLA + "," + VAR_FECHA_FIN_TABLA + ",'" + CStr(rs!ORDERED_DATE) + "','','" + rs!LINE_STATUS + "','','','" + CStr(rs!collector_id) + "','" + rs!Name + "','" + CStr(rs!CUST_ACCT_SITE_ID) + "','" + Mid(rs!CUSTOMER_NAME, 1, 50) + "'," + CStr(rs!ORDER_NUMBER) + ",'" + rs!ordered_item + "','" + rs!Description + "','" + IIf(IsNull(rs!linea), "", rs!linea) + "','" + CStr(var_es_Catalogo) + "', " + CStr(rs!cantidad_pedida) + "," + CStr(rs!unit_selling_price) + "," + CStr(rs!cantidad_surtida) + " , " + CStr(rs!CANTIDAD_NEGADA) + ",'" + IIf(IsNull(rs!REASON_CODE), "", rs!REASON_CODE) + "',0)"
                           End If
                        End If
                        'MsgBox var_cadena
                        rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        rs.MoveNext
                  Wend
                  rsaux.Open "SELECT DISTINCT CLAVE_AGENTE FROM TB_TEMP_ORACLE_REPORTE_NEGADO WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux.EOF
                        var_cadena = "SELECT hcsu.TERRITORY_ID FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc, qp_secu_list_headers_v plist, hz_contact_points email, hz_contact_points phone WHERE hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id "
                        var_cadena = var_cadena + " AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id AND hcsu.price_list_id = plist.list_header_id(+) AND email.contact_point_type(+) = 'EMAIL' AND email.owner_table_name(+) = 'HZ_PARTY_SITES' AND email.owner_table_id(+) = hcas.party_site_id AND phone.contact_point_type(+) = 'PHONE' AND phone.owner_table_name(+) = 'HZ_PARTY_SITES' AND phone.owner_table_id(+) = hcas.party_site_id"
                        var_cadena = var_cadena + " AND hcp.collector_id = " + CStr(IIf(IsNull(rsaux!CLAVE_AGENTE), 0, rsaux!CLAVE_AGENTE)) + " AND ROWNUM = 1"
                        rsaux1.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux1.EOF Then
                           rsaux2.Open "SELECT RT.TERRITORY_ID,RT.SEGMENT4 AS CANAL, CANAL.DESCRIPCION AS CANAL_DESC FROM XXVIA_AR_TERRITORIOS_SEG_V CANAL, RA_TERRITORIES RT WHERE CANAL.TIPO = 'CANAL' AND CANAL.VALOR = RT.SEGMENT4 and RT.TERRITORY_ID = " + CStr(IIf(IsNull(rsaux1!territory_id), 0, rsaux1!territory_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux2.EOF Then
                              rsaux3.Open "update TB_TEMP_ORACLE_REPORTE_NEGADO SET CANAL_VENTA = '" + CStr(IIf(IsNull(rsaux2!canal), "", rsaux2!canal)) + "', NOMBRE_CANAL_VENTA = '" + IIf(IsNull(rsaux2!CANAL_DESC), "", rsaux2!CANAL_DESC) + "' WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND CLAVE_AGENTE = '" + rsaux!CLAVE_AGENTE + "'", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux2.Close
                        End If
                        rsaux1.Close
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  rsaux.Open "SELECT * FROM TB_TEMP_ORACLE_REPORTE_NEGADO WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND CAUSA_NEGADO = 'PRODUCCION'", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux.EOF
                        rsaux1.Open "UPDATE TB_TEMP_ORACLE_REPORTE_NEGADO SET CAUSA_NEGADO = 'PRODUCCION' WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND CODIGO = '" + rsaux!codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  Set reporte = appl.OpenReport(App.Path + "\rep_oracle_negado.rpt")
                  reporte.RecordSelectionFormula = "{VW_ORACLE_NEGADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\Logistica_negado" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
              Else
                 MsgBox "No hay información", vbOKOnly, "ATENCION"
              End If
              rs.Close
              rs.Open "delete from TB_TEMP_ORACLE_REPORTE_NEGADO WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
           Else
              MsgBox "No existen pedidos para el periodo indicado", vbOKOnly, "ATENCION"
           End If
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3500
   txt_inicio = Date
   txt_fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_fin_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar la fecha"
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

Private Sub txt_fin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub txt_fin_LostFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = ""
End Sub

Private Sub txt_inicio_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar la fecha"
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

Private Sub txt_inicio_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_inicio_LostFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = ""
End Sub

