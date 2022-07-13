VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmconcentrado_orden_surtido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Concentrado de Orden de Surtido"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   9960
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      Picture         =   "frmconcentrado_orden_surtido.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame6 
      Caption         =   " Agrupado "
      Height          =   930
      Left            =   90
      TabIndex        =   33
      Top             =   2505
      Width           =   4050
      Begin VB.OptionButton opt_agrupado 
         Caption         =   "Agente"
         Height          =   345
         Left            =   210
         TabIndex        =   34
         Top             =   375
         Width           =   885
      End
   End
   Begin VB.Frame frm_horas 
      Height          =   1470
      Left            =   2115
      TabIndex        =   30
      Top             =   4395
      Width           =   2040
      Begin VB.ListBox lst_horas 
         Height          =   1035
         Left            =   60
         TabIndex        =   31
         Top             =   375
         Width           =   1905
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   " Canales de Venta"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   32
         Top             =   120
         Width           =   1965
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   " Hora "
      Height          =   705
      Left            =   105
      TabIndex        =   29
      Top             =   4560
      Width           =   4050
      Begin VB.TextBox txt_hora 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   210
         Width           =   1065
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Fecha "
      Height          =   705
      Left            =   90
      TabIndex        =   28
      Top             =   3495
      Width           =   4050
      Begin VB.TextBox txt_fecha_inicio 
         Height          =   315
         Left            =   1545
         TabIndex        =   8
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   1050
         TabIndex        =   35
         Top             =   330
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Tipo Reporte "
      Height          =   945
      Left            =   120
      TabIndex        =   21
      Top             =   495
      Width           =   4050
      Begin VB.OptionButton opt_tipo_reporte_2 
         Caption         =   "Detalle de Artículos"
         Height          =   405
         Left            =   225
         TabIndex        =   5
         Top             =   510
         Width           =   2340
      End
      Begin VB.OptionButton opt_tipo_reporte_1 
         Caption         =   "Agrupado por Linea"
         Height          =   315
         Left            =   225
         TabIndex        =   4
         Top             =   225
         Width           =   1740
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Filtrado "
      Height          =   930
      Left            =   120
      TabIndex        =   20
      Top             =   1530
      Width           =   4050
      Begin VB.OptionButton opt_tipo_filtrado_1 
         Caption         =   "Agente"
         Height          =   345
         Left            =   210
         TabIndex        =   6
         Top             =   210
         Width           =   1500
      End
      Begin VB.OptionButton opt_tipo_filtrado_2 
         Caption         =   "Canal de Venta"
         Height          =   255
         Left            =   210
         TabIndex        =   7
         Top             =   570
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9465
      Picture         =   "frmconcentrado_orden_surtido.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   375
      Picture         =   "frmconcentrado_orden_surtido.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmconcentrado_orden_surtido.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   -15
      TabIndex        =   0
      Top             =   345
      Width           =   9870
   End
   Begin VB.Frame frm_agentes 
      Height          =   3690
      Left            =   4215
      TabIndex        =   22
      Top             =   495
      Width           =   5610
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         Picture         =   "frmconcentrado_orden_surtido.frx":0940
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   30
         Picture         =   "frmconcentrado_orden_surtido.frx":0B56
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         Picture         =   "frmconcentrado_orden_surtido.frx":0C58
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   690
         Picture         =   "frmconcentrado_orden_surtido.frx":0D2A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Marcar (Enter)"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Picture         =   "frmconcentrado_orden_surtido.frx":0F74
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   360
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_agentes 
         Height          =   2955
         Left            =   75
         TabIndex        =   23
         Top             =   690
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   5212
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6967
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   " Agentes "
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   24
         Top             =   120
         Width           =   5535
      End
   End
   Begin VB.Frame frm_canales 
      Height          =   3690
      Left            =   4215
      TabIndex        =   25
      Top             =   495
      Width           =   5610
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Picture         =   "frmconcentrado_orden_surtido.frx":118A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   690
         Picture         =   "frmconcentrado_orden_surtido.frx":13A0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Marcar (Enter)"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         Picture         =   "frmconcentrado_orden_surtido.frx":15EA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   30
         Picture         =   "frmconcentrado_orden_surtido.frx":16BC
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         Picture         =   "frmconcentrado_orden_surtido.frx":17BE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   360
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_canales 
         Height          =   2925
         Left            =   60
         TabIndex        =   26
         Top             =   690
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   5159
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6967
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Canales de Venta"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   27
         Top             =   120
         Width           =   5535
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   2100
      Top             =   45
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
      Left            =   2655
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "frmconcentrado_orden_surtido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Private Sub cmd_imprimir_Click()
Dim var_n As Integer
Dim var_i As Integer
Dim var_primera_vez As Integer
Dim var_agentes As Integer
Dim var_cadena As String
var_primera_vez = 1
var_agentes = 0
rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
If IsDate(txt_fecha_inicio) Then
   If opt_tipo_filtrado_2.Value = True Then
      var_n = lv_canales.ListItems.Count
      VAR_cADENA_cANALES = ""
      For var_i = 1 To var_n
          lv_canales.ListItems(var_i).Selected = True
          If lv_canales.selectedItem.SubItems(2) = "*" Then
             If VAR_cADENA_cANALES = "" Then
                VAR_cADENA_cANALES = "'" + Me.lv_canales.selectedItem + "'"
             Else
                VAR_cADENA_cANALES = VAR_cADENA_cANALES + ",'" + Me.lv_canales.selectedItem + "'"
             End If
          End If
      Next var_i
      If VAR_cADENA_cANALES <> "" Then
         VAR_CADENA_TERRITORIOS = ""
         rs.Open "SELECT DISTINCT RT.TERRITORY_ID,RT.SEGMENT4 FROM XXVIA_AR_TERRITORIOS_SEG_V CANAL, RA_TERRITORIES RT WHERE CANAL.TIPO = 'CANAL' AND CANAL.VALOR = RT.SEGMENT4 AND RT.SEGMENT4 IN (" + VAR_cADENA_cANALES + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               If VAR_CADENA_TERRITORIOS = "" Then
                  VAR_CADENA_TERRITORIOS = CStr(rs!territory_id)
               Else
                  VAR_CADENA_TERRITORIOS = VAR_CADENA_TERRITORIOS + "," + CStr(rs!territory_id)
               End If
               rs.MoveNext
         Wend
         rs.Close
         var_cadena_agentes = ""
         If VAR_CADENA_TERRITORIOS <> "" Then
            var_cadena = "SELECT  DISTINCT hcp.collector_id FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc, qp_secu_list_headers_v plist, hz_contact_points email, hz_contact_points phone WHERE hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id "
            var_cadena = var_cadena + " AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id AND hcsu.price_list_id = plist.list_header_id(+) AND email.contact_point_type(+) = 'EMAIL' AND email.owner_table_name(+) = 'HZ_PARTY_SITES' AND email.owner_table_id(+) = hcas.party_site_id AND phone.contact_point_type(+) = 'PHONE' AND phone.owner_table_name(+) = 'HZ_PARTY_SITES' AND phone.owner_table_id(+) = hcas.party_site_id"
            var_cadena = var_cadena + " AND hcsu.TERRITORY_ID IN(" + VAR_CADENA_TERRITORIOS + ")"
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  If var_cadena_agentes = "" Then
                     var_cadena_agentes = CStr(rs!collector_id)
                  Else
                     var_cadena_agentes = var_cadena_agentes + "," + CStr(rs!collector_id)
                  End If
                  rs.MoveNext
            Wend
            rs.Close
         End If
      End If
   End If
   var_primera_vez = 1
   
   If opt_tipo_filtrado_1.Value = True Then
      var_cadena_agentes = ""
      var_n = lv_agentes.ListItems.Count
      For var_i = 1 To var_n
          lv_agentes.ListItems.Item(var_i).Selected = True
          If lv_agentes.selectedItem.SubItems(2) = "*" Then
             If var_primera_vez = 1 Then
                var_primera_vez = 0
                var_cadena_agentes = lv_agentes.selectedItem
             Else
                var_cadena_agentes = var_cadena_agentes + "," + lv_agentes.selectedItem
             End If
          End If
      Next var_i
   End If
   If Trim(var_cadena_agentes) <> "" Then
   
      var_fecha_fin_1 = CDate(txt_fecha_inicio) + 1
      var_dia = CStr(Day(CDate(txt_fecha_inicio)))
      var_mes = CStr(Month(CDate(txt_fecha_inicio)))
      var_año = CStr(Year(CDate(txt_fecha_inicio)))
      If Len(Trim(var_dia)) = 1 Then
         var_dia = "0" + var_dia
      End If
      If Len(Trim(var_mes)) = 1 Then
         var_mes = "0" + var_mes
      End If
      var_fecha_inicio = var_dia + "-" + var_mes + "-" + var_año
          
           
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
      var_fecha_sql = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
      var_cadena_agentes = var_cadena_agentes
      'var_cadena = "SELECT DISTINCT SOURCE_HEADER_NUMBER AS PEDIDO from WSH_DLVB_DLVY_V B , hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where A.delivery_id = B.delivery_id AND A.delivery_detail_id = B.delivery_detail_id AND HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id and TRUNC(B.LAST_UPDATE_DATE) >= to_date('" + var_fecha_inicio + "','DD-MM-YYYY') AND TRUNC(B.LAST_UPDATE_DATE) < TO_DATE('" + var_fecha_fin + "','DD-MM-YYYY')"
      'var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND E.COLLECTOR_ID IN (" + var_cadena_agentes + ") AND OHA.SHIP_FROM_ORG_ID = " + var_unidad_organizacional + " ORDER BY SOURCE_HEADER_NUMBER"
      'Text1 = var_cadena
      var_cadena = "SELECT DISTINCT SOURCE_HEADER_NUMBER AS PEDIDO from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors e Where  HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id and TRUNC(a.LAST_UPDATE_DATE) >= to_date('" + var_fecha_inicio + "','DD-MM-YYYY') AND TRUNC(a.LAST_UPDATE_DATE) < TO_DATE('" + var_fecha_fin + "','DD-MM-YYYY') AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id "
      var_cadena = var_cadena + " AND E.COLLECTOR_ID IN (" + var_cadena_agentes + ") AND OHA.SHIP_FROM_ORG_ID = " + CStr(var_unidad_organizacional) + " and a.released_status  = 'Y' ORDER BY SOURCE_HEADER_NUMBER"

      Text1 = var_cadena
      rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rsaux2.EOF Then
         cnn.BeginTrans
         rs.Open "select max(inte_tEM_consecutivo) from tb_Temp_oracle_concentrado_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            var_consecutivo = 1
         Else
            var_consecutivo = IIf(IsNull(rs(0).Value), 1, rs(0).Value + 1)
         End If
         rs.Close
         rs.Open "insert into tb_Temp_oracle_concentrado_orden_surtido (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
         
         var_cadena_Pedidos = ""
         While Not rsaux2.EOF
               If Trim(var_cadena_Pedidos) = "" Then
                  var_cadena_Pedidos = CStr(rsaux2!PEDIDO)
               Else
                  var_cadena_Pedidos = var_cadena_Pedidos + "," + CStr(rsaux2!PEDIDO)
               End If
               rsaux2.MoveNext
         Wend
         
         
        'var_cadena = "SELECT OL.LINE_ID, g.organization_id, oh.ordered_date,HCSU.CUST_ACCT_SITE_ID  AS CLAVE_CLIENTE, HL.ADDRESS1 AS CUSTOMER_NAME, HCSU2.CUST_ACCT_SITE_ID AS clave_establecimiento, HL2.ADDRESS1 AS nombre_Establecimiento, e.collector_id, E.NAME, oh.header_id, oh.order_number, oh.transactional_curr_code, NVL(ol.ordered_quantity,0)   AS CANTIDAD_PEDIDA, NVL(ol.cancelled_quantity,0) AS CANTIDAD_NEGADA, NVL(ol.shipped_quantity,0)   AS CANTIDAD_surtida, ol.line_id, ol.ordered_item, g.description, ol.order_quantity_uom, ol.inventory_item_id, ol.price_list_id, ol.unit_selling_price, ol.unit_list_price, DECODE(ol.cancelled_flag,'Y','CANCELADA','SURTIDA') line_status, ol.flow_status_code, NVL(OER.REASON_CODE,'') AS REASON_CODE, h.linea   FROM oe_order_headers_all oh, oe_order_lines_all ol, OE_ORDER_LINES_HISTORY OLH, OE_REASONS OER, hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, hz_cust_acct_sites_all HCAS2, HZ_PARTY_SITES HPS2, "
        ' var_cadena = var_cadena + " HZ_LOCATIONS HL2, HZ_CUST_SITE_USES_ALL HCSU2, hz_customer_profiles D2, ar_collectors E, hz_cust_acct_sites_all f, hz_cust_acct_sites_all f2, xxvia_system_items_b g, xxvia_vw_articulos_cat h WHERE order_number  IN (" + var_cadena_pedidos + ") AND oh.header_id = ol.header_id AND ol.ship_from_org_id = " + var_unidad_organizacional + " AND oL.header_id = oLh.header_id(+) AND OL.LINE_ID = OLH.LINE_ID(+) AND OER.HEADER_ID(+) = oL.header_id AND OER.ENTITY_ID(+) = OL.LINE_ID AND HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID = HL.LOCATION_ID AND HCSU.SITE_USE_ID = OH.INVOICE_TO_ORG_ID  AND HCSU.SITE_USE_ID = D.site_use_id AND HCSU.CUST_ACCT_SITE_ID  = HCAS.CUST_ACCT_SITE_ID AND f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID AND HCAS2.PARTY_SITE_ID = HPS2.PARTY_SITE_ID AND HPS2.LOCATION_ID = HL2.LOCATION_ID AND HCSU2.SITE_USE_ID = OH.ship_TO_ORG_ID AND HCSU2.SITE_USE_ID = D2.site_use_id AND HCSU2.CUST_ACCT_SITE_ID = HCAS2.CUST_ACCT_SITE_ID "
        ' var_cadena = var_cadena + " AND f2.cust_acct_site_id    = HCAS2.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id AND ol.inventory_item_id = g.inventory_item_id AND g.organization_id = ol.ship_from_org_id AND (OER.REASON_CODE <> 'PRODUCCION' OR OER.REASON_CODE IS NULL) AND h.item_id = g.inventory_item_id AND h.organization_id = g.organization_id"
        
        'var_Cadena_pedidos = "435095"
        
        var_cadena = "SELECT  OH.source_document_id, OL.source_header_type_name,g.organization_id, oh.ordered_date, HCSU.CUST_ACCT_SITE_ID  AS CLAVE_CLIENTE, HL.ADDRESS1 AS CUSTOMER_NAME, HCSU2.CUST_ACCT_SITE_ID AS clave_establecimiento, HL2.ADDRESS1 AS nombre_Establecimiento, e.collector_id, E.NAME, oh.header_id, oh.order_number, oh.transactional_curr_code, NVL(ol.requested_quantity,0)   AS CANTIDAD_PEDIDA, 0 AS CANTIDAD_NEGADA, 0   AS CANTIDAD_surtida, ol.source_line_id, g.segment1 as ordered_item, g.description, 0 as order_quantity_uom, ol.inventory_item_id,  ola.unit_selling_price unit_price, h.linea, h.volume_uom_code, h.unit_volume, j.name as vendedor"
        var_cadena = var_cadena + " FROM oe_order_lines_all ola,oe_order_headers_all oh, WSH_DELIVERABLES_V ol, OE_ORDER_LINES_HISTORY OLH, OE_REASONS OER, hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, hz_cust_acct_sites_all HCAS2, HZ_PARTY_SITES HPS2, HZ_LOCATIONS HL2, HZ_CUST_SITE_USES_ALL HCSU2, hz_customer_profiles D2, ar_collectors E, hz_cust_acct_sites_all f, hz_cust_acct_sites_all f2, xxvia_system_items_b g, "
        var_cadena = var_cadena + " xxvia_vw_articulos_cat h, XXVIA_VENDEDORES J WHERE order_number  IN (" + var_cadena_Pedidos + ") AND oh.header_id = ol.source_header_id AND ol.organization_id = " + var_unidad_organizacional + " AND oL.source_header_id = oLh.header_id(+) AND OL.source_LINE_ID = OLH.LINE_ID(+) AND OER.HEADER_ID(+) = oL.source_header_id AND OER.ENTITY_ID(+) = OL.source_LINE_ID AND HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID = HL.LOCATION_ID AND HCSU.SITE_USE_ID = OH.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND HCSU.CUST_ACCT_SITE_ID  = HCAS.CUST_ACCT_SITE_ID AND f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID AND HCAS2.PARTY_SITE_ID = HPS2.PARTY_SITE_ID AND HPS2.LOCATION_ID = HL2.LOCATION_ID AND HCSU2.SITE_USE_ID = OH.ship_TO_ORG_ID AND HCSU2.SITE_USE_ID = D2.site_use_id(+) AND HCSU2.CUST_ACCT_SITE_ID = HCAS2.CUST_ACCT_SITE_ID AND f2.cust_acct_site_id    = HCAS2.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id "
        var_cadena = var_cadena + " AND ol.inventory_item_id = g.inventory_item_id AND g.organization_id = ol.organization_id AND h.item_id = g.inventory_item_id"
        var_cadena = var_cadena + " AND h.organization_id = g.organization_id and released_status = 'Y' and oh.salesrep_id = j.salesrep_id and ola.line_id = ol.source_line_id"
        
         
         
         rsaux3.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux3.EOF
            If rsaux3!ORDERED_ITEM = "00002692" Then
               var_x = var_x
            End If
            If rsaux3!Linea = "CATALOGOS" Then
               var_es_Catalogo = 1
            Else
               If rsaux3!Linea = "POP" Or rsaux3!Linea = "EMPAQUE" Then
                  var_es_Catalogo = 1
                  'MsgBox rsaux3!ordered_item
               Else
                  var_es_Catalogo = 0
               End If
            End If
            
            var_establecimiento = ""
            VAR_NOMBRE_ESTABLECIMIENTO = ""
            VAR_AGENTE_str = ""
            
            If rsaux3!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rsaux3!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
               rsaux4.Open "SELECT A.ATTRIBUTE1, B.description, B.secondary_inventory_name FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux3!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                  var_establecimiento = rsaux3!clave_establecimiento
                  VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux4!Description), "", rsaux4!Description)
                  VAR_AGENTE_str = IIf(IsNull(rsaux4!Description), "", rsaux4!Description)
               End If
               rsaux4.Close
            Else
               var_establecimiento = rsaux3!clave_establecimiento
               VAR_NOMBRE_ESTABLECIMIENTO = rsaux3!nombre_Establecimiento
               VAR_AGENTE_str = CStr(rsaux3!collector_id)
            End If
            
            var_cadena = "insert into tb_Temp_oracle_concentrado_orden_surtido (INVENTORY_ITEM_ID, INTE_TEM_CONSECUTIVO, PEDIDO, FECHA, AGENTE, NOMBRE_AGENTE, ARTICULO, DESCRIPCION, VOLUMEN, PRECIO,  CANTIDAD, LINEA, ES_CATALOGO, CLIENTE, NOMBRE_CLIENTE, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, LINE_ID, VENDEDOR) values (" + CStr(rsaux3!inventory_item_id) + "," + CStr(var_consecutivo) + "," + CStr(rsaux3!order_number) + " ," + var_fecha_sql + " ,'" + CStr(VAR_AGENTE_str) + "','" + rsaux3!Name + "' ,'" + rsaux3!ORDERED_ITEM + "', '" + rsaux3!Description + "', " + CStr(Round(IIf(IsNull(rsaux3!unit_volume), 0, rsaux3!unit_volume), 4)) + ", " + CStr(rsaux3!unit_price) + "," + CStr(rsaux3!CANTIDAD_PEDIDA + rsaux3!CANTIDAD_NEGADA) + ",'" + rsaux3!Linea + "'," + CStr(var_es_Catalogo) + ",'" + CStr(rsaux3!clave_cliente) + "','"
            var_cadena = var_cadena + Replace(rsaux3!customer_name, "'", " ") + "', '" + CStr(var_establecimiento) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "'," + CStr(rsaux3!source_LINE_ID) + ",'" + IIf(IsNull(rsaux3!VENDEDOR), "", rsaux3!VENDEDOR) + "')"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            rsaux3.MoveNext
         Wend
         rsaux3.Close
         'MsgBox var_consecutivo
         
         rsaux3.Open "SELECT DISTINCT AGENTE FROM tb_Temp_oracle_concentrado_orden_surtido WHERE NOMBRE_AGENTE <> 'INTERCOMPAÑIAS'", cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux3.EOF
               rsaux4.Open "SELECT PERSON_ID, email_address FROM PER_ALL_PEOPLE_F, AR_COLLECTORS WHERE COLLECTOR_ID = " + CStr(rsaux3!Agente) + " AND PERSON_ID = EMPLOYEE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                  If IsNumeric(rsaux3!Agente) Then
                     rsaux5.Open "UPDATE tb_Temp_oracle_concentrado_orden_surtido SET EMPLEADO = " + CStr(IIf(IsNull(rsaux4!PERSON_ID), 0, rsaux4!PERSON_ID)) + ", CORREO = '" + IIf(IsNull(rsaux4!email_address), "", rsaux4!email_address) + "' WHERE INTE_TEM_cONSECUTIVO = " + CStr(var_consecutivo) + " AND AGENTE = '" + CStr(rsaux3!Agente) + "'", cnn, adOpenDynamic, adLockOptimistic
                  End If
               End If
               rsaux4.Close
               rsaux3.MoveNext
         Wend
         rsaux3.Close
                  
         rsaux3.Open "delete from tb_Temp_oracle_concentrado_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and fecha is null", cnn, adOpenDynamic, adLockOptimistic
         
         If opt_tipo_reporte_2.Value = True Then
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_concentrado_orden_surtido_detalle.rpt")
            var_cadena = "{VW_ORACLE_CONCENTRADO_ORDENES_SURTIDO_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
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
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_concentrado_orden_surtido_detalle.rpt")
               var_cadena = "{VW_ORACLE_CONCENTRADO_ORDENES_SURTIDO_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
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
               
               
               
                     Set oexcel = CreateObject("Excel.Application")
                     Set owbook = oexcel.Workbooks.Add
                     Set osheet = owbook.Worksheets(1)
                     osheet.Name = "CONCENTRADO OS"
                     Screen.MousePointer = vbHourglass
                     iFila = 1
                     ifila2 = 1
                     icol2 = 1
                     iCol = 1
                     var_cadena = "select PEDIDO, CLIENTE, NOMBRE_CLIENTE, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, CANTIDAD, VOLUMEN, ES_CATALOGO from VW_ORACLE_CONCENTRADO_ORDENES_SURTIDO_CLIENTE where INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and CLIENTE is not null"
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
                     owbook.SaveAs archivo = "c:\reportessid\concentrado_ordenes_surtido_RESUMEN_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     oexcel.Visible = True
                     Set oexcel = Nothing
                     Screen.MousePointer = vbDefault
                     rsaux10.Close
               
               
               
               
            End If
         
         
         
         End If
         If opt_tipo_reporte_1.Value = True Then
            'Este es el reporte que se tiene que modificar
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_concentrado_orden_surtido_linea.rpt")
            var_cadena = "{VW_ORACLE_CONCENTRADO_ORDEN_SURTIDO_LINEAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
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
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_concentrado_orden_surtido_linea.rpt")
               var_cadena = "{VW_ORACLE_CONCENTRADO_ORDEN_SURTIDO_LINEAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
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
               
               
                     Set oexcel = CreateObject("Excel.Application")
                     Set owbook = oexcel.Workbooks.Add
                     Set osheet = owbook.Worksheets(1)
                     osheet.Name = "CONCENTRADO OS"
                     Screen.MousePointer = vbHourglass
                     iFila = 1
                     ifila2 = 1
                     icol2 = 1
                     iCol = 1
                     var_cadena = "select PEDIDO, CLIENTE, NOMBRE_CLIENTE, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, CANTIDAD, VOLUMEN, ES_CATALOGO, TOTAL_PEDIDOS from VW_ORACLE_CONCENTRADO_ORDENES_SURTIDO_CLIENTE where INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and CLIENTE is not null"
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
                     archivo = "c:\reportessid\concentrado_ordenes_surtido_RESUMEN_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     owbook.SaveAs archivo
                     oexcel.Visible = True
                     Set oexcel = Nothing
                     Screen.MousePointer = vbDefault
                     rsaux10.Close
               
               
               
            End If
            var_si = MsgBox("¿Desea enviar los correos a los agentes?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               rsaux3.Open "SELECT DISTINCT AGENTE, nombre_agente, correo FROM  tb_Temp_oracle_concentrado_orden_surtido WHERE CORREO IS NOT NULL AND CORREO <> '' AND INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " GROUP BY INTE_TEM_CONSECUTIVO, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CORREO, PEDIDO", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux3.EOF
                    'var_correo_electronico = IIf(IsNull(rsaux3!correo), "", rsaux3!correo)
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
                       MAPIMessages1.MsgSubject = "Información del pedido cargados " + CStr(IIf(IsNull(rsaux3!Agente), "", rsaux3!Agente)) + " " + IIf(IsNull(rsaux3!NOMBRE_AGENTE), "", rsaux3!NOMBRE_AGENTE)

                       MAPIMessages1.MsgNoteText = "Se anexa archivo con información de los pedido cargados del agente  " + CStr(IIf(IsNull(rsaux3!Agente), "", rsaux3!Agente)) + " " + IIf(IsNull(rsaux3!NOMBRE_AGENTE), "", rsaux3!NOMBRE_AGENTE)
                       var_archivo = App.Path & "\agente_" + Trim(CStr(IIf(IsNull(rsaux3!Agente), "", rsaux3!Agente))) + "_" + CStr(Day(Date)) + "_" + CStr(Month(Date)) + "_" + CStr(Year(Date)) + ".txt"
                       Open (App.Path & "\agente_" + Trim(CStr(IIf(IsNull(rsaux3!Agente), "", rsaux3!Agente))) + "_" + CStr(Day(Date)) + "_" + CStr(Month(Date)) + "_" + CStr(Year(Date)) + ".txt") For Output As #1
                       rsaux4.Open "SELECT INTE_TEM_CONSECUTIVO, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CORREO, PEDIDO, SUM(CANTIDAD) AS CANTIDAD FROM  tb_Temp_oracle_concentrado_orden_surtido WHERE  AGENTE = " + Trim(CStr(IIf(IsNull(rsaux3!Agente), "", rsaux3!Agente))) + " AND INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " GROUP BY INTE_TEM_CONSECUTIVO, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CORREO, PEDIDO", cnn, adOpenDynamic, adLockOptimistic
                       While Not rsaux4.EOF
                             Print #1, "PEDIDO: " + Trim(CStr(IIf(IsNull(rsaux4!PEDIDO), "", rsaux4!PEDIDO))) + ",    CLIENTE: " + CStr(IIf(IsNull(rsaux4!nombre_cliente), "", rsaux4!nombre_cliente)) + ",    CANTIDAD: " + Format(CStr(IIf(IsNull(rsaux4!cantidad), 0, rsaux4!cantidad)), "###,###,##0.00") + " PIEZAS"
                             rsaux4.MoveNext
                       Wend
                       rsaux4.Close
                       Close #1
                       MAPIMessages1.AttachmentPathName = var_archivo
                       MAPIMessages1.send True
                       If MAPISession1.SessionID > 0 Then
                          MAPISession1.SignOff
                       End If
                     End If
                     rsaux3.MoveNext
               Wend
               rsaux3.Close
            End If
         End If
         rs.Open "DELETE FROM TB_TEMP_CONCENTRADO_ORDEN_SURTIDO_AGENTES WHERE INTE_TMP_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      Else
         MsgBox "Ne existen pedidos para esta solicitud", vbOKOnly, "ATENCION"
      End If
      If rsaux2.State = 1 Then
         rsaux2.Close
      End If
   Else
      If opt_tipo_filtrado_2.Value = True Then
         MsgBox "No se a seleccionado algún canal de venta", vbOKOnly, "ATENCION"
      Else
         MsgBox "No se a seleccionado algún agente", vbOKOnly, "ATENCION"
      End If
   End If
Else
   MsgBox "Fecha Incorrecta", vbOKOnly, "ATENCION"
End If
   
End Sub

Private Sub cmd_invertir_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_agentes.ListItems.Count
   For i = 1 To n
       If lv_agentes.ListItems.Item(i).SubItems(2) = "*" Then
          lv_agentes.ListItems.Item(i).SubItems(2) = " "
          lv_agentes.ListItems.Item(i).Bold = False
          lv_agentes.ListItems.Item(i).ForeColor = &H80000012
          lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
          lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
          lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
          lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
       Else
          lv_agentes.ListItems.Item(i).SubItems(2) = "*"
          lv_agentes.ListItems.Item(i).Bold = True
          lv_agentes.ListItems.Item(i).ForeColor = &H8000&
          lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
          lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
          lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
          lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
      End If
   Next
   lv_agentes.Refresh
End Sub

Private Sub cmd_marcar_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   i = lv_agentes.selectedItem.Index
   If lv_agentes.selectedItem.SubItems(2) = "*" Then
       lv_agentes.ListItems.Item(i).SubItems(2) = " "
       lv_agentes.ListItems.Item(i).Bold = False
       lv_agentes.ListItems.Item(i).ForeColor = &H80000012
       lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
       lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
       lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Else
      lv_agentes.ListItems.Item(i).SubItems(2) = "*"
      lv_agentes.ListItems.Item(i).Bold = True
      lv_agentes.ListItems.Item(i).ForeColor = &H8000&
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
  End If
  lv_agentes.Refresh
End Sub

Private Sub cmd_mes_1_Click()

End Sub

Private Sub cmd_ninguno_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_agentes.ListItems.Count
   For i = 1 To n
       lv_agentes.ListItems.Item(i).SubItems(2) = " "
       lv_agentes.ListItems.Item(i).Bold = False
       lv_agentes.ListItems.Item(i).ForeColor = &H80000012
       lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
       lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
       lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
    Next
    lv_agentes.Refresh

End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   primera_vez = False
   segunda_vez = False
   n = lv_agentes.ListItems.Count
   For i = 1 To n
       If lv_agentes.ListItems.Item(i).SubItems(2) = "*" And primera_vez = False Then
          numero_seleccionado1 = i
          primera_vez = True
       End If
       If lv_agentes.ListItems.Item(i).SubItems(2) = "*" And primera_vez = True Then
          numero_seleccionado2 = i
       End If
   Next
   For i = numero_seleccionado1 To numero_seleccionado2
       lv_agentes.ListItems.Item(i).SubItems(2) = "*"
       lv_agentes.ListItems.Item(i).Bold = True
       lv_agentes.ListItems.Item(i).ForeColor = &H8000&
       lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
       lv_agentes.Refresh
   Next
End Sub

Private Sub cmd_todos_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_agentes.ListItems.Count
   For i = 1 To n
       lv_agentes.ListItems.Item(i).SubItems(2) = "*"
       lv_agentes.ListItems.Item(i).Bold = True
       lv_agentes.ListItems.Item(i).ForeColor = &H8000&
       lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
   Next
   lv_agentes.Refresh
End Sub

Private Sub Command1_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_canales.ListItems.Count
   For i = 1 To n
       lv_canales.ListItems.Item(i).SubItems(2) = "*"
       lv_canales.ListItems.Item(i).Bold = True
       lv_canales.ListItems.Item(i).ForeColor = &H8000&
       lv_canales.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_canales.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
   Next
   lv_canales.Refresh

End Sub

Private Sub Command2_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_canales.ListItems.Count
   For i = 1 To n
       lv_canales.ListItems.Item(i).SubItems(2) = " "
       lv_canales.ListItems.Item(i).Bold = False
       lv_canales.ListItems.Item(i).ForeColor = &H80000012
       lv_canales.ListItems.Item(i).ListSubItems(1).Bold = False
       lv_canales.ListItems.Item(i).ListSubItems(2).Bold = False
       lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
    Next
    lv_canales.Refresh
End Sub

Private Sub Command3_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_canales.ListItems.Count
   For i = 1 To n
       If lv_canales.ListItems.Item(i).SubItems(2) = "*" Then
          lv_canales.ListItems.Item(i).SubItems(2) = " "
          lv_canales.ListItems.Item(i).Bold = False
          lv_canales.ListItems.Item(i).ForeColor = &H80000012
          lv_canales.ListItems.Item(i).ListSubItems(1).Bold = False
          lv_canales.ListItems.Item(i).ListSubItems(2).Bold = False
          lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
          lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
       Else
          lv_canales.ListItems.Item(i).SubItems(2) = "*"
          lv_canales.ListItems.Item(i).Bold = True
          lv_canales.ListItems.Item(i).ForeColor = &H8000&
          lv_canales.ListItems.Item(i).ListSubItems(1).Bold = True
          lv_canales.ListItems.Item(i).ListSubItems(2).Bold = True
          lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
          lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
      End If
   Next
   lv_canales.Refresh
End Sub

Private Sub Command4_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   i = lv_canales.selectedItem.Index
   If lv_canales.selectedItem.SubItems(2) = "*" Then
      lv_canales.ListItems.Item(i).SubItems(2) = " "
      lv_canales.ListItems.Item(i).Bold = False
      lv_canales.ListItems.Item(i).ForeColor = &H80000012
      lv_canales.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_canales.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Else
      lv_canales.ListItems.Item(i).SubItems(2) = "*"
      lv_canales.ListItems.Item(i).Bold = True
      lv_canales.ListItems.Item(i).ForeColor = &H8000&
      lv_canales.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_canales.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
      lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
   End If
   lv_canales.Refresh
End Sub

Private Sub Command5_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   primera_vez = False
   segunda_vez = False
   n = lv_canales.ListItems.Count
   For i = 1 To n
       If lv_canales.ListItems.Item(i).SubItems(2) = "*" And primera_vez = False Then
          numero_seleccionado1 = i
          primera_vez = True
       End If
       If lv_canales.ListItems.Item(i).SubItems(2) = "*" And primera_vez = True Then
          numero_seleccionado2 = i
       End If
   Next
   For i = numero_seleccionado1 To numero_seleccionado2
       lv_canales.ListItems.Item(i).SubItems(2) = "*"
       lv_canales.ListItems.Item(i).Bold = True
       lv_canales.ListItems.Item(i).ForeColor = &H8000&
       lv_canales.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_canales.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
       lv_canales.Refresh
   Next
End Sub

Private Sub Command6_Click()
Dim var_n As Integer
Dim var_i As Integer
Dim var_primera_vez As Integer
Dim var_agentes As Integer
Dim var_cadena As String
var_primera_vez = 1
var_agentes = 0
rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
If IsDate(txt_fecha_inicio) Then
   var_fecha_fin_1 = CDate(txt_fecha_inicio) + 1
   var_dia = CStr(Day(CDate(txt_fecha_inicio)))
   var_mes = CStr(Month(CDate(txt_fecha_inicio)))
   var_año = CStr(Year(CDate(txt_fecha_inicio)))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes)) = 1 Then
      var_mes = "0" + var_mes
   End If
   var_fecha_inicio = var_dia + "/" + var_mes + "/" + var_año
          
           
   var_dia = CStr(Day(var_fecha_fin_1))
   var_mes = CStr(Month(var_fecha_fin_1))
   var_año = CStr(Year(var_fecha_fin_1))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes)) = 1 Then
      var_mes = "0" + var_mes
   End If
   var_fecha_fin = var_dia + "/" + var_mes + "/" + var_año
   
         
         
        
   var_cadena = "SELECT '" + CStr(var_fecha_inicio) + "' AS FECHA_INICIO, C.codigo  CODIGO, A.ITEM_DESCRIPTION DESCRIPCION, c.linea, sum(src_requested_quantity) CANTIDAD, ubicacion_bcp UBICACION  from WSH_DLVB_DLVY_V B, hz_cust_acct_sites_all  HCAS, HZ_PARTY_SITES HPS,  HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA,  WSH_DELIVERABLES_V A,  HZ_CUST_SITE_USES_ALL HCSU,  XXVIA_VW_CATEGORIAS_ITEM_B C, hz_customer_profiles D, ar_collectors E Where A.delivery_id = B.delivery_id AND A.delivery_detail_id = B.delivery_detail_id AND HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID "
   var_cadena = var_cadena + " AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND HCSU.SITE_USE_ID = D.site_use_id "
   var_cadena = var_cadena + " And TRUNC(B.LAST_UPDATE_DATE) >= to_date('" + var_fecha_inicio + "','DD-MM-YYYY') "
   var_cadena = var_cadena + " And TRUNC(B.LAST_UPDATE_DATE) < TO_DATE('" + var_fecha_fin + "','DD-MM-YYYY') "
   var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = C.ITEM_ID AND A.ORGANIZATION_ID = C.ORGANIZATION_ID  AND D.collector_id = e.collector_id  AND OHA.SHIP_FROM_ORG_ID = 93 and released_status = 'Y'       group by C.codigo, A.ITEM_DESCRIPTION, c.linea, ubicacion_bcp"
   rsaux10.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
   If Not rsaux10.EOF Then
      
      Set oexcel = CreateObject("Excel.Application")
      Set owbook = oexcel.Workbooks.Add
      Set osheet = owbook.Worksheets(1)
      osheet.Name = "PIEZAS A SURTIR"
      Screen.MousePointer = vbHourglass
      iFila = 1
      ifila2 = 1
      icol2 = 1
      iCol = 1
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
      owbook.SaveAs archivo = "c:\reportessid\PIEZAS_SURTIR_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
      oexcel.Visible = True
      Set oexcel = Nothing
      Screen.MousePointer = vbDefault
   Else
      MsgBox "Ne existen pedidos para esta solicitud", vbOKOnly, "ATENCION"
   End If
   rsaux10.Close
Else
   MsgBox "Fecha Incorrecta", vbOKOnly, "ATENCION"
End If

End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 1500
   Left = 900
   frm_horas.Visible = False
   frm_canales.Visible = False
   frm_agentes.Visible = True
   opt_tipo_filtrado_1.Value = True
   opt_tipo_reporte_1.Value = True
   'var_cadena = "SELECT MAX(ORDERED_dATE) AS FECHA from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id "
   var_cadena = "SELECT MAX(ORDERED_dATE) AS FECHA From OE_ORDER_HEADERS_ALL"
   rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      If Not IsNull(rs!Fecha) Then
         If IsDate(rs!Fecha) Then
            If rs!Fecha <= Date + 2 Then
               txt_fecha_inicio = rs!Fecha
            Else
               Me.txt_fecha_inicio = Date
            End If
         Else
            MsgBox "Fecha incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         txt_fecha_inicio = Date
      End If
   Else
      txt_fecha_inicio = Date
   End If
   rs.Close
   Dim list_item As ListItem
   rs.Open "select COLLECTOR_ID, NAME from ar_collectors", cnnoracle_4, adOpenDynamic, adLockOptimistic
   numero_items_agentes = 0
   While Not rs.EOF
      Set list_item = lv_agentes.ListItems.Add(, , rs!collector_id)
      list_item.SubItems(1) = IIf(IsNull(rs!Name), "", rs!Name)
      list_item.SubItems(2) = " "
      rs.MoveNext:
      numero_items_agentes = numero_items_agentes + 1
    Wend
    rs.Close
   If numero_items_agentes > 12 Then
      lv_agentes.ColumnHeaders(1).Width = lv_agentes.ColumnHeaders(1).Width - 200
   End If
   rsaux.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   rs.Open "SELECT DISTINCT RT.SEGMENT4 AS VCHA_CAN_CANAL_VENTA_ID,CANAL.DESCRIPCION AS VCHA_CAN_NOMBRE FROM XXVIA_AR_TERRITORIOS_SEG_V CANAL, RA_TERRITORIES RT WHERE CANAL.TIPO = 'CANAL' AND CANAL.VALOR = RT.SEGMENT4 ORDER BY canal.descripcion", cnnoracle_4, adOpenDynamic, adLockOptimistic
   numero_items_canales = 0
   While Not rs.EOF
      Set list_item = lv_canales.ListItems.Add(, , rs!vcha_can_canal_venta_id)
      list_item.SubItems(1) = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
      list_item.SubItems(2) = " "
      rs.MoveNext:
      numero_items_canales = numero_items_canales + 1
    Wend
    rs.Close
   If numero_items_canales > 12 Then
      lv_canales.ColumnHeaders(1).Width = lv_canales.ColumnHeaders(1).Width - 200
   End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lst_horas_DblClick()
   Dim n As Integer
   n = lst_horas.ListIndex
   txt_hora = lst_horas.List(n)
   frm_horas.Visible = False
End Sub

Private Sub lst_horas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim n As Integer
      n = lst_horas.ListIndex
      txt_hora = lst_horas.List(n)
      frm_horas.Visible = False
   End If
End Sub

Private Sub lv_agentes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_agentes, ColumnHeader)
End Sub

Private Sub lv_agentes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim numero_lineas As Integer
      Dim numero_seleccionado1 As Integer
      Dim numero_seleccionado2 As Integer
      Dim primera_vez As Boolean
      Dim segunda_vez As Boolean
      Dim i As Integer
      Dim n As Integer
      Dim list_item As ListItem
      i = lv_agentes.selectedItem.Index
      If lv_agentes.selectedItem.SubItems(2) = "*" Then
         lv_agentes.ListItems.Item(i).SubItems(2) = " "
         lv_agentes.ListItems.Item(i).Bold = False
         lv_agentes.ListItems.Item(i).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_agentes.ListItems.Item(i).SubItems(2) = "*"
         lv_agentes.ListItems.Item(i).Bold = True
         lv_agentes.ListItems.Item(i).ForeColor = &H8000&
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
      End If
      lv_agentes.Refresh
   End If
End Sub

Private Sub lv_canales_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_canales, ColumnHeader)
End Sub

Private Sub lv_canales_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim numero_lineas As Integer
      Dim numero_seleccionado1 As Integer
      Dim numero_seleccionado2 As Integer
      Dim primera_vez As Boolean
      Dim segunda_vez As Boolean
      Dim i As Integer
      Dim n As Integer
      Dim list_item As ListItem
      n = lv_canales.ListItems.Count
      i = lv_canales.selectedItem.Index
      If lv_canales.ListItems.Item(i).SubItems(2) = "*" Then
      lv_canales.ListItems.Item(i).SubItems(2) = " "
             lv_canales.ListItems.Item(i).Bold = False
             lv_canales.ListItems.Item(i).ForeColor = &H80000012
             lv_canales.ListItems.Item(i).ListSubItems(1).Bold = False
             lv_canales.ListItems.Item(i).ListSubItems(2).Bold = False
             lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
             lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
          Else
             lv_canales.ListItems.Item(i).SubItems(2) = "*"
             lv_canales.ListItems.Item(i).Bold = True
             lv_canales.ListItems.Item(i).ForeColor = &H8000&
             lv_canales.ListItems.Item(i).ListSubItems(1).Bold = True
             lv_canales.ListItems.Item(i).ListSubItems(2).Bold = True
             lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
             lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
         End If
      lv_canales.Refresh
   End If
End Sub

Private Sub opt_tipo_filtrado_1_Click()
   If opt_tipo_filtrado_1.Value = True Then
      frm_agentes.Visible = True
      frm_canales.Visible = False
   Else
      frm_agentes.Visible = False
      frm_canales.Visible = True
   End If
End Sub

Private Sub opt_tipo_filtrado_1_Validate(Cancel As Boolean)
   If opt_tipo_filtrado_1.Value = True Then
      frm_agentes.Visible = True
      frm_canales.Visible = False
   Else
      frm_agentes.Visible = False
      frm_canales.Visible = True
   End If

End Sub

Private Sub opt_tipo_filtrado_2_Click()
   If opt_tipo_filtrado_1.Value = True Then
      frm_agentes.Visible = True
      frm_canales.Visible = False
   Else
      frm_agentes.Visible = False
      frm_canales.Visible = True
   End If
End Sub

Private Sub txt_hora_KeyDown(KeyCode As Integer, Shift As Integer)
  ' On Error GoTo salir:
   If KeyCode = 116 Then
      lst_horas.Clear
      If IsDate(txt_fecha) Then
         var_fecha_fin_1 = CDate(txt_fecha) + 1
         var_dia = CStr(Day(CDate(txt_fecha)))
         var_mes = CStr(Month(CDate(txt_fecha)))
         var_año = CStr(Year(CDate(txt_fecha)))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
              
         var_dia = CStr(Day(var_fecha_fin_1))
         var_mes = CStr(Month(var_fecha_fin_1))
         var_año = CStr(Year(var_fecha_fin_1))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
     
         rs.Open "SELECT DISTINCT TIME_ORS_HORA FROM TB_enc_ORDEN_SURTIDO WHERE DTIM_ORS_FECHA_CARGA >= " + var_fecha_inicio + " AND DTIM_ORS_FECHA_CARGA <= " + var_fecha_fin, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  If Not IsNull(rs!time_ors_hora) Then
                     lst_horas.AddItem (rs!time_ors_hora)
                  End If
                  rs.MoveNext
            Wend
            frm_horas.Visible = True
            lst_horas.SetFocus
         Else
            MsgBox "No existen ordenes de surtido de la fecha especificada", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
          MsgBox "Fecha Incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   Exit Sub
SALIR:
   MsgBox "Formato de fecha incorrecto", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
End Sub
