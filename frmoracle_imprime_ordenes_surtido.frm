VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_imprime_ordenes_surtido 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creación de tablas para asignación de embarques"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Height          =   480
      Left            =   90
      TabIndex        =   26
      Top             =   2295
      Width           =   5685
      Begin VB.ComboBox cmb_dia 
         Height          =   315
         ItemData        =   "frmoracle_imprime_ordenes_surtido.frx":0000
         Left            =   2205
         List            =   "frmoracle_imprime_ordenes_surtido.frx":0016
         TabIndex        =   27
         Top             =   135
         Width           =   1950
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Dia de carga:"
         Height          =   195
         Left            =   1125
         TabIndex        =   28
         Top             =   180
         Width           =   960
      End
   End
   Begin VB.CommandButton cmd_eliminar_pedidos_intercompañias 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   765
      Picture         =   "frmoracle_imprime_ordenes_surtido.frx":004D
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Eliminar pedidos internos intercompañias"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir_pedido_tiendas 
      Caption         =   "OS Tiendas"
      Height          =   285
      Left            =   3030
      TabIndex        =   24
      Top             =   15
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmd_imprimir_divididas_numero 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1515
      Picture         =   "frmoracle_imprime_ordenes_surtido.frx":014F
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Pedidos divididos por periodos de números"
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_creacion_tablas 
      Caption         =   "Crear tablas"
      Height          =   315
      Left            =   1860
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmd_imprimir_divididas 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1185
      Picture         =   "frmoracle_imprime_ordenes_surtido.frx":0251
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Pedidos divididos por periodos de fechas"
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir_entradas 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmoracle_imprime_ordenes_surtido.frx":0353
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Imprimir entradas"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5535
      Picture         =   "frmoracle_imprime_ordenes_surtido.frx":0455
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmoracle_imprime_ordenes_surtido.frx":0A8F
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   0
      TabIndex        =   13
      Top             =   270
      Width           =   5940
   End
   Begin VB.Frame Frame3 
      Height          =   3345
      Left            =   90
      TabIndex        =   6
      Top             =   2760
      Width           =   5715
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1365
         Picture         =   "frmoracle_imprime_ordenes_surtido.frx":0B91
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   705
         Picture         =   "frmoracle_imprime_ordenes_surtido.frx":0DA7
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Marcar (Enter)"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1035
         Picture         =   "frmoracle_imprime_ordenes_surtido.frx":0FF1
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   45
         Picture         =   "frmoracle_imprime_ordenes_surtido.frx":10C3
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmoracle_imprime_ordenes_surtido.frx":11C5
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   135
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_agentes 
         Height          =   2790
         Left            =   45
         TabIndex        =   12
         Top             =   480
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   4921
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
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Por número"
      Height          =   870
      Left            =   90
      TabIndex        =   5
      Top             =   1440
      Width           =   5715
      Begin VB.TextBox txt_numero_fin 
         Height          =   300
         Left            =   3255
         TabIndex        =   11
         Top             =   375
         Width           =   1425
      End
      Begin VB.TextBox txt_numero_inicio 
         Height          =   300
         Left            =   1275
         TabIndex        =   3
         Top             =   375
         Width           =   1425
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2880
         TabIndex        =   10
         Top             =   435
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   675
         TabIndex        =   9
         Top             =   435
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Por fechas"
      Height          =   870
      Left            =   90
      TabIndex        =   4
      Top             =   480
      Width           =   5715
      Begin VB.TextBox txt_fecha_fin 
         Height          =   315
         Left            =   3180
         TabIndex        =   2
         Top             =   315
         Width           =   1425
      End
      Begin VB.TextBox txt_fecha_inicio 
         Height          =   315
         Left            =   1275
         TabIndex        =   1
         Top             =   315
         Width           =   1425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2895
         TabIndex        =   8
         Top             =   375
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   675
         TabIndex        =   7
         Top             =   375
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmoracle_imprime_ordenes_surtido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_consecutivo_general As Double
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub crea_tablas()
        'rsaux2.Open "select * from tb_temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo_general), cnn, adOpenDynamic, adLockOptimistic
        'While Not rsaux2.EOF
        '      rsaux3.Open "select * from xxvia_tb_pedidos_divididos  where SOURCE_HEADER_NUMBER = " + rsaux2!source_header_number + " and DELIVERY_ID = " + CStr(rsaux2!delivery_id) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!DELIVERY_DETAIL_ID) + " and ORGANIZATION_ID = " + CStr(rsaux2!organization_id) + " and DELIVERY_LINE_ID = " + CStr(rsaux2!delivery_line_id) + " and INVENTORY_ITEM_ID = " + CStr(rsaux2!INVENTORY_ITEM_ID) + " and SOURCE_LINE_NUMBER = '" + rsaux2!source_line_number + "' and lote = " + CStr(rsaux2!lote), cnnoracle_4, adOpenDynamic, adLockOptimistic
        '      If rsaux3.EOF Then
        '         var_cadena = "INSERT INTO XXVIA_TB_PEDIDOS_DIVIDIDOS (SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,SEGMENT1,LOTE,CUSTOMER_NAME,COLLECTOR_ID,NAME,DATE_REQUESTED,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,CUST_ACCOUNT_ID,SOURCE_HEADER_TYPE_NAME,SOURCE_DOCUMENT_ID,SITE_USE_ID,linea,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,Embarque,ESTACION,ORDEN_CARGA) VALUES "
        '         var_cadena = var_cadena + "(" + CStr(rsaux2!source_header_number) + "," + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!DELIVERY_DETAIL_ID) + "," + CStr(rsaux2!organization_id) + ",'" + rsaux2!subinventory + "'," + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!INVENTORY_ITEM_ID) + ",'" + rsaux2!ITEM_DESCRIPTION + "','" + rsaux2!source_line_number + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "','" + rsaux2!segment1 + "'," + CStr(rsaux2!lote) + ",'" + rsaux2!CUSTOMER_NAME + "'," + CStr(rsaux2!collector_id) + ",'" + rsaux2!Name + "','" + CStr(rsaux2!DATE_REQUESTED) + "'," + CStr(rsaux2!ESTABLECIMIENTO) + ",'" + rsaux2!NOMBRE_ESTABLECIMIENTO + "'," + CStr(rsaux2!CUST_ACCOUNT_ID) + ",'" + rsaux2!source_header_type_name + "','" + CStr(IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id)) + "'," + CStr(rsaux2!site_use_id) + ",'" + IIf(IsNull(rsaux2!linea), "", rsaux2!linea) + "'," + CStr(IIf(IsNull(rsaux2!ruta), 0, rsaux2!ruta))
        '         var_cadena = var_cadena + ",'" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ",'" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_CARGA), 0, rsaux2!ORDEN_CARGA)) + ")"
        '         'MsgBox var_cadena
        '         rsaux4.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
        '      End If
        '      rsaux3.Close
        '      rsaux2.MoveNext
        'Wend
        'rsaux2.Close

End Sub

Private Sub cmd_creacion_tablas_Click()
   Dim var_consecutivo As Integer
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_cadena_agentes = ""
   For var_j = 1 To lv_agentes.ListItems.Count
       lv_agentes.ListItems.Item(var_j).Selected = True
       If lv_agentes.selectedItem.SubItems(2) = "*" Then
          If var_cadena_agentes = "" Then
             var_cadena_agentes = lv_agentes.selectedItem
          Else
             var_cadena_agentes = var_cadena_agentes + "," + Me.lv_agentes.selectedItem
          End If
       End If
   Next var_j
   If IsDate(Me.txt_fecha_inicio) Then
      If IsDate(Me.txt_fecha_fin) Then
         If var_cadena_agentes <> "" Then
            var_cadena = "SELECT DISTINCT a.SOURCE_HEADER_NUMBER AS PEDIDO from WSH_DLVB_DLVY_V B, hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, XXVIA_VW_AGENTES D Where a.delivery_id = B.delivery_id AND A.delivery_detail_id = B.delivery_detail_id AND HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID and TRUNC(B.LAST_UPDATE_DATE) >= to_date('" + Me.txt_fecha_inicio + "','DD-MM-YYYY') AND TRUNC(B.LAST_UPDATE_DATE) < TO_DATE('" + CStr(CDate(Me.txt_fecha_fin) + 1) + "','DD-MM-YYYY') AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.COLLECTOR_ID IN (" + var_cadena_agentes + ") "
            var_cadena = var_cadena + " AND HCAS.CUST_ACCOUNT_ID = D.CUST_ACCOUNT_ID AND OHA.SHIP_FROM_ORG_ID = " + var_unidad_organizacional + " AND A.RELEASED_STATUS = 'Y' ORDER BY a.SOURCE_HEADER_NUMBER"
            Text1 = var_cadena
            rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_Cadena_pedidos = ""
            While Not rsaux2.EOF
                  If Trim(var_Cadena_pedidos) = "" Then
                     var_Cadena_pedidos = CStr(rsaux2!pedido)
                  Else
                     var_Cadena_pedidos = var_Cadena_pedidos + "," + CStr(rsaux2!pedido)
                  End If
                  rsaux2.MoveNext
            Wend
            rsaux2.Close
            If var_Cadena_pedidos <> "" Then
               'MsgBox VAR_CADENA_PEDIDOS
               var_cadena = "SELECT CAT.LINEA, HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id, a.LAST_UPDATE_DATE,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,c.description,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, a.source_header_type_name, oha.source_document_id, C.ATTRIBUTE2, oha.attribute8, oha.attribute9, j.NAME as nombre_ruta, j.salesrep_id as clave_ruta from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, xxvia_vw_articulos_cat cat, XXVIA_VENDEDORES j Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) IN (" + var_Cadena_pedidos + ") "
               var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID "
               var_cadena = var_cadena + " AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y' AND A.inventory_item_id  = cat.item_id AND A.ORGANIZATION_ID = Cat.organization_id and oha.salesrep_id = j.salesrep_id "
               rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  cnn.BeginTrans
                  rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM tb_Temp_oracle_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
                  Else
                     var_consecutivo = 1
                  End If
                  rsaux.Close
                  rsaux1.Open "insert into tb_Temp_oracle_orden_surtido(inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  cnn.CommitTrans
                  While Not rs.EOF
                        var_establecimiento = rs!SHIP_TO_ORG_ID
                        rsaux.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'SHIP_TO' AND csu.LOCATION = ps.party_site_number and csu.site_use_id = " + CStr(var_establecimiento), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux!vcha_esb_nombre), "", rsaux!vcha_esb_nombre)
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
                        var_cadena = "insert into tb_Temp_oracle_orden_surtido(INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, DELIVERY_ID, DELIVERY_DETAIL_ID, ORGANIZATION_ID, SUBINVENTORY, DELIVERY_LINE_ID, INVENTORY_ITEM_ID, ITEM_DESCRIPTION, SOURCE_LINE_NUMBER, src_requested_quantity, RELEASED_STATUS, CUSTOMER_NAME, SEGMENT1, COLLECTOR_ID, NAME, date_requested, UBICACION, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, SITE_USE_ID, attribute8, attribute9, LINEA, RUTA, NOMBRE_RUTA)  values "
                        var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rs!source_header_number), "", rs!source_header_number) + "', " + CStr(IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)) + ", " + CStr(IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)) + ", " + CStr(IIf(IsNull(rs!organization_id), 0, rs!organization_id)) + ", '" + IIf(IsNull(rs!subinventory), "", rs!subinventory) + "', " + CStr(IIf(IsNull(rs!delivery_line_id), 0, rs!delivery_line_id)) + ", " + CStr(IIf(IsNull(rs!inventory_item_id), "", rs!inventory_item_id)) + ", '" + IIf(IsNull(rs!Description), "", rs!Description) + "', '" + IIf(IsNull(rs!SOURCE_LINE_NUMBER), "", rs!SOURCE_LINE_NUMBER) + "', " + CStr(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)) + ", '" + IIf(IsNull(rs!released_status), "", rs!released_status) + "', '" + IIf(IsNull(rs!customer_name), "", rs!customer_name) + "', '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'"
                        var_cadena = var_cadena + ", 0,''," + var_fecha + ",'" + IIf(IsNull(rs!attribute2), "", rs!attribute2) + "','" + CStr(var_establecimiento) + "','" + VAR_NOMBRE_ESTABLECIMIENTO + "'," + CStr(rs!CUST_ACCOUNT_ID) + ",'" + rs!source_header_type_name + "','" + CStr(IIf(IsNull(rs!source_document_id), "", rs!source_document_id)) + "'," + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)) + ",'" + IIf(IsNull(rs!attribute8), "", rs!attribute8) + "','" + IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9) + "','" + IIf(IsNull(rs!Linea), "", rs!Linea) + "','" + CStr(rs!CLAVE_RUTA) + "','" + rs!nombre_ruta + "')"
                        rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        rs.MoveNext
                  Wend
                  rsaux1.Open "delete from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number is null", cnn, adOpenDynamic, adLockOptimistic
                  If rsaux1.State = 1 Then
                     rsaux1.Close
                  End If
                  rsaux1.Open "select DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_Temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " group by DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA"
                  While Not rsaux1.EOF
                        If rsaux1!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rsaux1!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                           If var_pedido_tienda = 0 Then
                              If rsaux2.State = 1 Then
                                 rsaux2.Close
                              End If
                              rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux1!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           Else
                              rsaux2.Open "select a.attribute8, B.description from oe_order_headers_all a, MTL_SECONDARY_INVENTORIES b where order_number = " + CStr(rsaux1!source_header_number) + " and a.attribute8 = b.secondary_inventory_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           End If
                        Else
                           rsaux6.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           VAR_PROVEEDOR = rsaux6!collector_id
                           VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux6!Name), "", rsaux6!Name)
                           rsaux6.Close
                        End If
                        var_año_str = CStr(Year(rsaux1!DATE_REQUESTED))
                        If Len(var_año_str) < 2 Then
                           var_año_str = "20" + var_año_str
                        End If
                        var_mes_str = CStr(Month(rsaux1!DATE_REQUESTED))
                        If Len(var_mes_str) < 2 Then
                           var_mes_str = "0" + var_mes_str
                        End If
                        var_dia_str = CStr(Day(rsaux1!DATE_REQUESTED))
                        If Len(var_dia_str) < 2 Then
                           var_dia_str = "0" + var_dia_str
                        End If
                        var_hora_str = CStr(Hour(rsaux1!DATE_REQUESTED))
                        If Len(var_hora_str) < 2 Then
                           var_hora_str = "0" + var_hora_str
                        End If
                        VAR_MINUTO_STR = CStr(Minute(rsaux1!DATE_REQUESTED))
                        If Len(VAR_MINUTO_STR) < 2 Then
                           VAR_MINUTO_STR = "0" + VAR_MINUTO_STR
                        End If
                        VAR_SEGUNDO_STR = CStr(Second(rsaux1!DATE_REQUESTED))
                        If Len(VAR_SEGUNDO_STR) < 2 Then
                           VAR_SEGUNDO_STR = "0" + VAR_SEGUNDO_STR
                        End If
                        var_fecha_pedido = var_año_str + "-" + var_mes_str + "-" + var_dia_str + " " + var_hora_str + ":" + VAR_MINUTO_STR + ":" + VAR_SEGUNDO_STR
                        rsaux6.Open "update tb_Temp_oracle_orden_surtido set COLLECTOR_ID = " + CStr(VAR_PROVEEDOR) + ", NAME = '" + VAR_NOMBRE_PROVEEDOR + "' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and source_header_number = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        rsaux6.Open "SELECT * FROM XXVIA_TB_ORDENES_GRAFICA WHERE ORGANIZACION = " + var_unidad_organizacional + " AND PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           rsaux5.Open "INSERT INTO XXVIA_TB_ORDENES_GRAFICA (ORGANIZACION, PEDIDO, PROVEEDOR_ID, NOMBRE_PROVEEDOR, CLIENTE, ESTABLECIMIENTO, CANTIDAD, FECHA, TIPO_PEDIDO, CANTIDAD_LEIDA, FECHA_PEDIDO, RUTA, NOMBRE_RUTA) VALUES (" + var_unidad_organizacional + ", " + CStr(rsaux1!source_header_number) + "," + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "'," + CStr(rsaux1!cantidad) + ", '" + CStr(rsaux1!DATE_REQUESTED) + "','" + rsaux1!source_header_type_name + "',0, TO_DATE('" + var_fecha_pedido + "','YYYY-MM-DD HH24:MI:SS'),'" + rsaux1!ruta + "', '" + rsaux1!nombre_ruta + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET RUTA = '" + rsaux1!ruta + "', NOMBRE_RUTA = '" + rsaux1!nombre_ruta + "' WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        rsaux6.Open "select * from tb_oracle_pedidos_asignados_embarques where pedido = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           rsaux5.Open "insert into tb_oracle_pedidos_asignados_embarques (AGENTE, NOMBRE_AGENTE, PEDIDO, CLIENTE, PIEZAS, embarque, dia,  mes, AÑO, ORGANIZACION) values (" + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "'," + CStr(rsaux1!source_header_number) + ",'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "', " + CStr(rsaux1!cantidad) + ",0," + var_dia_str + "," + var_mes_str + "," + var_año_str + "," + CStr(var_unidad_organizacional) + ")", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  MsgBox "Termino el proceso de creación de tablas", vbOKOnly, "ATENCION"
               Else
                  MsgBox "No existen pedidos por cargar", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No existen pedidos por cargar", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se a seleccionado ningún agente", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha inicial incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_eliminar_pedidos_intercompañias_Click()
   rs.Open "DELETE from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where NOMBRE_AGENTE= 'INTERCOMPAÑIAS' and cliente like '%VIANNEY TEXTIL HOGAR SA DE CV%'  and isnull(embarque,0)  = 0 ", cnn, adOpenDynamic, adLockOptimistic
   MsgBox "Se han eliminado los pedidos", vbOKOnly, "ATENCION"
   
End Sub

Private Sub cmd_imprimir_Click()
   Dim var_consecutivo As Integer
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_cadena_agentes = ""
   For var_j = 1 To lv_agentes.ListItems.Count
       lv_agentes.ListItems.Item(var_j).Selected = True
       If lv_agentes.selectedItem.SubItems(2) = "*" Then
          If var_cadena_agentes = "" Then
             var_cadena_agentes = lv_agentes.selectedItem
          Else
             var_cadena_agentes = var_cadena_agentes + "," + Me.lv_agentes.selectedItem
          End If
       End If
   Next var_j
   If IsDate(Me.txt_fecha_inicio) Then
      If IsDate(Me.txt_fecha_fin) Then
         If var_cadena_agentes <> "" Then
            var_cadena = "SELECT DISTINCT a.SOURCE_HEADER_NUMBER AS PEDIDO from WSH_DLVB_DLVY_V B, hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, XXVIA_VW_AGENTES D Where a.delivery_id = B.delivery_id AND A.delivery_detail_id = B.delivery_detail_id AND HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID and TRUNC(B.LAST_UPDATE_DATE) >= to_date('" + Me.txt_fecha_inicio + "','DD-MM-YYYY') AND TRUNC(B.LAST_UPDATE_DATE) < TO_DATE('" + CStr(CDate(Me.txt_fecha_fin) + 1) + "','DD-MM-YYYY') AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.COLLECTOR_ID IN (" + var_cadena_agentes + ") "
            var_cadena = var_cadena + " AND HCAS.CUST_ACCOUNT_ID = D.CUST_ACCOUNT_ID AND nvl(OHA.SHIP_FROM_ORG_ID,93) = " + var_unidad_organizacional + " AND A.RELEASED_STATUS = 'Y' ORDER BY a.SOURCE_HEADER_NUMBER"
            'var_cadena = "SELECT distinct source_header_number as pedido FROM WSH_dELIVERABLES_V WHERE SOURCE_HEADER_NUMBER IN (662960, 662962, 662964, 662965, 662966, 662967, 662963) ORDER BY SOURCE_HEADER_NUMBER DESC"
            rsaux2.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            'var_cadena = "select DISTINCT SOURCE_HEADER_NUMBER AS PEDIDO from wsh_Deliverables_v where source_header_number in (665025, 665026, 665033, 665024, 665027, 665028, 665034, 665029, 665030, 665035, 665031, 665032, 665036, 665037, 665038, 665039, 665040, 665039, 665040, 665049, 665041, 665042, 664053,665050, 665044, 665045, 665046, 665051, 665047, 665048, 665052, 665024, 665025, 665026, 665027, 665028, 665029, 665030, 665031, 665032, 665033, 665034,665035, 665036,"
            'var_cadena = var_cadena + "665037, 665038, 665039, 665040, 665041, 665042, 665043, 665044, 665045, 665046, 665047, 665048)"
            
            rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_Cadena_pedidos = ""
            While Not rsaux2.EOF
                  If Trim(var_Cadena_pedidos) = "" Then
                     var_Cadena_pedidos = CStr(rsaux2!pedido)
                  Else
                     var_Cadena_pedidos = var_Cadena_pedidos + "," + CStr(rsaux2!pedido)
                  End If
                  rsaux2.MoveNext
            Wend
            rsaux2.Close
            'var_cadena_pedidos = "320903"

            If var_Cadena_pedidos <> "" Then
               'MsgBox VAR_CADENA_PEDIDOS
               var_cadena = "SELECT to_char(a.LAST_UPDATE_DATE,'day') DIA_SEMANA, CAT.LINEA, HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, "
               var_cadena = var_cadena + " HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id, a.LAST_UPDATE_DATE,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,c.description,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, a.source_header_type_name, oha.source_document_id, C.ATTRIBUTE2, oha.attribute8, oha.attribute9, j.NAME as nombre_ruta, j.salesrep_id as clave_ruta from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, xxvia_vw_articulos_cat cat, XXVIA_VENDEDORES j Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) IN (" + var_Cadena_pedidos + ") "
               var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID "
               var_cadena = var_cadena + " AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y' AND A.inventory_item_id  = cat.item_id AND A.ORGANIZATION_ID = Cat.organization_id and oha.salesrep_id = j.salesrep_id "
               Text1 = var_cadena
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
                        If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Then
                           rsaux1.Open "SELECT * FROM TB_ORACLE_ARTICULOS_MOTOR_LOGISTICO WHERE CODIGO = '" + rs!SEGMENT1 + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux1.EOF Then
                              strconsulta = "SELECT secondary_inventory_name, A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!source_document_id)
                                   .Parameters.Append parametro
                              End With
                              Set rsaux8 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                              If rsaux8.EOF Then
                                 var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                              Else
                                 var_almacen = rsaux8!secondary_inventory_name
                                 rsaux9.Open "SELECT * FROM TB_ORACLE_UBICACIONES_MOTOR_LOGISTICO WHERE CLAVE = '" + var_almacen + "' AND CODIGO = '" + rs!SEGMENT1 + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux9.EOF Then
                                    var_ubicacion = ""
                                 If Me.cmb_dia.Text = "Lunes" Then
                                    var_ubicacion = rsaux9!ubicacion_1
                                 End If
                                 If Me.cmb_dia.Text = "Martes" Then
                                    var_ubicacion = rsaux9!ubicacion_2
                                 End If
                                 If Me.cmb_dia.Text = "Miercoles" Then
                                    var_ubicacion = rsaux9!ubicacion_3
                                 End If
                                 If Me.cmb_dia.Text = "Jueves" Then
                                    var_ubicacion = rsaux9!ubicacion_4
                                 End If
                                 If Me.cmb_dia.Text = "Viernes" Then
                                    var_ubicacion = rsaux9!ubicacion_5
                                 End If
                                 If Me.cmb_dia.Text = "Sabado" Then
                                    var_ubicacion = rsaux9!ubicacion_6
                                 End If
                                 If IIf(IsNull(var_ubicacion), "", var_ubicacion) = var_ubicacion = "" Then
                                    var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                                 End If
                                    If var_ubicacion = "" Then
                                       var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                                    End If
                                 Else
                                    var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                                 End If
                                 rsaux9.Close
                              End If
                              rsaux8.Close
                           Else
                              var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                           End If
                           rsaux1.Close
                        Else
                           var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                        End If
                        var_establecimiento = rs!SHIP_TO_ORG_ID
                        rsaux.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'SHIP_TO' AND csu.LOCATION = ps.party_site_number and csu.site_use_id = " + CStr(var_establecimiento), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux!vcha_esb_nombre), "", rsaux!vcha_esb_nombre)
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
'''''direccion

                        
                     


''''' fin direccion
                        
                        var_cadena = "insert into tb_temp_oracle_orden_surtido (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, DELIVERY_ID, DELIVERY_DETAIL_ID, ORGANIZATION_ID, SUBINVENTORY, DELIVERY_LINE_ID, INVENTORY_ITEM_ID, ITEM_DESCRIPTION, SOURCE_LINE_NUMBER, src_requested_quantity, RELEASED_STATUS, CUSTOMER_NAME, SEGMENT1, COLLECTOR_ID, NAME, date_requested, UBICACION, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, SITE_USE_ID, attribute8, attribute9, LINEA, RUTA, NOMBRE_RUTA)  values "
                        var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rs!source_header_number), "", rs!source_header_number) + "', " + CStr(IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)) + ", " + CStr(IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)) + ", " + CStr(IIf(IsNull(rs!organization_id), 0, rs!organization_id)) + ", '" + IIf(IsNull(rs!subinventory), "", rs!subinventory) + "', " + CStr(IIf(IsNull(rs!delivery_line_id), 0, rs!delivery_line_id)) + ", " + CStr(IIf(IsNull(rs!inventory_item_id), "", rs!inventory_item_id)) + ", '" + IIf(IsNull(rs!Description), "", rs!Description) + "', '" + IIf(IsNull(rs!SOURCE_LINE_NUMBER), "", rs!SOURCE_LINE_NUMBER) + "', " + CStr(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)) + ", '" + IIf(IsNull(rs!released_status), "", rs!released_status) + "', '" + Replace(IIf(IsNull(rs!customer_name), "", rs!customer_name), "'", " ") + "', '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'"
                        'var_cadena = var_cadena + ", 0,''," + var_fecha + ",'" + IIf(IsNull(rs!attribute2), "", rs!attribute2) + "','" + CStr(VAR_ESTABLECIMIENTO) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "'," + CStr(rs!CUST_ACCOUNT_ID) + ",'" + rs!source_header_type_name + "','" + CStr(IIf(IsNull(rs!source_document_id), "", rs!source_document_id)) + "'," + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)) + ",'" + IIf(IsNull(rs!attribute8), "", rs!attribute8) + "','" + IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9) + "','" + IIf(IsNull(rs!linea), "", rs!linea) + "','" + CStr(rs!CLAVE_RUTA) + "','" + IIf(IsNull(rs!nombre_ruta), "", rs!nombre_ruta) + "')"
                        var_cadena = var_cadena + ", 0,''," + var_fecha + ",'" + var_ubicacion + "','" + CStr(var_establecimiento) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "'," + CStr(rs!CUST_ACCOUNT_ID) + ",'" + rs!source_header_type_name + "','" + CStr(IIf(IsNull(rs!source_document_id), "", rs!source_document_id)) + "'," + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)) + ",'" + IIf(IsNull(rs!attribute8), "", rs!attribute8) + "','" + IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9) + "','" + IIf(IsNull(rs!Linea), "", rs!Linea) + "','" + CStr(rs!CLAVE_RUTA) + "','" + IIf(IsNull(rs!nombre_ruta), "", rs!nombre_ruta) + "')"
                        
                        
                        
                        
                        'var_cadena = "insert into tb_temp_oracle_orden_surtido (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, DELIVERY_ID, DELIVERY_DETAIL_ID, ORGANIZATION_ID, SUBINVENTORY, DELIVERY_LINE_ID, INVENTORY_ITEM_ID, ITEM_DESCRIPTION, SOURCE_LINE_NUMBER, src_requested_quantity, RELEASED_STATUS, CUSTOMER_NAME, SEGMENT1, COLLECTOR_ID, NAME, date_requested, ubicacion, pais, municipio, estado, ciudad, colonia, direccion, cp, PAQUETERIA, telefono, nombre_establecimiento, COMENTARIO)  values "
                        'var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rs!source_header_number), "", rs!source_header_number) + "', " + CStr(IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)) + ", " + CStr(IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)) + ", " + CStr(IIf(IsNull(rs!organization_id), 0, rs!organization_id)) + ", '" + IIf(IsNull(rs!subinventory), "", rs!subinventory) + "', " + CStr(IIf(IsNull(rs!delivery_line_id), 0, rs!delivery_line_id)) + ", " + CStr(IIf(IsNull(rs!inventory_item_id), "", rs!inventory_item_id)) + ", '" + IIf(IsNull(rs!Description), "", rs!Description) + "', '" + IIf(IsNull(rs!SOURCE_LINE_NUMBER), "", rs!SOURCE_LINE_NUMBER) + "', " + CStr(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)) + ", '" + IIf(IsNull(rs!released_status), "", rs!released_status) + "', '" + IIf(IsNull(rs!customer_name), "", rs!customer_name) + "', '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'"
                        'var_cadena = var_cadena + ", " + CStr(IIf(IsNull(rsaux!collector_id), 0, rsaux!collector_id)) + ",'" + IIf(IsNull(rsaux!Name), "", rsaux!Name) + "'," + var_fecha + ",'" + IIf(IsNull(rs!attribute2), "", rs!attribute2) + "','" + var_pais + "','" + VAR_MUNICIPIO + "','" + var_estado + "', '" + var_ciudad + "','" + VAR_COLONIA + "','" + VAR_DIRECCION + "','" + VAR_CP + "','" + var_paqueteria + "','" + var_telefono + "','" + var_nombre + "','" + VAR_COMENTARIOS + "')"
                        
                        
                        
                        
                        
                        rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        rs.MoveNext
                  Wend
                  rsaux1.Open "delete from tb_temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number is null", cnn, adOpenDynamic, adLockOptimistic
                  If rsaux1.State = 1 Then
                     rsaux1.Close
                  End If
                  rsaux1.Open "select min(DATE_REQUESTED) as DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " group by source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA"
                  While Not rsaux1.EOF
                        If rsaux1!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rsaux1!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                           If var_pedido_tienda = 0 Then
                              If rsaux2.State = 1 Then
                                 rsaux2.Close
                              End If
                              rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux1!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 If rsaux4.State = 1 Then
                                    rsaux4.Close
                                 End If
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           Else
                              rsaux2.Open "select a.attribute8, B.description from oe_order_headers_all a, MTL_SECONDARY_INVENTORIES b where order_number = " + CStr(rsaux1!source_header_number) + " and a.attribute8 = b.secondary_inventory_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           End If
                        Else
                           rsaux6.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux6.EOF Then
                              VAR_PROVEEDOR = rsaux6!collector_id
                              VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux6!Name), "", rsaux6!Name)
                           Else
                              VAR_PROVEEDOR = ""
                              VAR_NOMBRE_PROVEEDOR = ""
                           End If
                           rsaux6.Close
                        End If
                        var_año_str = CStr(Year(rsaux1!DATE_REQUESTED))
                        If Len(var_año_str) < 2 Then
                           var_año_str = "20" + var_año_str
                        End If
                        var_mes_str = CStr(Month(rsaux1!DATE_REQUESTED))
                        If Len(var_mes_str) < 2 Then
                           var_mes_str = "0" + var_mes_str
                        End If
                        var_dia_str = CStr(Day(rsaux1!DATE_REQUESTED))
                        If Len(var_dia_str) < 2 Then
                           var_dia_str = "0" + var_dia_str
                        End If
                        var_hora_str = CStr(Hour(rsaux1!DATE_REQUESTED))
                        If Len(var_hora_str) < 2 Then
                           var_hora_str = "0" + var_hora_str
                        End If
                        VAR_MINUTO_STR = CStr(Minute(rsaux1!DATE_REQUESTED))
                        If Len(VAR_MINUTO_STR) < 2 Then
                           VAR_MINUTO_STR = "0" + VAR_MINUTO_STR
                        End If
                        VAR_SEGUNDO_STR = CStr(Second(rsaux1!DATE_REQUESTED))
                        If Len(VAR_SEGUNDO_STR) < 2 Then
                           VAR_SEGUNDO_STR = "0" + VAR_SEGUNDO_STR
                        End If
                        var_fecha_pedido = var_año_str + "-" + var_mes_str + "-" + var_dia_str + " " + var_hora_str + ":" + VAR_MINUTO_STR + ":" + VAR_SEGUNDO_STR
                        If VAR_PROVEEDOR = "" Then
                           VAR_PROVEEDOR = "0"
                        End If
                        
                        
'''''direccion
                        rsaux6.Open "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.invoice_TO_ORG_ID AND oha.order_number = '" + IIf(IsNull(CStr(rsaux1!source_header_number)), "", CStr(rsaux1!source_header_number)) + "' AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux6.EOF Then
                           rsaux5.Open "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = '" + IIf(IsNull(CStr(rsaux1!source_header_number)), "", CStr(rsaux1!source_header_number)) + "' AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux5.EOF Then
                              var_nombre = IIf(IsNull(rsaux5!customer_name), "", rsaux5!customer_name)
                              var_tel = IIf(IsNull(rsaux5!tel), 0, rsaux5!tel)
                              VAR_DIRECCION = IIf(IsNull(rsaux5!calle), "", rsaux5!calle) + " " + IIf(IsNull(rsaux5!numero), "", rsaux5!numero)
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
                              VAR_DIRECCION = IIf(IsNull(rsaux6!calle), "", rsaux6!calle) + " " + IIf(IsNull(rsaux6!numero), "", rsaux6!numero)
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
                     
                     
                        var_cadena = "SELECT HCAS.CUST_ACCOUNT_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, c.attribute2, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) "
                        var_cadena = var_cadena + " Between " + CStr(rsaux1!source_header_number) + " And " + CStr(rsaux1!source_header_number) + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID  AND released_status = 'Y'"
                        rsaux.Open "select shipping_method_code, packing_instructions from oe_order_headers_all where order_number = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        var_paqueteria = ""
                        If Not rsaux.EOF Then
                           VAR_COMENTARIOS = IIf(IsNull(rsaux!packing_instructions), "", rsaux!packing_instructions)
                           var_tipo_metodo = IIf(IsNull(rsaux(0).Value), "", rsaux(0).Value)
                           If var_tipo_metodo <> "" Then
                              rsaux2.Open "SELECT description FROM fnd_lookup_values where lookup_type = 'SHIP_METHOD' and lookup_code = '" + var_tipo_metodo + "' AND LANGUAGE = 'ESA'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 var_paqueteria = IIf(IsNull(rsaux2(0).Value), "", rsaux2(0).Value)
                              End If
                              rsaux2.Close
                           End If
                        End If
                        rsaux.Close
                        
                        
'''''fin direccion
                        
                        
                        
                        
                        
                        rsaux6.Open "update tb_temp_oracle_orden_surtido set COLLECTOR_ID = " + CStr(IIf(IsNull(VAR_PROVEEDOR), 0, VAR_PROVEEDOR)) + ", NAME = '" + VAR_NOMBRE_PROVEEDOR + "', pais = '" + var_pais + "',estado = '" + var_estado + "',municipio = '" + VAR_MUNICIPIO + "', ciudad = '" + var_ciudad + "', colonia = '" + VAR_COLONIA + "', direccion = '" + VAR_DIRECCION + "', cp = '" + VAR_CP + "', paqueteria = '" + var_paqueteria + "', comentario = '" + var_comentario + "' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and source_header_number = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        ' esto se quita por el problema del servidor
                        x = 1
                        If x = 1 Then
                        rsaux6.Open "SELECT * FROM XXVIA_TB_ORDENES_GRAFICA WHERE ORGANIZACION = " + var_unidad_organizacional + " AND PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           rsaux5.Open "INSERT INTO XXVIA_TB_ORDENES_GRAFICA (ORGANIZACION, PEDIDO, PROVEEDOR_ID, NOMBRE_PROVEEDOR, CLIENTE, ESTABLECIMIENTO, CANTIDAD, FECHA, TIPO_PEDIDO, CANTIDAD_LEIDA, FECHA_PEDIDO, RUTA, NOMBRE_RUTA) VALUES (" + var_unidad_organizacional + ", " + CStr(rsaux1!source_header_number) + "," + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "'," + CStr(rsaux1!cantidad) + ", '" + CStr(rsaux1!DATE_REQUESTED) + "','" + rsaux1!source_header_type_name + "',0, TO_DATE('" + var_fecha_pedido + "','YYYY-MM-DD HH24:MI:SS'),'" + rsaux1!ruta + "', '" + rsaux1!nombre_ruta + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET RUTA = '" + rsaux1!ruta + "', NOMBRE_RUTA = '" + rsaux1!nombre_ruta + "' WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        'MsgBox cnn.ConnectionString
                        rsaux6.Open "select * from tb_oracle_pedidos_asignados_embarques where pedido = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           rsaux5.Open "insert into tb_oracle_pedidos_asignados_embarques (AGENTE, NOMBRE_AGENTE, PEDIDO, CLIENTE, PIEZAS, embarque, dia,  mes, AÑO, ORGANIZACION) values (" + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "'," + CStr(rsaux1!source_header_number) + ",'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "', " + CStr(rsaux1!cantidad) + ",0," + var_dia_str + "," + var_mes_str + "," + var_año_str + "," + CStr(var_unidad_organizacional) + ")", cnn, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux5.Open "UPDATE TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES SET PIEZAS = " + CStr(rsaux1!cantidad) + " WHERE PEDIDO = '" + CStr(rsaux1!source_header_number) + "'", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        End If
                        'hasta aqui
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
               
                  ' se quita por el problema del servidor
                  x = 1
                  If x = 1 Then
                  If rsaux1.State = 1 Then
                     rsaux1.Close
                  End If
                  rsaux1.Open "select source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " AND (LINEA <> 'CATALOGOS' OR LINEA IS NULL) group by source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME"
                  While Not rsaux1.EOF
                        rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET CANTIDAD_SIN_CATALOGOS = " + CStr(IIf(IsNull(rsaux1!cantidad), 0, rsaux1!cantidad)) + " WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  End If
                  'hasta aqui
                  rsaux.Open "select distinct source_header_number from tb_temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER is not null", cnn, adOpenDynamic, adLockOptimistic
                  var_si = 0
                  var_si = MsgBox("¿Se imprimiran las ordenes de surtido?", vbYesNo, "ATENCION")
                  
                  If Not rsaux.EOF Then
                     While Not rsaux.EOF
                           x = 1
                           If x = 1 Then
                              rsaux5.Open "SELECT * FROM tb_oracle_pedidos_asignados_embarques WHERE PEDIDO = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux5.EOF Then
                                 rsaux3.Open "UPDATE tb_temp_oracle_orden_surtido SET EMBARQUE = " + CStr(IIf(IsNull(rsaux5!Embarque), 0, rsaux5!Embarque)) + ", ORDEN_SURTIDO = " + CStr(IIf(IsNull(rsaux5!orden_pedido), 0, rsaux5!orden_pedido)) + ", ESTACION = '" + CStr(IIf(IsNull(rsaux5!estacion), 0, rsaux5!estacion)) + "' WHERE SOURCE_HEADER_NUMBER = " + CStr(rsaux(0).Value) + " AND inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux5.Close
                              
                              If var_si = 6 Then
                                 x = 0
                                 If x = 1 Then
                                    Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_ft.rpt")
                                    reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.SOURCE_HEADER_NUMBER} = '" + rsaux(0).Value + "' and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                    'Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido.rpt")
                                    'reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.SOURCE_HEADER_NUMBER} = '" + rsaux(0).Value + "' and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                    For ntablas = 1 To reporte.Database.Tables.Count
                                        reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                    Next ntablas
                                    reporte.PrintOut False
                                    Set reporte = Nothing
                                 Else
                                    Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_ft.rpt")
                                    reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.SOURCE_HEADER_NUMBER} = '" + rsaux(0).Value + "' and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                    frmvistasprevias.cr.ReportSource = reporte
                                    For ntablas = 1 To reporte.Database.Tables.Count
                                        reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                    Next ntablas
                                    frmvistasprevias.cr.ViewReport
                                    frmvistasprevias.Caption = "Ordenes de surtido historica"
                                    frmvistasprevias.Show 1
                                    Set reporte = Nothing
                                 
                                 End If
                              End If
                           End If
                           rsaux.MoveNext
                     Wend
                  End If
                  rsaux.Close
               Else
                  MsgBox "No existen ordenes de surtido", vbOKOnly, "ATENCION"
               End If
               rs.Close
            Else
               MsgBox "No existen ordenes de surtido", vbOKOnly, "ATENCION"
            End If
            rs.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "No se seleccionaron agentes", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Rango de fechas incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      If IsNumeric(Me.txt_numero_inicio) Then
         If IsNumeric(Me.txt_numero_fin) Then
         
         rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         
         var_cadena = " SELECT to_char(a.LAST_UPDATE_DATE,'day') DIA_SEMANA, HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, a.source_header_type_name, oha.source_document_id, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.LAST_UPDATE_DATE,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,c.description,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, C.ATTRIBUTE2, j.NAME as nombre_ruta, j.salesrep_id as clave_ruta  from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C,  xxvia_vw_articulos_cat cat, XXVIA_VENDEDORES j  "
         var_cadena = var_cadena + " Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND  to_number(source_header_number)  BETWEEN " + CStr(Me.txt_numero_inicio) + " AND " + CStr(Me.txt_numero_fin) + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID "
         var_cadena = var_cadena + " AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y' and A.inventory_item_id  = cat.item_id AND A.ORGANIZATION_ID = Cat.ORGANIZATION_ID  and oha.salesrep_id = j.salesrep_id "
         
         
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
                  var_establecimiento = rs!SHIP_TO_ORG_ID
                  rsaux.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'SHIP_TO' AND csu.LOCATION = ps.party_site_number and csu.site_use_id = " + CStr(var_establecimiento), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux!vcha_esb_nombre), "", rsaux!vcha_esb_nombre)
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
                           VAR_PROVEEDOR = rsaux4!collector_id
                           VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                           rsaux4.Close
                        Else
                           rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID) + " AND SITE_USE_ID = " + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           VAR_PROVEEDOR = rsaux4!collector_id
                           VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                          rsaux4.Close
                        End If
                        rsaux2.Close
                     Else
                        rsaux2.Open "select a.attribute8, B.description from oe_order_headers_all a, MTL_SECONDARY_INVENTORIES b where order_number = " + CStr(rs!source_header_number) + " and a.attribute8 = b.secondary_inventory_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           VAR_PROVEEDOR = rsaux4!collector_id
                           VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                           rsaux4.Close
                        Else
                           rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           VAR_PROVEEDOR = rsaux4!collector_id
                           VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                           rsaux4.Close
                        End If
                        rsaux2.Close
                     End If
                  Else
                     rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(IIf(IsNull(rs!CUST_ACCOUNT_ID), 0, rs!CUST_ACCOUNT_ID)) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     'rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID) + " AND SITE_USE_ID = " + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux4.EOF Then
                        VAR_PROVEEDOR = rsaux4!collector_id
                        VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                     Else
                        VAR_PROVEEDOR = 0
                        VAR_NOMBRE_PROVEEDOR = ""
                     End If
                     rsaux4.Close
                  End If
                  
'''''

                  If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Then
                     rsaux1.Open "SELECT * FROM TB_ORACLE_ARTICULOS_MOTOR_LOGISTICO WHERE CODIGO = '" + rs!SEGMENT1 + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux1.EOF Then
                        strconsulta = "SELECT secondary_inventory_name, A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!source_document_id)
                             .Parameters.Append parametro
                        End With
                        Set rsaux8 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If rsaux8.EOF Then
                           var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                        Else
                           var_almacen = rsaux8!secondary_inventory_name
                           rsaux9.Open "SELECT * FROM TB_ORACLE_UBICACIONES_MOTOR_LOGISTICO WHERE CLAVE = '" + var_almacen + "' AND CODIGO = '" + rs!SEGMENT1 + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux9.EOF Then
                              var_ubicacion = ""
                              If Me.cmb_dia.Text = "Lunes" Then
                                 var_ubicacion = rsaux9!ubicacion_1
                              End If
                              If Me.cmb_dia.Text = "Martes" Then
                                 var_ubicacion = rsaux9!ubicacion_2
                              End If
                              If Me.cmb_dia.Text = "Miercoles" Then
                                 var_ubicacion = rsaux9!ubicacion_3
                              End If
                              If Me.cmb_dia.Text = "Jueves" Then
                                 var_ubicacion = rsaux9!ubicacion_4
                              End If
                              If Me.cmb_dia.Text = "Viernes" Then
                                 var_ubicacion = rsaux9!ubicacion_5
                              End If
                              If Me.cmb_dia.Text = "Sabado" Then
                                 var_ubicacion = rsaux9!ubicacion_6
                              End If
                              If IIf(IsNull(var_ubicacion), "", var_ubicacion) = var_ubicacion = "" Then
                                 var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                              End If
                              If var_ubicacion = "" Then
                                 var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                              End If
                           Else
                              var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                           End If
                           rsaux9.Close
                        End If
                        rsaux8.Close
                     Else
                        var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                     End If
                     rsaux1.Close
                  Else
                     var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                  End If
                 
                  
'''''
                  var_cadena = "insert into tb_temp_oracle_orden_surtido (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, DELIVERY_ID, DELIVERY_DETAIL_ID, ORGANIZATION_ID, SUBINVENTORY, DELIVERY_LINE_ID, INVENTORY_ITEM_ID, ITEM_DESCRIPTION, SOURCE_LINE_NUMBER, src_requested_quantity, RELEASED_STATUS, CUSTOMER_NAME, SEGMENT1, COLLECTOR_ID, NAME, date_requested, UBICACION, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, RUTA, NOMBRE_RUTA)  values "
                  var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rs!source_header_number), "", rs!source_header_number) + "', " + CStr(IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)) + ", " + CStr(IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)) + ", " + CStr(IIf(IsNull(rs!organization_id), 0, rs!organization_id)) + ", '" + IIf(IsNull(rs!subinventory), "", rs!subinventory) + "', " + CStr(IIf(IsNull(rs!delivery_line_id), 0, rs!delivery_line_id)) + ", " + CStr(IIf(IsNull(rs!inventory_item_id), "", rs!inventory_item_id)) + ", '" + IIf(IsNull(rs!Description), "", rs!Description) + "', '" + IIf(IsNull(rs!SOURCE_LINE_NUMBER), "", rs!SOURCE_LINE_NUMBER) + "', " + CStr(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)) + ", '" + IIf(IsNull(rs!released_status), "", rs!released_status) + "', '" + IIf(IsNull(rs!customer_name), "", rs!customer_name) + "', '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'"
                  'var_cadena = var_cadena + ", " + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "'," + var_fecha + ",'" + IIf(IsNull(rs!ATTRIBUTE2), "", rs!ATTRIBUTE2) + "','" + CStr(VAR_ESTABLECIMIENTO) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "', '" + CStr(rs!CLAVE_RUTA) + "','" + IIf(IsNull(rs!nombre_ruta), "", rs!nombre_ruta) + "')"
                  var_cadena = var_cadena + ", " + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "'," + var_fecha + ",'" + var_ubicacion + "','" + CStr(var_establecimiento) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "', '" + CStr(rs!CLAVE_RUTA) + "','" + IIf(IsNull(rs!nombre_ruta), "", rs!nombre_ruta) + "')"
                  'MsgBox var_cadena
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
                     x = 1
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
         Else
            MsgBox "No existen ordenes de surtido", vbOKOnly, "ATENCION"
         End If
         rs.Close
         rs.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "Numeracion incorrecta", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub cmd_imprimir_divididas_Click()
   Dim var_consecutivo As Integer
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_cadena_agentes = ""
   For var_j = 1 To lv_agentes.ListItems.Count
       lv_agentes.ListItems.Item(var_j).Selected = True
       If lv_agentes.selectedItem.SubItems(2) = "*" Then
          If var_cadena_agentes = "" Then
             var_cadena_agentes = lv_agentes.selectedItem
          Else
             var_cadena_agentes = var_cadena_agentes + "," + Me.lv_agentes.selectedItem
          End If
       End If
   Next var_j
   If IsDate(Me.txt_fecha_inicio) Then
      If IsDate(Me.txt_fecha_fin) Then
         If var_cadena_agentes <> "" Then
            var_cadena = "SELECT DISTINCT a.SOURCE_HEADER_NUMBER AS PEDIDO from WSH_DLVB_DLVY_V B, hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, XXVIA_VW_AGENTES D Where a.delivery_id = B.delivery_id AND A.delivery_detail_id = B.delivery_detail_id AND HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID and TRUNC(B.LAST_UPDATE_DATE) >= to_date('" + Me.txt_fecha_inicio + "','DD-MM-YYYY') AND TRUNC(B.LAST_UPDATE_DATE) < TO_DATE('" + CStr(CDate(Me.txt_fecha_fin) + 1) + "','DD-MM-YYYY') AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.COLLECTOR_ID IN (" + var_cadena_agentes + ") "
            var_cadena = var_cadena + " AND HCAS.CUST_ACCOUNT_ID = D.CUST_ACCOUNT_ID AND OHA.SHIP_FROM_ORG_ID = " + var_unidad_organizacional + " AND A.RELEASED_STATUS = 'Y' ORDER BY a.SOURCE_HEADER_NUMBER"
            rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_Cadena_pedidos = ""
            While Not rsaux2.EOF
                  If Trim(var_Cadena_pedidos) = "" Then
                     var_Cadena_pedidos = CStr(rsaux2!pedido)
                  Else
                     var_Cadena_pedidos = var_Cadena_pedidos + "," + CStr(rsaux2!pedido)
                  End If
                  rsaux2.MoveNext
            Wend
            rsaux2.Close
            If var_Cadena_pedidos <> "" Then
               var_cadena = "SELECT CAT.LINEA, HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id, a.LAST_UPDATE_DATE,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,c.description,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, a.source_header_type_name, oha.source_document_id, C.ATTRIBUTE2, oha.attribute8, oha.attribute9, j.NAME as nombre_ruta, j.salesrep_id as clave_ruta from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, xxvia_vw_articulos_cat cat, XXVIA_VENDEDORES j Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) IN (" + var_Cadena_pedidos + ") "
               var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y' AND A.inventory_item_id  = cat.item_id AND A.ORGANIZATION_ID = Cat.organization_id and oha.salesrep_id = j.salesrep_id "
               rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
'---------------------------------------
               If Not rs.EOF Then
                  cnn.BeginTrans
                  rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM tb_Temp_oracle_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
                  Else
                     var_consecutivo = 1
                  End If
                  rsaux.Close
                  rsaux1.Open "insert into tb_Temp_oracle_orden_surtido(inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  cnn.CommitTrans
                  While Not rs.EOF
                        var_establecimiento = rs!SHIP_TO_ORG_ID
                        rsaux.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'SHIP_TO' AND csu.LOCATION = ps.party_site_number and csu.site_use_id = " + CStr(var_establecimiento), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux!vcha_esb_nombre), "", rsaux!vcha_esb_nombre)
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
                        var_cadena = "insert into tb_Temp_oracle_orden_surtido(INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, DELIVERY_ID, DELIVERY_DETAIL_ID, ORGANIZATION_ID, SUBINVENTORY, DELIVERY_LINE_ID, INVENTORY_ITEM_ID, ITEM_DESCRIPTION, SOURCE_LINE_NUMBER, src_requested_quantity, RELEASED_STATUS, CUSTOMER_NAME, SEGMENT1, COLLECTOR_ID, NAME, date_requested, UBICACION, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, SITE_USE_ID, attribute8, attribute9, LINEA, RUTA, NOMBRE_RUTA)  values "
                        var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rs!source_header_number), "", rs!source_header_number) + "', " + CStr(IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)) + ", " + CStr(IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)) + ", " + CStr(IIf(IsNull(rs!organization_id), 0, rs!organization_id)) + ", '" + IIf(IsNull(rs!subinventory), "", rs!subinventory) + "', " + CStr(IIf(IsNull(rs!delivery_line_id), 0, rs!delivery_line_id)) + ", " + CStr(IIf(IsNull(rs!inventory_item_id), "", rs!inventory_item_id)) + ", '" + IIf(IsNull(rs!Description), "", rs!Description) + "', '" + IIf(IsNull(rs!SOURCE_LINE_NUMBER), "", rs!SOURCE_LINE_NUMBER) + "', " + CStr(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)) + ", '" + IIf(IsNull(rs!released_status), "", rs!released_status) + "', '" + IIf(IsNull(rs!customer_name), "", rs!customer_name) + "', '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'"
                        var_cadena = var_cadena + ", 0,''," + var_fecha + ",'" + IIf(IsNull(rs!attribute2), "", rs!attribute2) + "','" + CStr(var_establecimiento) + "','" + VAR_NOMBRE_ESTABLECIMIENTO + "'," + CStr(rs!CUST_ACCOUNT_ID) + ",'" + rs!source_header_type_name + "','" + CStr(IIf(IsNull(rs!source_document_id), "", rs!source_document_id)) + "'," + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)) + ",'" + IIf(IsNull(rs!attribute8), "", rs!attribute8) + "','" + IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9) + "','" + IIf(IsNull(rs!Linea), "", rs!Linea) + "','" + CStr(rs!CLAVE_RUTA) + "','" + rs!nombre_ruta + "')"
                        rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        rs.MoveNext
                  Wend
                  rsaux1.Open "delete from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number is null", cnn, adOpenDynamic, adLockOptimistic
                  If rsaux1.State = 1 Then
                     rsaux1.Close
                  End If
                  rsaux1.Open "select DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_Temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " group by DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA"
                  While Not rsaux1.EOF
                        If rsaux1!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rsaux1!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                           If var_pedido_tienda = 0 Then
                              If rsaux2.State = 1 Then
                                 rsaux2.Close
                              End If
                              rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux1!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           Else
                              rsaux2.Open "select a.attribute8, B.description from oe_order_headers_all a, MTL_SECONDARY_INVENTORIES b where order_number = " + CStr(rsaux1!source_header_number) + " and a.attribute8 = b.secondary_inventory_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           End If
                        Else
                           rsaux6.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           VAR_PROVEEDOR = rsaux6!collector_id
                           VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux6!Name), "", rsaux6!Name)
                           rsaux6.Close
                        End If
                        var_año_str = CStr(Year(rsaux1!DATE_REQUESTED))
                        If Len(var_año_str) < 2 Then
                           var_año_str = "20" + var_año_str
                        End If
                        var_mes_str = CStr(Month(rsaux1!DATE_REQUESTED))
                        If Len(var_mes_str) < 2 Then
                           var_mes_str = "0" + var_mes_str
                        End If
                        var_dia_str = CStr(Day(rsaux1!DATE_REQUESTED))
                        If Len(var_dia_str) < 2 Then
                           var_dia_str = "0" + var_dia_str
                        End If
                        var_hora_str = CStr(Hour(rsaux1!DATE_REQUESTED))
                        If Len(var_hora_str) < 2 Then
                           var_hora_str = "0" + var_hora_str
                        End If
                        VAR_MINUTO_STR = CStr(Minute(rsaux1!DATE_REQUESTED))
                        If Len(VAR_MINUTO_STR) < 2 Then
                           VAR_MINUTO_STR = "0" + VAR_MINUTO_STR
                        End If
                        VAR_SEGUNDO_STR = CStr(Second(rsaux1!DATE_REQUESTED))
                        If Len(VAR_SEGUNDO_STR) < 2 Then
                           VAR_SEGUNDO_STR = "0" + VAR_SEGUNDO_STR
                        End If
                        var_fecha_pedido = var_año_str + "-" + var_mes_str + "-" + var_dia_str + " " + var_hora_str + ":" + VAR_MINUTO_STR + ":" + VAR_SEGUNDO_STR
                        rsaux6.Open "update tb_Temp_oracle_orden_surtido set COLLECTOR_ID = " + CStr(VAR_PROVEEDOR) + ", NAME = '" + VAR_NOMBRE_PROVEEDOR + "' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and source_header_number = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        rsaux6.Open "SELECT * FROM XXVIA_TB_ORDENES_GRAFICA WHERE ORGANIZACION = " + var_unidad_organizacional + " AND PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           rsaux5.Open "INSERT INTO XXVIA_TB_ORDENES_GRAFICA (ORGANIZACION, PEDIDO, PROVEEDOR_ID, NOMBRE_PROVEEDOR, CLIENTE, ESTABLECIMIENTO, CANTIDAD, FECHA, TIPO_PEDIDO, CANTIDAD_LEIDA, FECHA_PEDIDO, RUTA, NOMBRE_RUTA) VALUES (" + var_unidad_organizacional + ", " + CStr(rsaux1!source_header_number) + "," + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "'," + CStr(rsaux1!cantidad) + ", '" + CStr(rsaux1!DATE_REQUESTED) + "','" + rsaux1!source_header_type_name + "',0, TO_DATE('" + var_fecha_pedido + "','YYYY-MM-DD HH24:MI:SS'),'" + rsaux1!ruta + "', '" + rsaux1!nombre_ruta + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET RUTA = '" + rsaux1!ruta + "', NOMBRE_RUTA = '" + rsaux1!nombre_ruta + "' WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        rsaux6.Open "select * from tb_oracle_pedidos_asignados_embarques where pedido = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           rsaux5.Open "insert into tb_oracle_pedidos_asignados_embarques (AGENTE, NOMBRE_AGENTE, PEDIDO, CLIENTE, PIEZAS, embarque, dia,  mes, AÑO, ORGANIZACION) values (" + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "'," + CStr(rsaux1!source_header_number) + ",'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "', " + CStr(rsaux1!cantidad) + ",0," + var_dia_str + "," + var_mes_str + "," + var_año_str + "," + CStr(var_unidad_organizacional) + ")", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "select DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_Temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " AND (LINEA <> 'CATALOGOS' OR LINEA IS NULL) group by DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME"
                  While Not rsaux1.EOF
                        rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET CANTIDAD_SIN_CATALOGOS = " + CStr(IIf(IsNull(rsaux1!cantidad), 0, rsaux1!cantidad)) + " WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "DELETE from tb_Temp_oracle_orden_surtido_aux_2", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "SELECT * FROM tb_Temp_oracle_orden_surtido where inte_tem_consecutivo =  " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        If rsaux1!Linea = "CATALOGOS" Then
                           var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           If Len(Trim(var_dia)) = 1 Then
                              var_dia = "0" + var_dia
                           End If
                           If Len(Trim(var_mes)) = 1 Then
                              var_mes = "0" + var_mes
                           End If
                           var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                           var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                           var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION) "
                           var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux1!src_requested_quantity) + ",'" + rsaux1!released_status + "',"
                           var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                           var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                           var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "')"
                           rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        Else
                           var_cantidad = rsaux1!src_requested_quantity
                           While var_cantidad > 0
                                 var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(1) + ",'" + rsaux1!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                 var_cantidad = var_cantidad - 1
                           Wend
                        End If
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "delete from tb_Temp_oracle_orden_surtido_aux_1", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "select distinct source_header_number from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        var_lote = 1
                        var_contador = 0
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              If var_contador = 50 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + "  and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + ")"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 var_contador = var_contador + 1
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "insert TB_TEMP_ORACLE_ORDEN_SURTIDO (inte_tem_consecutivo, segment1) values (" + CStr(var_consecutivo) + ",'---------')", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and segment1 <> '---------'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "insert into TB_TEMP_ORACLE_ORDEN_SURTIDO select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and segment1 = '---------'", cnn, adOpenDynamic, adLockOptimistic
                  var_consecutivo_general = var_consecutivo
                  Call crea_tablas
                  rsaux.Open "select distinct source_header_number, lote from tb_Temp_oracle_orden_surtido_aux_1 where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     While Not rsaux.EOF
                           x = 1
                           If x = 1 Then
                              rsaux2.Open "SELECT * FROM tb_oracle_pedidos_asignados_embarques WHERE PEDIDO = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 rsaux3.Open "UPDATE tb_Temp_oracle_orden_surtido SET EMBARQUE = " + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", ORDEN_SURTIDO = " + CStr(IIf(IsNull(rsaux2!orden_pedido), 0, rsaux2!orden_pedido)) + ", ESTACION = '" + CStr(IIf(IsNull(rsaux2!estacion), 0, rsaux2!estacion)) + "' WHERE SOURCE_HEADER_NUMBER = " + CStr(rsaux(0).Value) + " AND inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux2.Close
                              Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA.rpt")
                              reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.SOURCE_HEADER_NUMBER} = '" + rsaux(0).Value + "' and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_ORDEN_SURTIDO.LOTE} = " + CStr(rsaux(1).Value)
                              For ntablas = 1 To reporte.Database.Tables.Count
                                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                              Next ntablas
                              reporte.PrintOut False
                              Set reporte = Nothing
                           End If
                           rsaux.MoveNext
                     Wend
                  End If
                  rsaux.Close
               Else
                  MsgBox "No existen ordenes de surtido", vbOKOnly, "ATENCION"
               End If
               rs.Close
'------------------
               rs.Open "delete from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            Else
               MsgBox "No existen pedidos para el periodo seleccionado", vbOKOnly, "ATENCION"
            End If
            
         Else
            MsgBox "No se seleccionaron agentes", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha inicial incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_divididas_numero_Click()
      If IsNumeric(Me.txt_numero_inicio) Then
         If IsNumeric(Me.txt_numero_fin) Then
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = "SELECT CAT.LINEA, HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id, a.LAST_UPDATE_DATE,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,c.description,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, a.source_header_type_name, oha.source_document_id, C.ATTRIBUTE2, oha.attribute8, oha.attribute9, j.NAME as nombre_ruta, j.salesrep_id as clave_ruta from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, xxvia_vw_articulos_cat cat, XXVIA_VENDEDORES j Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND "
            var_cadena = var_cadena + " to_number(source_header_number)  BETWEEN " + CStr(Me.txt_numero_inicio) + " AND " + CStr(Me.txt_numero_fin) + ""
            var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID "
            var_cadena = var_cadena + " AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y' AND A.inventory_item_id  = cat.item_id AND A.ORGANIZATION_ID = Cat.organization_id and oha.salesrep_id = j.salesrep_id "
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
'--------------------------
            If Not rs.EOF Then
               cnn.BeginTrans
               rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM tb_Temp_oracle_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
               Else
                  var_consecutivo = 1
               End If
               rsaux.Close
               rsaux1.Open "insert into tb_Temp_oracle_orden_surtido(inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               While Not rs.EOF
                     var_establecimiento = rs!SHIP_TO_ORG_ID
                     rsaux.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'SHIP_TO' AND csu.LOCATION = ps.party_site_number and csu.site_use_id = " + CStr(var_establecimiento), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux!vcha_esb_nombre), "", rsaux!vcha_esb_nombre)
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
                     var_cadena = "insert into tb_Temp_oracle_orden_surtido(INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, DELIVERY_ID, DELIVERY_DETAIL_ID, ORGANIZATION_ID, SUBINVENTORY, DELIVERY_LINE_ID, INVENTORY_ITEM_ID, ITEM_DESCRIPTION, SOURCE_LINE_NUMBER, src_requested_quantity, RELEASED_STATUS, CUSTOMER_NAME, SEGMENT1, COLLECTOR_ID, NAME, date_requested, UBICACION, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, SITE_USE_ID, attribute8, attribute9, LINEA, RUTA, NOMBRE_RUTA)  values "
                     var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rs!source_header_number), "", rs!source_header_number) + "', " + CStr(IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)) + ", " + CStr(IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)) + ", " + CStr(IIf(IsNull(rs!organization_id), 0, rs!organization_id)) + ", '" + IIf(IsNull(rs!subinventory), "", rs!subinventory) + "', " + CStr(IIf(IsNull(rs!delivery_line_id), 0, rs!delivery_line_id)) + ", " + CStr(IIf(IsNull(rs!inventory_item_id), "", rs!inventory_item_id)) + ", '" + IIf(IsNull(rs!Description), "", rs!Description) + "', '" + IIf(IsNull(rs!SOURCE_LINE_NUMBER), "", rs!SOURCE_LINE_NUMBER) + "', " + CStr(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)) + ", '" + IIf(IsNull(rs!released_status), "", rs!released_status) + "', '" + IIf(IsNull(rs!customer_name), "", rs!customer_name) + "', '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'"
                     var_cadena = var_cadena + ", 0,''," + var_fecha + ",'" + IIf(IsNull(rs!attribute2), "", rs!attribute2) + "','" + CStr(var_establecimiento) + "','" + VAR_NOMBRE_ESTABLECIMIENTO + "'," + CStr(rs!CUST_ACCOUNT_ID) + ",'" + rs!source_header_type_name + "','" + CStr(IIf(IsNull(rs!source_document_id), "", rs!source_document_id)) + "'," + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)) + ",'" + IIf(IsNull(rs!attribute8), "", rs!attribute8) + "','" + IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9) + "','" + IIf(IsNull(rs!Linea), "", rs!Linea) + "','" + CStr(rs!CLAVE_RUTA) + "','" + IIf(IsNull(rs!nombre_ruta), "", rs!nombre_ruta) + "')"
                     rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rsaux1.Open "delete from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number is null", cnn, adOpenDynamic, adLockOptimistic
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
               rsaux1.Open "select DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_Temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " group by DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux1.EOF
                     If rsaux1!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rsaux1!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                        If var_pedido_tienda = 0 Then
                           If rsaux2.State = 1 Then
                              rsaux2.Close
                           End If
                           rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux1!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux2.EOF Then
                              rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              VAR_PROVEEDOR = rsaux4!collector_id
                              VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                              rsaux4.Close
                           Else
                              rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              VAR_PROVEEDOR = rsaux4!collector_id
                              VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                              rsaux4.Close
                           End If
                           rsaux2.Close
                        Else
                           rsaux2.Open "select a.attribute8, B.description from oe_order_headers_all a, MTL_SECONDARY_INVENTORIES b where order_number = " + CStr(rsaux1!source_header_number) + " and a.attribute8 = b.secondary_inventory_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux2.EOF Then
                              rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              VAR_PROVEEDOR = rsaux4!collector_id
                              VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                              rsaux4.Close
                           Else
                              rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              VAR_PROVEEDOR = rsaux4!collector_id
                              VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                              rsaux4.Close
                           End If
                           rsaux2.Close
                        End If
                     Else
                        rsaux6.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        VAR_PROVEEDOR = rsaux6!collector_id
                        VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux6!Name), "", rsaux6!Name)
                        rsaux6.Close
                     End If
                     var_año_str = CStr(Year(rsaux1!DATE_REQUESTED))
                     If Len(var_año_str) < 2 Then
                        var_año_str = "20" + var_año_str
                     End If
                     var_mes_str = CStr(Month(rsaux1!DATE_REQUESTED))
                     If Len(var_mes_str) < 2 Then
                        var_mes_str = "0" + var_mes_str
                     End If
                     var_dia_str = CStr(Day(rsaux1!DATE_REQUESTED))
                     If Len(var_dia_str) < 2 Then
                        var_dia_str = "0" + var_dia_str
                     End If
                     var_hora_str = CStr(Hour(rsaux1!DATE_REQUESTED))
                     If Len(var_hora_str) < 2 Then
                        var_hora_str = "0" + var_hora_str
                     End If
                     VAR_MINUTO_STR = CStr(Minute(rsaux1!DATE_REQUESTED))
                     If Len(VAR_MINUTO_STR) < 2 Then
                        VAR_MINUTO_STR = "0" + VAR_MINUTO_STR
                     End If
                     VAR_SEGUNDO_STR = CStr(Second(rsaux1!DATE_REQUESTED))
                     If Len(VAR_SEGUNDO_STR) < 2 Then
                        VAR_SEGUNDO_STR = "0" + VAR_SEGUNDO_STR
                     End If
                     var_fecha_pedido = var_año_str + "-" + var_mes_str + "-" + var_dia_str + " " + var_hora_str + ":" + VAR_MINUTO_STR + ":" + VAR_SEGUNDO_STR
                     rsaux6.Open "update tb_Temp_oracle_orden_surtido set COLLECTOR_ID = " + CStr(VAR_PROVEEDOR) + ", NAME = '" + VAR_NOMBRE_PROVEEDOR + "' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and source_header_number = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                     rsaux6.Open "SELECT * FROM XXVIA_TB_ORDENES_GRAFICA WHERE ORGANIZACION = " + var_unidad_organizacional + " AND PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If rsaux6.EOF Then
                        rsaux5.Open "INSERT INTO XXVIA_TB_ORDENES_GRAFICA (ORGANIZACION, PEDIDO, PROVEEDOR_ID, NOMBRE_PROVEEDOR, CLIENTE, ESTABLECIMIENTO, CANTIDAD, FECHA, TIPO_PEDIDO, CANTIDAD_LEIDA, FECHA_PEDIDO, RUTA, NOMBRE_RUTA) VALUES (" + var_unidad_organizacional + ", " + CStr(rsaux1!source_header_number) + "," + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "'," + CStr(rsaux1!cantidad) + ", '" + CStr(rsaux1!DATE_REQUESTED) + "','" + rsaux1!source_header_type_name + "',0, TO_DATE('" + var_fecha_pedido + "','YYYY-MM-DD HH24:MI:SS'),'" + rsaux1!ruta + "', '" + rsaux1!nombre_ruta + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     Else
                        rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET RUTA = '" + rsaux1!ruta + "', NOMBRE_RUTA = '" + rsaux1!nombre_ruta + "' WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux6.Close
                     rsaux6.Open "select * from tb_oracle_pedidos_asignados_embarques where pedido = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                     If rsaux6.EOF Then
                        rsaux5.Open "insert into tb_oracle_pedidos_asignados_embarques (AGENTE, NOMBRE_AGENTE, PEDIDO, CLIENTE, PIEZAS, embarque, dia,  mes, AÑO, ORGANIZACION) values (" + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "'," + CStr(rsaux1!source_header_number) + ",'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "', " + CStr(rsaux1!cantidad) + ",0," + var_dia_str + "," + var_mes_str + "," + var_año_str + "," + CStr(var_unidad_organizacional) + ")", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux6.Close
                     rsaux1.MoveNext
               Wend
               rsaux1.Close
               rsaux1.Open "select DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_Temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " AND (LINEA <> 'CATALOGOS' OR LINEA IS NULL) group by DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME"
               While Not rsaux1.EOF
                     rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET CANTIDAD_SIN_CATALOGOS = " + CStr(IIf(IsNull(rsaux1!cantidad), 0, rsaux1!cantidad)) + " WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     rsaux1.MoveNext
               Wend
               rsaux1.Close
               rsaux1.Open "DELETE from tb_Temp_oracle_orden_surtido_aux_2", cnn, adOpenDynamic, adLockOptimistic
               rsaux1.Open "SELECT * FROM tb_Temp_oracle_orden_surtido where inte_tem_consecutivo =  " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux1.EOF
                     If rsaux1!Linea = "CATALOGOS" Then
                        var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                        var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                        var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                        If Len(Trim(var_dia)) = 1 Then
                           var_dia = "0" + var_dia
                        End If
                        If Len(Trim(var_mes)) = 1 Then
                           var_mes = "0" + var_mes
                        End If
                        var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                        var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                        var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION) "
                        var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux1!src_requested_quantity) + ",'" + rsaux1!released_status + "',"
                        var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                        var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                        var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "')"
                        rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     Else
                        var_cantidad = rsaux1!src_requested_quantity
                        While var_cantidad > 0
                              var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                              var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                              var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                              If Len(Trim(var_dia)) = 1 Then
                                 var_dia = "0" + var_dia
                              End If
                              If Len(Trim(var_mes)) = 1 Then
                                 var_mes = "0" + var_mes
                              End If
                              var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                              var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                              var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION) "
                              var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(1) + ",'" + rsaux1!released_status + "',"
                              var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                              var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                              var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "')"
                              rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              var_cantidad = var_cantidad - 1
                        Wend
                     End If
                     rsaux1.MoveNext
               Wend
               rsaux1.Close
               rsaux1.Open "delete from tb_Temp_oracle_orden_surtido_aux_1", cnn, adOpenDynamic, adLockOptimistic
               rsaux1.Open "select distinct source_header_number from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux1.EOF
                     var_lote = 1
                     var_contador = 0
                     rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                     While Not rsaux2.EOF
                           If var_contador = 50 Then
                              var_lote = var_lote + 1
                              var_contador = 0
                           End If
                           rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux3.EOF Then
                              rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                           Else
                              var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                              var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                              var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                              If Len(Trim(var_dia)) = 1 Then
                                 var_dia = "0" + var_dia
                              End If
                              If Len(Trim(var_mes)) = 1 Then
                                 var_mes = "0" + var_mes
                              End If
                              var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                              
                              var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                              var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE) "
                              var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                              var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                              var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                              var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + ")"
                              rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux3.Close
                           If rsaux2!Linea <> "CATALOGOS" Then
                              var_contador = var_contador + 1
                           End If
                           rsaux2.MoveNext
                     Wend
                     rsaux2.Close
                     rsaux1.MoveNext
               Wend
               rsaux1.Close
               rsaux1.Open "insert TB_TEMP_ORACLE_ORDEN_SURTIDO (inte_tem_consecutivo, segment1) values (" + CStr(var_consecutivo) + ",'---------')", cnn, adOpenDynamic, adLockOptimistic
               rsaux1.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and segment1 <> '---------'", cnn, adOpenDynamic, adLockOptimistic
               rsaux2.Open "insert into TB_TEMP_ORACLE_ORDEN_SURTIDO select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               rsaux2.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and segment1 = '---------'", cnn, adOpenDynamic, adLockOptimistic
               var_consecutivo_general = var_consecutivo
               Call crea_tablas
               rsaux.Open "select distinct source_header_number, lote from tb_Temp_oracle_orden_surtido_aux_1 where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  While Not rsaux.EOF
                        x = 1
                        If x = 1 Then
                           rsaux2.Open "SELECT * FROM tb_oracle_pedidos_asignados_embarques WHERE PEDIDO = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux2.EOF Then
                              rsaux3.Open "UPDATE tb_Temp_oracle_orden_surtido SET EMBARQUE = " + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", ORDEN_SURTIDO = " + CStr(IIf(IsNull(rsaux2!orden_pedido), 0, rsaux2!orden_pedido)) + ", ESTACION = '" + CStr(IIf(IsNull(rsaux2!estacion), 0, rsaux2!estacion)) + "' WHERE SOURCE_HEADER_NUMBER = " + CStr(rsaux(0).Value) + " AND inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux2.Close
                           x = 1
                           If x = 1 Then
                           Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA.rpt")
                           reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.SOURCE_HEADER_NUMBER} = '" + rsaux(0).Value + "' and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_ORDEN_SURTIDO.LOTE} = " + CStr(rsaux(1).Value)
                           For ntablas = 1 To reporte.Database.Tables.Count
                               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                           Next ntablas
                           reporte.PrintOut False
                           Set reporte = Nothing
                           End If
                        End If
                        rsaux.MoveNext
                  Wend
               End If
               rsaux.Close
            Else
               MsgBox "No existen ordenes de surtido", vbOKOnly, "ATENCION"
            End If
            rs.Close
            rs.Open "delete from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "Número superior incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Número inferior incorrecto"
      End If
End Sub

Private Sub cmd_imprimir_entradas_Click()
   frmoracle_imprimir_ordenes_surtido_entradas.Show 1
End Sub

Private Sub cmd_imprimir_pedido_tiendas_Click()
   x = 0
   If x = 0 Then
   rs.Open "select distinct codigo from pedidos_tiendas_120313", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         strconsulta = "select attribute2 ubicacion1, attribute3  ubicacion2, attribute4  ubicacion3, attribute5  ubicacion4, attribute6 ubicacion5, attribute7 ubicacion6, inventory_item_id item_id, segment1 item_number, description item_description from  xxvia_system_items_b where  organization_id  = ? and segment1 = ? order by segment1"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_unidad_organizacional))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!codigo)
              .Parameters.Append parametro
         End With
         Set rsaux9 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rsaux9.EOF Then
            rsaux.Open "update pedidos_tiendas_120313 set ubicacion = '" + IIf(IsNull(rsaux9!UBICACION1), "", rsaux9!UBICACION1) + "' where codigo = '" + rs!codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux9.Close
         strconsulta = "select * from Xxvia_vw_existencias_inv where organization_id = ? and subinventory_code = ? and segment1 = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_unidad_organizacional))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, "CDI_ALMPT")
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!codigo)
             .Parameters.Append parametro
         End With
         Set rsaux9 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         'rsaux9.Open "SELECT * FROM Xxvia_vw_existencias_inv WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND SEGMENT1 = '" + Me.txt_codigo + "'  AND SUBINVENTORY_CODE = '" + Me.txt_clave_almacen + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux9.EOF Then
            txt_fisico = IIf(IsNull(rsaux9!CANTMANO), 0, rsaux9!CANTMANO)
            txt_apartado = IIf(IsNull(rsaux9!RESERVADA), 0, rsaux9!RESERVADA)
            txt_disponible = IIf(IsNull(rsaux9!Disponible), 0, rsaux9!Disponible)
         Else
            txt_fisico = 0
            txt_apartado = 0
            txt_disponible = 0
         End If
         If Not rsaux9.EOF Then
            rsaux.Open "update pedidos_tiendas_120313 set existencias = " + CStr(txt_disponible) + " where codigo = '" + rs!codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux9.Close
         
         rs.MoveNext
   
   Wend
   rs.Close
   End If
   rs.Open "select distinct tienda_id from pedidos_tiendas_120313", cnn, adOpenDynamic, adLockOptimistic
   var_i = 0
   While Not rs.EOF
         var_i = var_i + 1
         'If var_i = 1 Then
         Set reporte = appl.OpenReport(App.Path + "\rep_oracle_os_tiendas_sin_pedido.rpt")
         reporte.RecordSelectionFormula = "{VW_ORACLE_OS_TIENDAS_SIN_PEDIDO.Tienda_id} = '" + rs!tienda_id + "' AND {VW_ORACLE_OS_TIENDAS_SIN_PEDIDO.existencias}>0"
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         reporte.PrintOut False
         Set reporte = Nothing
         'End If
         rs.MoveNext
   Wend
   rs.Close
End Sub


Private Sub cmd_invertir_Click()
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      If lv_agentes.selectedItem.SubItems(2) = "*" Then
         lv_agentes.selectedItem.SubItems(2) = ""
         lv_agentes.ListItems.Item(i).Bold = False
         lv_agentes.ListItems.Item(i).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_agentes.selectedItem.SubItems(2) = "*"
         lv_agentes.ListItems.Item(i).Bold = True
         lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   i = lv_agentes.selectedItem.Index
   If lv_agentes.selectedItem.SubItems(2) = "*" Then
      lv_agentes.selectedItem.SubItems(2) = ""
      lv_agentes.ListItems.Item(i).Bold = False
      lv_agentes.ListItems.Item(i).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_agentes.Refresh
   Else
      lv_agentes.selectedItem.SubItems(2) = "*"
      lv_agentes.ListItems.Item(i).Bold = True
      lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_agentes.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      lv_agentes.selectedItem.SubItems(2) = ""
      lv_agentes.ListItems.Item(i).Bold = False
      lv_agentes.ListItems.Item(i).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
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
      If var_encontro = True And lv_agentes.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_agentes.selectedItem.SubItems(2) = "*"
         lv_agentes.ListItems.Item(i).Bold = True
         lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_agentes.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_agentes.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub cmd_todos_Click()
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      lv_agentes.selectedItem.SubItems(2) = "*"
      lv_agentes.ListItems.Item(i).Bold = True
      lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_agentes.Refresh
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Form_Load()
   Top = 800
   Left = 3000
   Dim list_item As ListItem
   rs.Open "select COLLECTOR_ID, NAME from ar_collectors", cnnoracle_4, adOpenDynamic, adLockOptimistic
   numero_items_permisos = 0
   While Not rs.EOF
      Set list_item = lv_agentes.ListItems.Add(, , rs!collector_id)
      list_item.SubItems(1) = IIf(IsNull(rs!Name), "", rs!Name)
      list_item.SubItems(2) = ""
      rs.MoveNext:
      numero_items_permisos = numero_items_permisos + 1
   Wend
   rs.Close
   Me.txt_fecha_fin = Date
   Me.txt_fecha_inicio = Date - 7
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      lv_agentes.selectedItem.SubItems(2) = "*"
      lv_agentes.ListItems.Item(i).Bold = True
      lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_agentes.Refresh



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
         If lv_agentes.selectedItem.SubItems(2) = "*" Then
            lv_agentes.selectedItem.SubItems(2) = ""
            lv_agentes.ListItems.Item(i).Bold = False
            lv_agentes.ListItems.Item(i).ForeColor = &H80000012
            lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
            lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
            lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
            lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
            lv_agentes.Refresh
         Else
            lv_agentes.selectedItem.SubItems(2) = "*"
            lv_agentes.ListItems.Item(i).Bold = True
            lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
            lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
            lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
            lv_agentes.Refresh
         End If
      End If
   End If
End Sub

Private Sub txt_fecha_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha_fin) Then
         frmcalendario.mes = CDate(Me.txt_fecha_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fecha_fin = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_fin_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_fecha_inicio_Change()
   Me.txt_numero_fin = ""
   Me.txt_numero_inicio = ""
End Sub

Private Sub txt_fecha_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha_inicio) Then
         frmcalendario.mes = CDate(Me.txt_fecha_inicio)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fecha_inicio = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_inicio_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_fin_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_inicio_Change()
   Me.txt_fecha_inicio = ""
   Me.txt_fecha_fin = ""
End Sub

Private Sub txt_numero_inicio_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub
