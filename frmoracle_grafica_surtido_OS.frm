VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmoracle_grafica_surtido_OS 
   BorderStyle     =   0  'None
   Caption         =   "Grafica surtido"
   ClientHeight    =   9480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm_mensaje 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   3915
      TabIndex        =   14
      Top             =   3915
      Width           =   7365
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Actualizando ventana, espere un momento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   7380
      End
   End
   Begin VB.CommandButton cmd_cerrar 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   14970
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   45
      Width           =   255
   End
   Begin MSComCtl2.MonthView mes 
      Height          =   2370
      Left            =   2130
      TabIndex        =   3
      Top             =   6420
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   58392577
      CurrentDate     =   38148
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   45
      TabIndex        =   4
      Top             =   8640
      Width           =   6000
      Begin VB.CommandButton Command12 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5160
         Picture         =   "frmoracle_grafica_surtido_OS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fecha Final"
         Top             =   270
         Width           =   330
      End
      Begin VB.CommandButton Command11 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2220
         Picture         =   "frmoracle_grafica_surtido_OS.frx":1272
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fecha Inicial"
         Top             =   270
         Width           =   330
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   1080
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   660
         TabIndex        =   10
         Top             =   330
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3705
         TabIndex        =   9
         Top             =   330
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Reporte "
      Height          =   720
      Left            =   6090
      TabIndex        =   0
      Top             =   8640
      Width           =   9060
      Begin VB.CommandButton Command2 
         Height          =   345
         Left            =   4470
         Picture         =   "frmoracle_grafica_surtido_OS.frx":24E4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exportar Gráfica"
         Top             =   270
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   345
         Left            =   4095
         Picture         =   "frmoracle_grafica_surtido_OS.frx":27F6
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Actualiza Grafica"
         Top             =   270
         Width           =   375
      End
   End
   Begin MSComctlLib.ListView lv_grafica 
      Height          =   8205
      Left            =   60
      TabIndex        =   11
      Top             =   315
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   14473
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "O.S."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2910
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Agente"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cliente"
         Object.Width           =   6262
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Surtir"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Surtido"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "%"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Tipo"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Embarque"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Estatus"
         Object.Width           =   1412
      EndProperty
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "  Grafica de surtido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   15
      TabIndex        =   12
      Top             =   30
      Width           =   15225
   End
End
Attribute VB_Name = "frmoracle_grafica_surtido_OS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_fecha_inicio As String
Dim var_fecha_fin As String
Dim var_tipo_mes As Integer

Private Sub cmd_cerrar_Click()
   Me.frm_mensaje.Visible = True
   Me.Refresh
   var_contraseña_cerrar_pantalla = ""
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         var_filtrado = 0
         rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_fecha_inicio = CStr(CDate(Me.txt_inicio) - 1)
         var_fecha_inicio = var_fecha_inicio + " 19:59:59"
         var_cadena = "SELECT DISTINCT a.SOURCE_HEADER_NUMBER AS PEDIDO from WSH_DLVB_DLVY_V B, hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, XXVIA_VW_AGENTES D Where a.delivery_id = B.delivery_id AND A.delivery_detail_id = B.delivery_detail_id AND HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID and B.creation_date >= to_date('" + var_fecha_inicio + "','DD/MM/YYYY HH24:MI:SS') AND TRUNC(B.creation_date) < TO_DATE('" + CStr(CDate(Me.txt_fin) + 1) + "','DD-MM-YYYY') AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID "
         var_cadena = var_cadena + " AND HCAS.CUST_ACCOUNT_ID = D.CUST_ACCOUNT_ID AND OHA.SHIP_FROM_ORG_ID = " + var_unidad_organizacional + " ORDER BY a.SOURCE_HEADER_NUMBER"
         rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_cadena_pedidos = ""
         var_i = 0
         While Not rsaux2.EOF
               If Trim(var_cadena_pedidos) = "" Then
                  var_cadena_pedidos = CStr(rsaux2!PEDIDO)
               Else
                  var_cadena_pedidos = var_cadena_pedidos + "," + CStr(rsaux2!PEDIDO)
               End If
               var_i = var_i + 1
               rsaux2.MoveNext
         Wend
         rsaux2.Close
         Me.Refresh
         'var_cadena = "SELECT HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.source_header_number,A.organization_id,A.subinventory,A.released_status, a.source_header_type_name, oha.source_document_id, sum(A.requested_quantity) as cantidad, max(a.LAST_UPDATE_DATE) as fecha from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C "
         'var_cadena = var_cadena + " Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) IN (" + var_cadena_pedidos + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID  AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y'"
         'var_cadena = var_cadena + " group by HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1, A.source_header_number,A.organization_id,A.subinventory,A.released_status, a.source_header_type_name, oha.source_document_id"
         var_cadena = "SELECT HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID, HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, to_number(A.source_header_number) source_header_number, A.organization_id, A.subinventory, A.released_status, a.source_header_type_name, oha.source_document_id, SUM(A.requested_quantity) + (select nvl(sum(cantidad),0) from xxvia_tb_negado_distribucion where source_header_number = a.source_header_number and cantidad > 0) AS cantidad, MAX(A.LAST_UPDATE_DATE)   As Fecha FROM hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID=HL.LOCATION_ID AND HCSU.SITE_USE_ID = OHA.INVOICE_TO_ORG_ID "
         var_cadena = var_cadena + " AND to_number(source_header_number) IN (" + var_cadena_pedidos + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status in ('Y','C') GROUP BY source_header_number, HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID, HPS.LOCATION_ID, HL.ADDRESS1, A.organization_id, A.subinventory, A.released_status, a.source_header_type_name, oha.source_document_id"
         
         'MsgBox var_cadena
         rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            cnn.BeginTrans
            rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_tEMP_ORACLE_ORDEN_SURTIDO_grafica", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rsaux.Close
            rsaux1.Open "insert into tb_Temp_oracle_orden_surtido_grafica (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
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
                  var_dia = CStr(Day(CDate(rs!Fecha)))
                  var_mes = CStr(Month(CDate(rs!Fecha)))
                  var_año = CStr(Year(CDate(rs!Fecha)))
                  var_hora = CStr(Hour(CDate(rs!Fecha)))
                  var_minuto = CStr(Minute(CDate(rs!Fecha)))
                  var_segundo = CStr(Second(CDate(rs!Fecha)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  If Len(Trim(var_hora)) = 1 Then
                     var_hora = "0" + var_hora
                  End If
                  If Len(Trim(var_minuto)) = 1 Then
                     var_minuto = "0" + var_minuto
                  End If
                  If Len(Trim(var_segundo)) = 1 Then
                     var_segundo = "0" + var_segundo
                  End If
                  
                  var_fecha = "{ts '" + var_año + "-" + var_mes + "-" + var_dia + " " + var_hora + ":" + var_minuto + ":" + var_segundo + ".000'}"
                  
                  
                  
                  
                  var_cadena = "insert into tb_temp_oracle_orden_surtido_grafica (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, ORGANIZATION_ID, SUBINVENTORY, src_requested_quantity, RELEASED_STATUS, CUSTOMER_NAME, COLLECTOR_ID, NAME, date_requested, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, SITE_USE_ID)  values "
                  
                  var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", '" + CStr(IIf(IsNull(rs!source_header_number), "", rs!source_header_number)) + "', " + CStr(IIf(IsNull(rs!organization_id), 0, rs!organization_id)) + ", '" + IIf(IsNull(rs!subinventory), "", rs!subinventory) + "', " + CStr(rs!Cantidad) + ",'" + IIf(IsNull(rs!released_status), "", rs!released_status) + "', '" + Replace(IIf(IsNull(rs!customer_name), "", rs!customer_name), "'", " ") + "' "
                  var_cadena = var_cadena + ", 0,''," + var_fecha + ",'" + CStr(VAR_ESTABLECIMIENTO) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "'," + CStr(rs!CUST_ACCOUNT_ID) + ",'" + rs!source_header_type_name + "','" + CStr(IIf(IsNull(rs!source_document_id), "", rs!source_document_id)) + "'," + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)) + ")"
                  'MsgBox var_cadena
                  rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  Me.Refresh
                  rs.MoveNext
            Wend
            rsaux1.Open "delete from tb_temp_oracle_orden_surtido_grafica where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number is null", cnn, adOpenDynamic, adLockOptimistic
            If rsaux1.State = 1 Then
               rsaux1.Close
            End If
            rsaux1.Open "select DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_temp_oracle_orden_surtido_grafica where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " group by DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME"
            While Not rsaux1.EOF
                  If rsaux1!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rsaux1!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
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
                     rsaux6.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux6.EOF Then
                        VAR_PROVEEDOR = rsaux6!collector_id
                        VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux6!Name), "", rsaux6!Name)
                     Else
                        VAR_PROVEEDOR = 0
                        VAR_NOMBRE_PROVEEDOR = ""
                     End If
                     rsaux6.Close
                  End If
                  rsaux6.Open "update tb_temp_oracle_orden_surtido_grafica set COLLECTOR_ID = " + CStr(VAR_PROVEEDOR) + ", NAME = '" + VAR_NOMBRE_PROVEEDOR + "' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and source_header_number = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                  rsaux6.Open "SELECT * FROM XXVIA_TB_ORDENES_GRAFICA WHERE ORGANIZACION = " + var_unidad_organizacional + " AND PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If rsaux6.EOF Then
                     var_tipo_pedido = ""
                     If Trim(rsaux1!source_header_type_name) = "VIA_EXPORTACION" Then
                        var_tipo_pedido = "E"
                     End If
                     If Trim(rsaux1!source_header_type_name) = "VIA_PEDIDO_INTERNO" Then
                        var_tipo_pedido = "T"
                     End If
                     If Trim(rsaux1!source_header_type_name) = "VIA_MAYOREO_NACIONAL" Then
                        var_tipo_pedido = "M"
                     End If
                     If Trim(rsaux1!source_header_type_name) = "VIA_VTAS_X_TELEFONO" Then
                        var_tipo_pedido = "VXT"
                     End If

                     var_dia = CStr(Day(CDate(rsaux1!DATE_REQUESTED)))
                     var_mes = CStr(Month(CDate(rsaux1!DATE_REQUESTED)))
                     var_año = CStr(Year(CDate(rsaux1!DATE_REQUESTED)))
                     var_hora = CStr(Hour(CDate(rsaux1!DATE_REQUESTED)))
                     var_minuto = CStr(Minute(CDate(rsaux1!DATE_REQUESTED)))
                     var_segundo = CStr(Second(CDate(rsaux1!DATE_REQUESTED)))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     If Len(Trim(var_hora)) = 1 Then
                        var_hora = "0" + var_hora
                     End If
                     If Len(Trim(var_minuto)) = 1 Then
                        var_minuto = "0" + var_minuto
                     End If
                     If Len(Trim(var_segundo)) = 1 Then
                        var_segundo = "0" + var_segundo
                     End If
                     Me.Refresh
                     rsaux5.Open "INSERT INTO XXVIA_TB_ORDENES_GRAFICA (ORGANIZACION, PEDIDO, PROVEEDOR_ID, NOMBRE_PROVEEDOR, CLIENTE, ESTABLECIMIENTO, CANTIDAD, FECHA, TIPO_PEDIDO, CANTIDAD_LEIDA) VALUES (" + var_unidad_organizacional + ", " + CStr(rsaux1!source_header_number) + "," + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "'," + CStr(rsaux1!Cantidad) + ", TO_DATE('" + var_año + "-" + var_mes + "-" + var_dia + " " + var_hora + ":" + var_minuto + ":" + var_segundo + "','YYYY-MM-DD HH24:MI:SS'" + "),'" + var_tipo_pedido + "',0)", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  Else
                     var_tipo_pedido = ""
                     If Trim(rsaux1!source_header_type_name) = "VIA_EXPORTACION" Then
                        var_tipo_pedido = "E"
                     End If
                     If Trim(rsaux1!source_header_type_name) = "VIA_PEDIDO_INTERNO" Then
                        var_tipo_pedido = "T"
                     End If
                     If Trim(rsaux1!source_header_type_name) = "VIA_MAYOREO_NACIONAL" Then
                        var_tipo_pedido = "M"
                     End If
                     If Trim(rsaux1!source_header_type_name) = "VIA_VTAS_X_TELEFONO" Then
                        var_tipo_pedido = "VXT"
                     End If
                     rsaux5.Open "update XXVIA_TB_ORDENES_GRAFICA set tipo_pedido = '" + var_tipo_pedido + "' where pedido = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux6.Close
                  Me.Refresh
                  rsaux1.MoveNext
            Wend
            rsaux1.Close
                  
            var_dia = CStr(Day(CDate(Me.txt_inicio) - 1))
            var_mes = CStr(Month(CDate(Me.txt_inicio) - 1))
            var_año = CStr(Year(CDate(Me.txt_inicio) - 1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = var_año + "-" + var_mes + "-" + var_dia + " 19:59:59"
            
            
            var_dia = CStr(Day(CDate(Me.txt_fin) + 1))
            var_mes = CStr(Month(CDate(Me.txt_fin) + 1))
            var_año = CStr(Year(CDate(Me.txt_fin) + 1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = var_año + "-" + var_mes + "-" + var_dia
         End If
         rs.Close
         x = 1
         If x = 0 Then
            rs.Open "SELECT DISTINCT PEDIDO FROM XXVIA_tB_ORDENES_GRAFICA WHERE pedido in (" + var_cadena_pedidos + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux5.Open "SELECT INTE_EMB_EMBARQUE, SUM(FLOA_SAL_cANTIDAD_LEIDA) FROM XXVIA_tB_SALIDAS WHERE source_header_number = " + CStr(rs!PEDIDO) + " GROUP BY INTE_EMB_EMBARQUE", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux5.EOF Then
                     var_cantidad = IIf(IsNull(rsaux5(0).Value), 0, rsaux5(0).Value)
                  Else
                      var_cantidad = 0
                  End If
                  If var_cantidad > 0 Then
                     rsaux7.Open "UPDATE XXVIA_tB_ORDENES_GRAFICA SET CANTIDAD_LEIDA = " + CStr(IIf(IsNull(rsaux5(1).Value), 0, rsaux5(1).Value)) + ", EMBARQUE = " + CStr(IIf(IsNull(rsaux5(0).Value), 0, rsaux5(0).Value)) + " WHERE PEDIDO = " + CStr(rs!PEDIDO), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     rsaux5.Close
                  Else
                     rsaux5.Close
                     rsaux5.Open "SELECT INTE_EMB_EMBARQUE, SUM(FLOA_SAL_cANTIDAD_LEIDA) FROM XXVIA_tB_SALIDAS_cAJAS WHERE source_header_number = " + CStr(rs!PEDIDO) + " GROUP BY INTE_EMB_EMBARQUE", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux5.EOF Then
                        var_cantidad = IIf(IsNull(rsaux5(0).Value), 0, rsaux5(0).Value)
                     Else
                        var_cantidad = 0
                     End If
                     If var_cantidad > 0 Then
                        rsaux6.Open "UPDATE XXVIA_tB_ORDENES_GRAFICA SET CANTIDAD_LEIDA = " + CStr(IIf(IsNull(rsaux5(1).Value), 0, rsaux5(1).Value)) + ", EMBARQUE = " + CStr(IIf(IsNull(rsaux5(0).Value), 0, rsaux5(0).Value)) + " WHERE PEDIDO = " + CStr(rs!PEDIDO), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux5.Close
                  End If
                  rs.MoveNext
                  Me.Refresh
            Wend
            rs.Close
         Else
            rsaux5.Open "SELECT INTE_EMB_EMBARQUE as embarque, SOURCE_HEADER_NUMBER as pedido, SUM(FLOA_SAL_cANTIDAD_LEIDA) as cantidad FROM XXVIA_tB_SALIDAS WHERE source_header_number in (" + var_cadena_pedidos + ") GROUP BY INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux5.EOF
                  If Not rsaux5.EOF Then
                     var_cantidad = IIf(IsNull(rsaux5!Cantidad), 0, rsaux5!Cantidad)
                  Else
                     var_cantidad = 0
                  End If
                  If var_cantidad > 0 Then
                     rsaux6.Open "UPDATE XXVIA_tB_ORDENES_GRAFICA SET CANTIDAD_LEIDA = " + CStr(IIf(IsNull(rsaux5!Cantidad), 0, rsaux5!Cantidad)) + ", EMBARQUE = " + CStr(IIf(IsNull(rsaux5!Embarque), 0, rsaux5!Embarque)) + " WHERE PEDIDO = " + CStr(rsaux5!PEDIDO), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux5.MoveNext
                  Me.Refresh
            Wend
            rsaux5.Close
            rsaux5.Open "SELECT INTE_EMB_EMBARQUE as embarque, SOURCE_HEADER_NUMBER as pedido, SUM(FLOA_SAL_cANTIDAD_LEIDA) as cantidad FROM XXVIA_tB_SALIDAS_cajas WHERE source_header_number in (" + var_cadena_pedidos + ") GROUP BY INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux5.EOF
                  If Not rsaux5.EOF Then
                     var_cantidad = IIf(IsNull(rsaux5!Cantidad), 0, rsaux5!Cantidad)
                  Else
                     var_cantidad = 0
                  End If
                  If var_cantidad > 0 Then
                     rsaux6.Open "UPDATE XXVIA_tB_ORDENES_GRAFICA SET CANTIDAD_LEIDA = " + CStr(IIf(IsNull(rsaux5!Cantidad), 0, rsaux5!Cantidad)) + ", EMBARQUE = " + CStr(IIf(IsNull(rsaux5!Embarque), 0, rsaux5!Embarque)) + " WHERE PEDIDO = " + CStr(rsaux5!PEDIDO), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux5.MoveNext
                  Me.Refresh
            Wend
            rsaux5.Close
         
         End If
         
         rs.Open "SELECT DISTINCT EMBARQUE FROM XXVIA_tB_ORDENES_GRAFICA WHERE pedido in (" + var_cadena_pedidos + ") AND EMBARQUE IS NOT NULL", cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux5.Open "SELECT CHAR_EMB_ESTATUS FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + CStr(rs!Embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
               rsaux6.Open "UPDATE XXVIA_tB_ORDENES_GRAFICA SET ESTATUS = '" + CStr(IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)) + "' WHERE EMBARQUE = " + CStr(rs!Embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
               rsaux5.Close
               rs.MoveNext
               Me.Refresh
         Wend
         rs.Close
         '----
         Me.lv_grafica.ListItems.Clear
         'rs.Open "select * FROM XXVIA_tB_ORDENES_GRAFICA WHERE FECHA >= TO_DATE('" + var_fecha_inicio + "','YYYY-MM-DD HH24:MI:SS') AND FECHA < TO_DATE('" + var_fecha_fin + "','YYYY-MM-DD')", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If var_filtrado = 6 Then
            rs.Open "select * FROM XXVIA_tB_ORDENES_GRAFICA WHERE pedido in (" + var_cadena_pedidos + ") ", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  VAR_ESTATUS = IIf(IsNull(rs!estatus), "", rs!estatus)
                  If VAR_ESTATUS = "" Then
                     Set list_item = Me.lv_grafica.ListItems.Add(, , rs!PEDIDO)
                     list_item.SubItems(1) = Format(CDate(IIf(IsNull(rs!Fecha), "", rs!Fecha)), "short date")
                     list_item.SubItems(2) = IIf(IsNull(rs!nombre_proveedor), "", rs!nombre_proveedor)
                     list_item.SubItems(3) = IIf(IsNull(rs!Cliente), "", rs!Cliente)
                     list_item.SubItems(4) = IIf(IsNull(rs!Cantidad), "0", rs!Cantidad)
                     list_item.SubItems(5) = IIf(IsNull(rs!cantidad_leida), "0", rs!cantidad_leida)
                     If IIf(IsNull(rs!Cantidad), 0, rs!Cantidad) = 0 Then
                        var_porcentaje = 0
                     Else
                        var_porcentaje = (IIf(IsNull(rs!cantidad_leida), 0, rs!cantidad_leida) * 100) / IIf(IsNull(rs!Cantidad), 1, rs!Cantidad)
                     End If
                     list_item.SubItems(6) = Format(var_porcentaje, "##0.00")
                     list_item.SubItems(7) = IIf(IsNull(rs!tipo_pedido), "", rs!tipo_pedido)
                     list_item.SubItems(8) = IIf(IsNull(rs!Embarque), "", rs!Embarque)
                     list_item.SubItems(9) = IIf(IsNull(rs!estatus), "", rs!estatus)
                  End If
                  rs.MoveNext
                  Me.Refresh
            Wend
            rs.Close
         Else
            rs.Open "select * FROM XXVIA_tB_ORDENES_GRAFICA WHERE pedido in (" + var_cadena_pedidos + ") ", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = Me.lv_grafica.ListItems.Add(, , rs!PEDIDO)
                  list_item.SubItems(1) = Format(CDate(IIf(IsNull(rs!Fecha), "", rs!Fecha)), "short date")
                  list_item.SubItems(2) = IIf(IsNull(rs!nombre_proveedor), "", rs!nombre_proveedor)
                  list_item.SubItems(3) = IIf(IsNull(rs!Cliente), "", rs!Cliente)
                  list_item.SubItems(4) = IIf(IsNull(rs!Cantidad), "0", rs!Cantidad)
                  list_item.SubItems(5) = IIf(IsNull(rs!cantidad_leida), "0", rs!cantidad_leida)
                  If IIf(IsNull(rs!Cantidad), 0, rs!Cantidad) = 0 Then
                     var_porcentaje = 0
                  Else
                     var_porcentaje = (IIf(IsNull(rs!cantidad_leida), 0, rs!cantidad_leida) * 100) / IIf(IsNull(rs!Cantidad), 1, rs!Cantidad)
                  End If
                  list_item.SubItems(6) = Format(var_porcentaje, "##0.00")
                  list_item.SubItems(7) = IIf(IsNull(rs!tipo_pedido), "", rs!tipo_pedido)
                  list_item.SubItems(8) = IIf(IsNull(rs!Embarque), "", rs!Embarque)
                  list_item.SubItems(9) = IIf(IsNull(rs!estatus), "", rs!estatus)
                  rs.MoveNext
                  Me.Refresh
            Wend
            rs.Close
         End If
         
         
         For var_i = 1 To lv_grafica.ListItems.Count
             lv_grafica.ListItems(var_i).Selected = True
             If (lv_grafica.selectedItem.SubItems(6) * 1) > 25 Then
                lv_grafica.ListItems.Item(var_i).ForeColor = vbBlue
                'lv_grafica.ListItems(var_i).ListSubItems(1).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(1).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(2).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(2).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(3).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(3).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(4).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(4).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(5).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(5).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(6).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(6).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(7).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(7).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(8).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(8).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(9).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(9).ForeColor = vbBlue
                lv_grafica.selectedItem.Bold = True
             End If
             If (lv_grafica.selectedItem.SubItems(6) * 1) > 50 Then
                lv_grafica.ListItems.Item(var_i).ForeColor = &HC000C0
                'lv_grafica.ListItems(var_i).ListSubItems(1).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(1).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(2).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(2).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(3).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(3).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(4).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(4).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(5).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(5).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(6).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(6).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(7).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(7).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(8).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(8).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(9).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(9).ForeColor = &HC000C0
                lv_grafica.selectedItem.Bold = True
             End If
             If (lv_grafica.selectedItem.SubItems(6) * 1) = 100 Then
                lv_grafica.ListItems.Item(var_i).ForeColor = vbRed
                'lv_grafica.ListItems(var_i).ListSubItems(1).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(1).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(2).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(2).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(3).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(3).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(4).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(4).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(5).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(5).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(6).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(6).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(7).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(7).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(8).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(8).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(9).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(9).ForeColor = vbRed
                lv_grafica.selectedItem.Bold = True
             End If
         Next var_i
         
            
      End If
   End If
   Me.Enabled = False
   frmoracle_autoriza_cerrar_pantalla.Show
   Me.frm_mensaje.Visible = False
   If var_contraseña_cerrar_pantalla <> "" Then
      Unload Me
   End If
   Me.Enabled = True
End Sub

Private Sub Command1_Click()
   Dim var_filtrado As Integer
   
   Dim var_consecutivo As Integer
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_cadena_agentes = ""
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         var_filtrado = MsgBox("¿Desea el reporte filtrado", vbYesNo, "ATENCION")
         
         var_fecha_inicio = CStr(CDate(Me.txt_inicio) - 1)
         var_fecha_inicio = var_fecha_inicio + " 19:59:59"
         var_cadena = "SELECT DISTINCT a.SOURCE_HEADER_NUMBER AS PEDIDO from WSH_DLVB_DLVY_V B, hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, XXVIA_VW_AGENTES D Where a.delivery_id = B.delivery_id AND A.delivery_detail_id = B.delivery_detail_id AND HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID and B.creation_date >= to_date('" + var_fecha_inicio + "','DD/MM/YYYY HH24:MI:SS') AND TRUNC(B.creation_date) < TO_DATE('" + CStr(CDate(Me.txt_fin) + 1) + "','DD-MM-YYYY') AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID "
         var_cadena = var_cadena + " AND HCAS.CUST_ACCOUNT_ID = D.CUST_ACCOUNT_ID AND OHA.SHIP_FROM_ORG_ID = " + var_unidad_organizacional + " ORDER BY a.SOURCE_HEADER_NUMBER"
         rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_cadena_pedidos = ""
         var_i = 0
         While Not rsaux2.EOF
               If Trim(var_cadena_pedidos) = "" Then
                  var_cadena_pedidos = CStr(rsaux2!PEDIDO)
               Else
                  var_cadena_pedidos = var_cadena_pedidos + "," + CStr(rsaux2!PEDIDO)
               End If
               var_i = var_i + 1
               rsaux2.MoveNext
         Wend
         rsaux2.Close
         Text1 = var_cadena_pedidos
         var_cadena = "SELECT HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID, HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, to_number(A.source_header_number) source_header_number, A.organization_id, A.subinventory, A.released_status, a.source_header_type_name, oha.source_document_id, SUM(A.requested_quantity) + (select nvl(sum(cantidad),0) from xxvia_tb_negado_distribucion where source_header_number = a.source_header_number and cantidad > 0) AS cantidad, MAX(A.LAST_UPDATE_DATE)   As Fecha FROM hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID=HL.LOCATION_ID AND HCSU.SITE_USE_ID = OHA.INVOICE_TO_ORG_ID "
         var_cadena = var_cadena + " AND to_number(source_header_number) IN (" + var_cadena_pedidos + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status in ('Y','C') GROUP BY source_header_number, HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID, HPS.LOCATION_ID, HL.ADDRESS1, A.organization_id, A.subinventory, A.released_status, a.source_header_type_name, oha.source_document_id"

         'var_cadena = "SELECT HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.source_header_number,A.organization_id,A.subinventory,A.released_status, a.source_header_type_name, oha.source_document_id, sum(A.requested_quantity) as cantidad, max(a.LAST_UPDATE_DATE) as fecha from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C "
         'var_cadena = var_cadena + " Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) IN (" + var_cadena_pedidoS + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID  AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status IN ('Y')"
         'var_cadena = var_cadena + " group by HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1, A.source_header_number,A.organization_id,A.subinventory,A.released_status, a.source_header_type_name, oha.source_document_id"

         rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            cnn.BeginTrans
            rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_tEMP_ORACLE_ORDEN_SURTIDO_grafica", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rsaux.Close
            rsaux1.Open "insert into tb_Temp_oracle_orden_surtido_grafica (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
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
                  var_dia = CStr(Day(CDate(rs!Fecha)))
                  var_mes = CStr(Month(CDate(rs!Fecha)))
                  var_año = CStr(Year(CDate(rs!Fecha)))
                  var_hora = CStr(Hour(CDate(rs!Fecha)))
                  var_minuto = CStr(Minute(CDate(rs!Fecha)))
                  var_segundo = CStr(Second(CDate(rs!Fecha)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  If Len(Trim(var_hora)) = 1 Then
                     var_hora = "0" + var_hora
                  End If
                  If Len(Trim(var_minuto)) = 1 Then
                     var_minuto = "0" + var_minuto
                  End If
                  If Len(Trim(var_segundo)) = 1 Then
                     var_segundo = "0" + var_segundo
                  End If
                  
                  var_fecha = "{ts '" + var_año + "-" + var_mes + "-" + var_dia + " " + var_hora + ":" + var_minuto + ":" + var_segundo + ".000'}"
                  
                  
                  
                  
                  var_cadena = "insert into tb_temp_oracle_orden_surtido_grafica (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, ORGANIZATION_ID, SUBINVENTORY, src_requested_quantity, RELEASED_STATUS, CUSTOMER_NAME, COLLECTOR_ID, NAME, date_requested, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, SITE_USE_ID)  values "
                  
                  var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", '" + CStr(IIf(IsNull(rs!source_header_number), "", rs!source_header_number)) + "', " + CStr(IIf(IsNull(rs!organization_id), 0, rs!organization_id)) + ", '" + IIf(IsNull(rs!subinventory), "", rs!subinventory) + "', " + CStr(rs!Cantidad) + ",'" + IIf(IsNull(rs!released_status), "", rs!released_status) + "', '" + Replace(IIf(IsNull(rs!customer_name), "", rs!customer_name), "'", " ") + "' "
                  var_cadena = var_cadena + ", 0,''," + var_fecha + ",'" + CStr(VAR_ESTABLECIMIENTO) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "'," + CStr(rs!CUST_ACCOUNT_ID) + ",'" + rs!source_header_type_name + "','" + CStr(IIf(IsNull(rs!source_document_id), "", rs!source_document_id)) + "'," + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)) + ")"
                  'MsgBox var_cadena
                  rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rsaux1.Open "delete from tb_temp_oracle_orden_surtido_grafica where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number is null", cnn, adOpenDynamic, adLockOptimistic
            If rsaux1.State = 1 Then
               rsaux1.Close
            End If
            rsaux1.Open "select DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_temp_oracle_orden_surtido_grafica where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " group by DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME"
            While Not rsaux1.EOF
                  If rsaux1!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rsaux1!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
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
                     rsaux6.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux6.EOF Then
                        VAR_PROVEEDOR = rsaux6!collector_id
                        VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux6!Name), "", rsaux6!Name)
                     Else
                        VAR_PROVEEDOR = 0
                        VAR_NOMBRE_PROVEEDOR = ""
                     End If
                     rsaux6.Close
                  End If
                  rsaux6.Open "update tb_temp_oracle_orden_surtido_grafica set COLLECTOR_ID = " + CStr(VAR_PROVEEDOR) + ", NAME = '" + VAR_NOMBRE_PROVEEDOR + "' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and source_header_number = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                  rsaux6.Open "SELECT * FROM XXVIA_TB_ORDENES_GRAFICA WHERE ORGANIZACION = " + var_unidad_organizacional + " AND PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If rsaux6.EOF Then
                     var_tipo_pedido = ""
                     If Trim(rsaux1!source_header_type_name) = "VIA_EXPORTACION" Then
                        var_tipo_pedido = "E"
                     End If
                     If Trim(rsaux1!source_header_type_name) = "VIA_PEDIDO_INTERNO" Then
                        var_tipo_pedido = "T"
                     End If
                     If Trim(rsaux1!source_header_type_name) = "VIA_MAYOREO_NACIONAL" Then
                        var_tipo_pedido = "M"
                     End If
                     If Trim(rsaux1!source_header_type_name) = "VIA_VTAS_X_TELEFONO" Then
                        var_tipo_pedido = "VXT"
                     End If

                     var_dia = CStr(Day(CDate(rsaux1!DATE_REQUESTED)))
                     var_mes = CStr(Month(CDate(rsaux1!DATE_REQUESTED)))
                     var_año = CStr(Year(CDate(rsaux1!DATE_REQUESTED)))
                     var_hora = CStr(Hour(CDate(rsaux1!DATE_REQUESTED)))
                     var_minuto = CStr(Minute(CDate(rsaux1!DATE_REQUESTED)))
                     var_segundo = CStr(Second(CDate(rsaux1!DATE_REQUESTED)))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     If Len(Trim(var_hora)) = 1 Then
                        var_hora = "0" + var_hora
                     End If
                     If Len(Trim(var_minuto)) = 1 Then
                        var_minuto = "0" + var_minuto
                     End If
                     If Len(Trim(var_segundo)) = 1 Then
                        var_segundo = "0" + var_segundo
                     End If

                     rsaux5.Open "INSERT INTO XXVIA_TB_ORDENES_GRAFICA (ORGANIZACION, PEDIDO, PROVEEDOR_ID, NOMBRE_PROVEEDOR, CLIENTE, ESTABLECIMIENTO, CANTIDAD, FECHA, TIPO_PEDIDO, CANTIDAD_LEIDA) VALUES (" + var_unidad_organizacional + ", " + CStr(rsaux1!source_header_number) + "," + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "'," + CStr(rsaux1!Cantidad) + ", TO_DATE('" + var_año + "-" + var_mes + "-" + var_dia + " " + var_hora + ":" + var_minuto + ":" + var_segundo + "','YYYY-MM-DD HH24:MI:SS'" + "),'" + var_tipo_pedido + "',0)", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  Else
                     var_tipo_pedido = ""
                     If Trim(rsaux1!source_header_type_name) = "VIA_EXPORTACION" Then
                        var_tipo_pedido = "E"
                     End If
                     If Trim(rsaux1!source_header_type_name) = "VIA_PEDIDO_INTERNO" Then
                        var_tipo_pedido = "T"
                     End If
                     If Trim(rsaux1!source_header_type_name) = "VIA_MAYOREO_NACIONAL" Then
                        var_tipo_pedido = "M"
                     End If
                     If Trim(rsaux1!source_header_type_name) = "VIA_VTAS_X_TELEFONO" Then
                        var_tipo_pedido = "VXT"
                     End If
                     rsaux5.Open "update XXVIA_TB_ORDENES_GRAFICA set tipo_pedido = '" + var_tipo_pedido + "', cantidad = " + CStr(rsaux1!Cantidad) + " where pedido = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux6.Close
                  rsaux1.MoveNext
            Wend
            rsaux1.Close
                  
            var_dia = CStr(Day(CDate(Me.txt_inicio) - 1))
            var_mes = CStr(Month(CDate(Me.txt_inicio) - 1))
            var_año = CStr(Year(CDate(Me.txt_inicio) - 1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = var_año + "-" + var_mes + "-" + var_dia + " 19:59:59"
            
            
            var_dia = CStr(Day(CDate(Me.txt_fin) + 1))
            var_mes = CStr(Month(CDate(Me.txt_fin) + 1))
            var_año = CStr(Year(CDate(Me.txt_fin) + 1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = var_año + "-" + var_mes + "-" + var_dia
         End If
         rs.Close
         x = 1
         If x = 0 Then
            rs.Open "SELECT DISTINCT PEDIDO FROM XXVIA_tB_ORDENES_GRAFICA WHERE pedido in (" + var_cadena_pedidos + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux5.Open "SELECT INTE_EMB_EMBARQUE, SUM(FLOA_SAL_cANTIDAD_LEIDA) FROM XXVIA_tB_SALIDAS WHERE source_header_number = " + CStr(rs!PEDIDO) + " GROUP BY INTE_EMB_EMBARQUE", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux5.EOF Then
                     var_cantidad = IIf(IsNull(rsaux5(0).Value), 0, rsaux5(0).Value)
                  Else
                      var_cantidad = 0
                  End If
                  If var_cantidad > 0 Then
                     rsaux7.Open "UPDATE XXVIA_tB_ORDENES_GRAFICA SET CANTIDAD_LEIDA = " + CStr(IIf(IsNull(rsaux5(1).Value), 0, rsaux5(1).Value)) + ", EMBARQUE = " + CStr(IIf(IsNull(rsaux5(0).Value), 0, rsaux5(0).Value)) + " WHERE PEDIDO = " + CStr(rs!PEDIDO), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     rsaux5.Close
                  Else
                     rsaux5.Close
                     rsaux5.Open "SELECT INTE_EMB_EMBARQUE, SUM(FLOA_SAL_cANTIDAD_LEIDA) FROM XXVIA_tB_SALIDAS_cAJAS WHERE source_header_number = " + CStr(rs!PEDIDO) + " GROUP BY INTE_EMB_EMBARQUE", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux5.EOF Then
                        var_cantidad = IIf(IsNull(rsaux5(0).Value), 0, rsaux5(0).Value)
                     Else
                        var_cantidad = 0
                     End If
                     If var_cantidad > 0 Then
                        rsaux6.Open "UPDATE XXVIA_tB_ORDENES_GRAFICA SET CANTIDAD_LEIDA = " + CStr(IIf(IsNull(rsaux5(1).Value), 0, rsaux5(1).Value)) + ", EMBARQUE = " + CStr(IIf(IsNull(rsaux5(0).Value), 0, rsaux5(0).Value)) + " WHERE PEDIDO = " + CStr(rs!PEDIDO), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux5.Close
                  End If
                  rs.MoveNext
            Wend
            rs.Close
         Else
            rsaux5.Open "SELECT INTE_EMB_EMBARQUE as embarque, SOURCE_HEADER_NUMBER as pedido, SUM(FLOA_SAL_cANTIDAD_LEIDA) as cantidad FROM XXVIA_tB_SALIDAS WHERE source_header_number in (" + var_cadena_pedidos + ") GROUP BY INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux5.EOF
                  If Not rsaux5.EOF Then
                     var_cantidad = IIf(IsNull(rsaux5!Cantidad), 0, rsaux5!Cantidad)
                  Else
                     var_cantidad = 0
                  End If
                  If var_cantidad > 0 Then
                     rsaux6.Open "UPDATE XXVIA_tB_ORDENES_GRAFICA SET CANTIDAD_LEIDA = " + CStr(IIf(IsNull(rsaux5!Cantidad), 0, rsaux5!Cantidad)) + ", EMBARQUE = " + CStr(IIf(IsNull(rsaux5!Embarque), 0, rsaux5!Embarque)) + " WHERE PEDIDO = " + CStr(rsaux5!PEDIDO), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux5.MoveNext
            Wend
            rsaux5.Close
            rsaux5.Open "SELECT INTE_EMB_EMBARQUE as embarque, SOURCE_HEADER_NUMBER as pedido, SUM(FLOA_SAL_cANTIDAD_LEIDA) as cantidad FROM XXVIA_tB_SALIDAS_cajas WHERE source_header_number in (" + var_cadena_pedidos + ") GROUP BY INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux5.EOF
                  If Not rsaux5.EOF Then
                     var_cantidad = IIf(IsNull(rsaux5!Cantidad), 0, rsaux5!Cantidad)
                  Else
                     var_cantidad = 0
                  End If
                  If var_cantidad > 0 Then
                     rsaux6.Open "UPDATE XXVIA_tB_ORDENES_GRAFICA SET CANTIDAD_LEIDA = " + CStr(IIf(IsNull(rsaux5!Cantidad), 0, rsaux5!Cantidad)) + ", EMBARQUE = " + CStr(IIf(IsNull(rsaux5!Embarque), 0, rsaux5!Embarque)) + " WHERE PEDIDO = " + CStr(rsaux5!PEDIDO), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux5.MoveNext
            Wend
            rsaux5.Close
         
         End If
         
         rs.Open "SELECT DISTINCT EMBARQUE FROM XXVIA_tB_ORDENES_GRAFICA WHERE pedido in (" + var_cadena_pedidos + ") AND EMBARQUE IS NOT NULL", cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux5.Open "SELECT CHAR_EMB_ESTATUS FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + CStr(rs!Embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
               rsaux6.Open "UPDATE XXVIA_tB_ORDENES_GRAFICA SET ESTATUS = '" + CStr(IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)) + "' WHERE EMBARQUE = " + CStr(rs!Embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
               rsaux5.Close
               rs.MoveNext
         Wend
         rs.Close
         '----
         Me.lv_grafica.ListItems.Clear
         'rs.Open "select * FROM XXVIA_tB_ORDENES_GRAFICA WHERE FECHA >= TO_DATE('" + var_fecha_inicio + "','YYYY-MM-DD HH24:MI:SS') AND FECHA < TO_DATE('" + var_fecha_fin + "','YYYY-MM-DD')", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If var_filtrado = 6 Then
            rs.Open "select * FROM XXVIA_tB_ORDENES_GRAFICA WHERE pedido in (" + var_cadena_pedidos + ") ", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  VAR_ESTATUS = IIf(IsNull(rs!estatus), "", rs!estatus)
                  If VAR_ESTATUS = "" Then
                     Set list_item = Me.lv_grafica.ListItems.Add(, , rs!PEDIDO)
                     list_item.SubItems(1) = Format(CDate(IIf(IsNull(rs!Fecha), "", rs!Fecha)), "short date")
                     list_item.SubItems(2) = IIf(IsNull(rs!nombre_proveedor), "", rs!nombre_proveedor)
                     list_item.SubItems(3) = IIf(IsNull(rs!Cliente), "", rs!Cliente)
                     list_item.SubItems(4) = IIf(IsNull(rs!Cantidad), "0", rs!Cantidad)
                     list_item.SubItems(5) = IIf(IsNull(rs!cantidad_leida), "0", rs!cantidad_leida)
                     If IIf(IsNull(rs!Cantidad), 0, rs!Cantidad) = 0 Then
                        var_porcentaje = 0
                     Else
                        var_porcentaje = (IIf(IsNull(rs!cantidad_leida), 0, rs!cantidad_leida) * 100) / IIf(IsNull(rs!Cantidad), 1, rs!Cantidad)
                     End If
                     list_item.SubItems(6) = Format(var_porcentaje, "##0.00")
                     list_item.SubItems(7) = IIf(IsNull(rs!tipo_pedido), "", rs!tipo_pedido)
                     list_item.SubItems(8) = IIf(IsNull(rs!Embarque), "", rs!Embarque)
                     list_item.SubItems(9) = IIf(IsNull(rs!estatus), "", rs!estatus)
                  End If
                  rs.MoveNext
            Wend
            rs.Close
         Else
            rs.Open "select * FROM XXVIA_tB_ORDENES_GRAFICA WHERE pedido in (" + var_cadena_pedidos + ") ", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = Me.lv_grafica.ListItems.Add(, , rs!PEDIDO)
                  list_item.SubItems(1) = Format(CDate(IIf(IsNull(rs!Fecha), "", rs!Fecha)), "short date")
                  list_item.SubItems(2) = IIf(IsNull(rs!nombre_proveedor), "", rs!nombre_proveedor)
                  list_item.SubItems(3) = IIf(IsNull(rs!Cliente), "", rs!Cliente)
                  list_item.SubItems(4) = IIf(IsNull(rs!Cantidad), "0", rs!Cantidad)
                  list_item.SubItems(5) = IIf(IsNull(rs!cantidad_leida), "0", rs!cantidad_leida)
                  If IIf(IsNull(rs!Cantidad), 0, rs!Cantidad) = 0 Then
                     var_porcentaje = 0
                  Else
                     var_porcentaje = (IIf(IsNull(rs!cantidad_leida), 0, rs!cantidad_leida) * 100) / IIf(IsNull(rs!Cantidad), 1, rs!Cantidad)
                  End If
                  list_item.SubItems(6) = Format(var_porcentaje, "##0.00")
                  list_item.SubItems(7) = IIf(IsNull(rs!tipo_pedido), "", rs!tipo_pedido)
                  list_item.SubItems(8) = IIf(IsNull(rs!Embarque), "", rs!Embarque)
                  list_item.SubItems(9) = IIf(IsNull(rs!estatus), "", rs!estatus)
                  rs.MoveNext
            Wend
            rs.Close
         End If
         
         
         For var_i = 1 To lv_grafica.ListItems.Count
             lv_grafica.ListItems(var_i).Selected = True
             If (lv_grafica.selectedItem.SubItems(6) * 1) > 25 Then
                lv_grafica.ListItems.Item(var_i).ForeColor = vbBlue
                'lv_grafica.ListItems(var_i).ListSubItems(1).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(1).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(2).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(2).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(3).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(3).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(4).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(4).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(5).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(5).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(6).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(6).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(7).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(7).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(8).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(8).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(9).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(9).ForeColor = vbBlue
                lv_grafica.selectedItem.Bold = True
             End If
             If (lv_grafica.selectedItem.SubItems(6) * 1) > 50 Then
                lv_grafica.ListItems.Item(var_i).ForeColor = &HC000C0
                'lv_grafica.ListItems(var_i).ListSubItems(1).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(1).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(2).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(2).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(3).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(3).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(4).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(4).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(5).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(5).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(6).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(6).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(7).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(7).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(8).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(8).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(9).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(9).ForeColor = &HC000C0
                lv_grafica.selectedItem.Bold = True
             End If
             If (lv_grafica.selectedItem.SubItems(6) * 1) = 100 Then
                lv_grafica.ListItems.Item(var_i).ForeColor = vbRed
                'lv_grafica.ListItems(var_i).ListSubItems(1).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(1).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(2).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(2).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(3).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(3).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(4).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(4).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(5).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(5).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(6).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(6).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(7).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(7).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(8).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(8).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(9).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(9).ForeColor = vbRed
                lv_grafica.selectedItem.Bold = True
             End If
         Next var_i
         
            
      End If
   End If
End Sub

Private Sub Command11_Click()
   If IsDate(Me.txt_inicio) Then
      Me.mes.Value = CDate(Me.txt_inicio)
   Else
      mes.Value = Date
   End If
   var_tipo_mes = 1
   mes.Visible = True
   mes.SetFocus
End Sub

Private Sub Command12_Click()
   If IsDate(Me.txt_fin) Then
      mes.Value = CDate(Me.txt_fin)
   Else
      mes.Value = Date
   End If
   var_tipo_mes = 2
   mes.Visible = True
   mes.SetFocus
End Sub

Private Sub Form_Load()
   Me.mes.Visible = False
   Me.txt_fin = Date
   Me.txt_inicio = Date
   Me.frm_mensaje.Visible = False
   Top = 0
   Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_grafica_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_grafica, ColumnHeader)
End Sub

Private Sub mes_DateDblClick(ByVal DateDblClicked As Date)
   If var_tipo_mes = 1 Then
      txt_inicio = mes.Value
      
   End If
   If var_tipo_mes = 2 Then
      txt_fin = mes.Value
   End If
   mes.Visible = False
End Sub

Private Sub mes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      mes.Visible = False
   End If
End Sub
