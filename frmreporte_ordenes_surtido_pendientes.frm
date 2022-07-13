VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_ordenes_surtido_pendientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Surtido Pendientes"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5805
   Begin VB.Frame Frame2 
      Caption         =   " Agentes "
      Height          =   5175
      Left            =   60
      TabIndex        =   3
      Top             =   405
      Width           =   5685
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   435
         Picture         =   "frmreporte_ordenes_surtido_pendientes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   105
         Picture         =   "frmreporte_ordenes_surtido_pendientes.frx":0216
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1095
         Picture         =   "frmreporte_ordenes_surtido_pendientes.frx":0318
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         Picture         =   "frmreporte_ordenes_surtido_pendientes.frx":03EA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Marcar (Enter)"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1425
         Picture         =   "frmreporte_ordenes_surtido_pendientes.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   240
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_agentes 
         Height          =   4365
         Left            =   45
         TabIndex        =   10
         Top             =   720
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   7699
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
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Frame Frame7 
         Height          =   120
         Left            =   30
         TabIndex        =   9
         Top             =   525
         Width           =   5610
      End
   End
   Begin VB.Frame Frame5 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   345
      Width           =   5715
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmreporte_ordenes_surtido_pendientes.frx":084A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5340
      Picture         =   "frmreporte_ordenes_surtido_pendientes.frx":094C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "frmreporte_ordenes_surtido_pendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   Dim var_cadena_agentes As String
   Dim var_cadena_agentes_tiendas As String
   Dim var_n As Integer, var_i As Integer, var_contador_agentes As Integer, var_consecutivo As Integer
   cnn.CommandTimeout = 6000
   var_n = 0
   var_n = lv_agentes.ListItems.Count
   If var_n > 0 Then
      var_contador_agentes = 0
      var_si_tiendas = 0
      For var_i = 1 To var_n
          lv_agentes.ListItems.Item(var_i).Selected = True
          If Trim(lv_agentes.selectedItem.SubItems(2)) = "*" Then
             var_contador_agentes = var_contador_agentes + 1
             If CDbl(Me.lv_agentes.selectedItem) = -3 Then
                var_si_tiendas = 1
             End If
          End If
      Next var_i
      If var_si_tiendas = 1 Then
          
      End If
      If var_contador_agentes > 0 Then
         var_contador_agentes = 0
         var_cadena_agentes_tiendas = ""
         For var_i = 1 To var_n
             lv_agentes.ListItems.Item(var_i).Selected = True
             If Trim(lv_agentes.selectedItem.SubItems(2)) = "*" Then
                If var_contador_agentes = 0 Then
                   var_cadena_agentes = Trim(lv_agentes.selectedItem)
                   var_contador_agentes = 1
                Else
                   var_cadena_agentes = var_cadena_agentes + ", " + Trim(lv_agentes.selectedItem)
                End If
                
             End If
         Next var_i
         var_cadena_agentes = var_cadena_agentes + ")"
         rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         
         var_cadena = "SELECT oha.source_document_id, a.source_header_type_name, OHA.ORDERED_DATE, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, HPS2.LOCATION_ID AS ESTABLECIMIENTO, HL2.ADDRESS1 AS NOMBRE_ESTABLECIMIENTO, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, j.NAME, j.salesrep_id as collector, EXISTENCIA from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, hz_cust_acct_sites_all HCAS2, HZ_PARTY_SITES HPS2, HZ_LOCATIONS HL2, HZ_CUST_SITE_USES_ALL HCSU2, xxvia_system_items_b C, hz_customer_profiles D,  XXVIA_VENDEDORES J, XXVIA_vw_existencias XVE "
         var_cadena = var_cadena + " Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HCAS2.PARTY_SITE_ID = HPS2.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HPS2.LOCATION_ID =HL2.LOCATION_ID "
         var_cadena = var_cadena + " AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU2.SITE_USE_ID= OHA.SHIP_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND HCSU2.CUST_ACCT_SITE_ID = HCAS2.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y' AND A.INVENTORY_ITEM_ID = XVE.INVENTORY_ITEM_ID AND A.ORGANIZATION_ID = XVE.ORGANIZATION_ID AND A.subinventory = XVE.SUBINVENTORY_CODE and oha.salesrep_id = j.salesrep_id and j.salesrep_id in (" + var_cadena_agentes + " and A.organization_id = " + var_unidad_organizacional

         
         Text1 = var_cadena
         
         rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            cnn.BeginTrans
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_ORACLE_ORDENES_PEDNIDENTES_SURTIR", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_consecutivo = IIf(IsNull(rsaux!NUMERO), 0, rsaux!NUMERO)
            Else
               var_consecutivo = 0
            End If
            rsaux.Close
            var_consecutivo = var_consecutivo + 1
            rsaux.Open "insert into TB_TEMP_ORACLE_ORDENES_PEDNIDENTES_SURTIR (INTE_TEM_CONSECUTIVO) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            While Not rs.EOF
                  var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDENES_PEDNIDENTES_SURTIR (INTE_TEM_CONSECUTIVO, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, PEDIDO, CODIGO, DESCRIPCION, UBICACION, EXISTENCIA, SURTIR, SURTIDA, EMPACADO, FECHA, TIPO_PEDIDO, DOCUMENTO_ID, ENTREGA) "
                  var_cadena = var_cadena + "values  (" + CStr(var_consecutivo) + ", '" + CStr(rs!collector) + "', '" + rs!Name + "', '" + CStr(rs!LOCATION_ID) + "', '" + rs!CUSTOMER_NAME + "', '" + CStr(rs!ESTABLECIMIENTO) + "', '" + CStr(rs!NOMBRE_ESTABLECIMIENTO) + "', " + CStr(rs!source_header_number) + ",'" + rs!segment1 + "','" + rs!Description + "',''," + CStr(rs!EXISTENCIA) + "," + CStr(rs!requested_quantity) + ",0,0,'" + CStr(rs!ORDERED_DATE) + "','" + IIf(IsNull(rs!source_header_type_name), "", rs!source_header_type_name) + "'," + CStr(IIf(IsNull(rs!source_document_id), 0, rs!source_document_id)) + ", " + CStr(IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)) + ")"
                  rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rsaux.Open "select distinct tipo_pedido, documento_id, pedido from TB_TEMP_ORACLE_ORDENES_PEDNIDENTES_SURTIR where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and TIPO_PEDIDO = 'VIA_PEDIDO_INTERNO'", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux!documento_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     rsaux3.Open "UPDATE TB_TEMP_ORACLE_ORDENES_PEDNIDENTES_SURTIR SET agente = '" + IIf(IsNull(rsaux2!attribute1), "", rsaux2!attribute1) + "', nombre_agente = '" + IIf(IsNull(rsaux2!Description), "", rsaux2!Description) + "', NOMBRE_CLIENTE = '" + IIf(IsNull(rsaux2!Description), "", rsaux2!Description) + "' WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux!pedido), cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux2.Close

                  rsaux.MoveNext
            Wend
            rsaux.Close
            If var_si_tiendas = 1 Then
               var_consecutivo_tiendas = var_consecutivo
               frmoracle_seleccion_tiendas.Show 1
            End If
            
            rsaux.Open "select distinct PEDIDO from TB_TEMP_ORACLE_ORDENES_PEDNIDENTES_SURTIR where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido is not null", cnn, adOpenDynamic, adLockOptimistic
            VAR_CADENA_PEDIDOS = ""
            While Not rsaux.EOF
                  If VAR_CADENA_PEDIDOS = "" Then
                     VAR_CADENA_PEDIDOS = CStr(rsaux!pedido)
                  Else
                     VAR_CADENA_PEDIDOS = VAR_CADENA_PEDIDOS + ", " + CStr(rsaux!pedido)
                  End If
                  rsaux.MoveNext
            Wend
            rsaux.Close
            Text1 = VAR_CADENA_PEDIDOS
            rsaux.Open "delete from xxvia_tb_temp_Cant_leidas where consecutivo = " + CStr(var_consecutivo), cnnoracle_4, adOpenDynamic, adLockOptimistic
        
            rsaux2.Open "insert into xxvia_tb_Temp_cant_leidas (CONSECUTIVO, PEDIDO, CODIGO, CANTIDAD, TIPO, ENTREGA) select " + CStr(var_consecutivo) + ",source_header_number, segment1, sum(floa_sal_cantidad_leida) as cantidad, 'S', DELIVERY_ID from xxvia_Tb_salidas where source_header_number in (" + VAR_CADENA_PEDIDOS + ") group by " + CStr(var_consecutivo) + ", source_header_number, segment1, 'S', DELIVERY_ID UNION all select " + CStr(var_consecutivo) + ", source_header_number, segment1, sum(floa_sal_cantidad_leida) as cantiada,'E', DELIVERY_ID from xxvia_Tb_salidas_cajas where source_header_number in (" + VAR_CADENA_PEDIDOS + ") group by " + CStr(var_consecutivo) + ",source_header_number, segment1, 'E', DELIVERY_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rsaux.Open "select * from xxvia_tb_Temp_cant_leidas where consecutivo = " + CStr(var_consecutivo), cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                 If rsaux!TIPO = "S" Then
                  rsaux1.Open "update TB_TEMP_ORACLE_ORDENES_PEDNIDENTES_SURTIR set SURTIDA = " + CStr(rsaux!Cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux!pedido) + " AND CODIGO = '" + CStr(rsaux!CODIGO) + "' AND ENTREGA = " + CStr(IIf(IsNull(rsaux!ENTREGA), 0, rsaux!ENTREGA)), cnn, adOpenDynamic, adLockOptimistic
                 Else
                  rsaux1.Open "update TB_TEMP_ORACLE_ORDENES_PEDNIDENTES_SURTIR set EMPACADO = " + CStr(rsaux!Cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux!pedido) + " AND CODIGO = '" + CStr(rsaux!CODIGO) + "' AND ENTREGA = " + CStr(IIf(IsNull(rsaux!ENTREGA), 0, rsaux!ENTREGA)), cnn, adOpenDynamic, adLockOptimistic
                 End If
                  rsaux.MoveNext
            Wend
            rsaux.Close
            'rsaux.Open "SELECT * FROM TB_TEMP_ORACLE_ORDENES_PEDNIDENTES_SURTIR WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND AGENTE IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
            'While Not rsaux.EOF
            '      rsaux1.Open "SELECT * FROM XXVIA_TB_SALIDAS WHERE SOURCE_HEADER_NUMBER = " + CStr(rsaux!pedido) + " AND SEGMENT1 = '" + rsaux!CODIGO + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            '      If Not rsaux1.EOF Then
            '         rsaux2.Open "Update TB_TEMP_ORACLE_ORDENES_PEDNIDENTES_SURTIR SET SURTIDA = " + CStr(rsaux1!FLOA_sAL_cANTIDAD_LEIDA) + " WHERE PEDIDO = " + CStr(rsaux!pedido) + " AND CODIGO = '" + rsaux!CODIGO + "'", cnn, adOpenDynamic, adLockOptimistic
            '      End If
            '      rsaux1.Close
            '      rsaux.MoveNext
            'Wend
            'rsaux.Close
           '
           ' rsaux.Open "SELECT * FROM TB_TEMP_ORACLE_ORDENES_PEDNIDENTES_SURTIR WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND AGENTE IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
           ' While Not rsaux.EOF
           '       rsaux1.Open "SELECT SOURCE_HEADER_NUMBER, SEGMENT1, SUM(FLOA_sAL_CANTIDAD_LEIDA) AS FLOA_sAL_cANTIDAD_LEIDA FROM XXVIA_TB_SALIDAS_CAJAS WHERE SOURCE_HEADER_NUMBER = " + CStr(rsaux!pedido) + " AND SEGMENT1 = '" + rsaux!CODIGO + "' GROUP BY SOURCE_HEADER_NUMBER, SEGMENT1", cnnoracle_4, adOpenDynamic, adLockOptimistic
           '       If Not rsaux1.EOF Then
           '          rsaux2.Open "Update TB_TEMP_ORACLE_ORDENES_PEDNIDENTES_SURTIR SET EMPACADO = " + CStr(rsaux1!FLOA_sAL_cANTIDAD_LEIDA) + " WHERE PEDIDO = " + CStr(rsaux!pedido) + " AND CODIGO = '" + rsaux!CODIGO + "'", cnn, adOpenDynamic, adLockOptimistic
           '      End If
           '       rsaux1.Close
           '       rsaux.MoveNext
           ' Wend
           ' rsaux.Close
            
            rsaux.Open "delete from TB_TEMP_ORACLE_ORDENES_PEDNIDENTES_SURTIR where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and agente is null", cnn, adOpenDynamic, adLockOptimistic
            'rsaux.Open "delete from TB_TEMP_ORACLE_ORDENES_PEDNIDENTES_SURTIR where SURTIR - (SURTIDA+EMPACADO) <= 0", cnn, adOpenDynamic, adLockOptimistic
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_pendientes_concentrado.rpt")
            reporte.RecordSelectionFormula = "{VW_ORACLE_ORDENES_PENDIENTES_SURTIR_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + Str(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Ordenes de surtido pendientes de empacar o facturar"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_pendientes_concentrado.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_ORDENES_PENDIENTES_SURTIR_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + Str(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\Reporte_ordenes_pendientes_concentrado_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
   
   
   
   
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_pendientes.rpt")
            reporte.RecordSelectionFormula = "{VW_ORACLE_ORDENES_PENDIENTES_SURTIR.INTE_TEM_CONSECUTIVO} = " + Str(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Ordenes de surtido pendientes de empacar o facturar"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_pendientes.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_ORDENES_PENDIENTES_SURTIR.INTE_TEM_CONSECUTIVO} = " + Str(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\Reporte_ordenes_pendientes_detalle_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
   
            rsaux.Open "delete from TB_TEMP_ORACLE_ORDENES_PEDNIDENTES_SURTIR where inte_tem_Consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            
            
         Else
            MsgBox "No existen ordenes de surtido pendientes para los agentes seleccionados", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "No se a seleccionado ningun agente", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No existe ningun agente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
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

Private Sub Command2_Click()
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

Private Sub Command3_Click()
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

Private Sub Command4_Click()
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

Private Sub Command5_Click()
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

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 1000
   Left = 3000
   'rs.Open "select * from ar_collectors ", cnnoracle_4, adOpenDynamic, adLockOptimistic
   rs.Open "SELECT salesrep_id collector_id, name  FROM XXVIA_VENDEDORES where name is not null", cnnoracle_4, adOpenDynamic, adLockOptimistic
   numero_items_agentes = 0
   While Not rs.EOF
      Set list_item = lv_agentes.ListItems.Add(, , rs!COLlECTOR_ID)
      list_item.SubItems(1) = IIf(IsNull(rs!Name), "", rs!Name)
      list_item.SubItems(2) = ""
      rs.MoveNext:
      numero_items_agentes = numero_items_agentes + 1
   Wend
   rs.Close
   If numero_items_agentes > 12 Then
      lv_agentes.ColumnHeaders(2).Width = 4200.71
   Else
      lv_agentes.ColumnHeaders(2).Width = 4499.71
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro = False
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_agentes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_agentes, ColumnHeader)
End Sub

Private Sub lv_agentes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
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
End Sub
